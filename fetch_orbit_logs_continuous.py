#!/usr/bin/env python3
"""
Continuous Orbit log collector — GUI edition.

Startup flow:
  1. Setup dialog: Continue from existing file  OR  Start new collection
  2. Main window: shows next retrieval time, countdown, Retrieve Now button,
     and a live activity log.

Threading model (replaces the old subprocess launcher/worker):
  Main thread  → tkinter GUI
  poll-worker  → background thread; fetches and appends rows on schedule
  watchdog     → background thread; restarts poll-worker if it dies
"""
from __future__ import annotations

import queue
import re
import threading
import time
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from bosdyn.orbit.client import Client


# =============================================================================
# Configuration
# =============================================================================
ORBIT_HOST      = "172.24.43.164"
ORBIT_API_TOKEN = "98bbd976-0b14-4cc5-b69d-d68d1f9b88a9"
TLS_VERIFY      = False

RUN_FETCH_LIMIT    = 10000
POLL_INTERVAL_SECS = 30 * 60   # default interval between automatic retrievals
RESTART_DELAY_SECS = 15        # wait before watchdog restarts a dead poll thread
BOOTSTRAP_MINUTES  = 30        # look-back when no existing data and no explicit start_dt

HEADERS      = ["Robot Serial Number", "Date and time", "Route", "Orbit log"]
OUT_TIME_FMT = "%Y/%m/%d %H:%M:%S"
JST          = timezone(timedelta(hours=9))

INCLUDE_KEYWORDS = [""]
IGNORE_WORDS     = []
MISSION_IGNORE   = []

PATTERNS_XLSX  = "log_patterns_example.xlsx"
PATTERNS_SHEET = "Log Patterns"

# Row colour per category — tiered by priority
CATEGORY_COLORS: Dict[str, str] = {
    # Critical
    "Software / System Issues":          "FF4C4C",
    "Navigation — Route Conflict":       "FF8042",
    # High
    "Navigation — Stuck (Goal Blocked)": "FF9900",
    "Navigation — Stuck":                "FFB347",
    # Warning
    "Navigation — Path Not Found":       "FFD966",
    "Entity Detection":                  "FFDC73",
    # Informational
    "Task Completion":                   "C6EFCE",
    "Docking / Undocking":               "BDD7EE",
    "Run Lifecycle":                     "DDEEFF",
    "Task Skipped":                      "F2F2F2",
    "Mission Interruption":              "F0EBFF",
    "Battery":                           "FEFCE8",
}
COLOR_FALLBACK = "F3F3F3"

# Overrides matched against reference PATTERN TEXT (known patterns)
PATTERN_OVERRIDES: List[tuple] = [
    ("after stuck",          "FFD966"),
    ("after route conflict", "FFD966"),
]

# Overrides matched against actual LOG TEXT at runtime (new/unknown entries)
KEYWORD_OVERRIDES: List[tuple] = [
    ("battery temp",    "FFD966"),
    ("battery error",   "FFD966"),
    ("battery fault",   "FFD966"),
    ("battery failure", "FFD966"),
    ("thermal",         "FFD966"),
]
# =============================================================================


# ---------------------------------------------------------------------------
# Regex helpers
# ---------------------------------------------------------------------------
def _build_re(phrases: List[str]) -> re.Pattern:
    if not phrases:
        return re.compile(r"(?!x)x")
    return re.compile("|".join(re.escape(p) for p in phrases), re.IGNORECASE)

INCLUDE_RE        = _build_re(INCLUDE_KEYWORDS)
IGNORE_RE         = _build_re(IGNORE_WORDS)
MISSION_IGNORE_RE = _build_re(MISSION_IGNORE)


# ---------------------------------------------------------------------------
# Log-pattern colour helpers
# ---------------------------------------------------------------------------
def _pattern_to_regex(pattern: str) -> re.Pattern:
    escaped = re.escape(pattern)
    escaped = re.sub(r"'\\\{[a-z_]+\\\}'", r"(?:'[^']*'|\\S+)", escaped)
    escaped = re.sub(r"\\\{[a-z_]+\\\}", r".+?", escaped)
    return re.compile(r"^\s*" + escaped + r"\s*$", re.IGNORECASE)

_category_matchers: List[tuple] = []   # (category, regex, PatternFill)

def load_category_matchers() -> None:
    global _category_matchers
    try:
        wb = load_workbook(PATTERNS_XLSX, read_only=True, data_only=True)
        ws = wb[PATTERNS_SHEET]
        matchers = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            category = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            pattern  = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            if not category or not pattern:
                continue
            override = next((c for kw, c in PATTERN_OVERRIDES if kw in pattern.lower()), None)
            color    = override or CATEGORY_COLORS.get(category, COLOR_FALLBACK)
            matchers.append((category, _pattern_to_regex(pattern), PatternFill("solid", fgColor=color)))
        wb.close()
        _category_matchers = matchers
    except Exception:
        _category_matchers = []

def get_row_fill(log_text: str) -> Optional[PatternFill]:
    for _cat, rx, fill in _category_matchers:
        if rx.match(log_text):
            return fill
    log_lower = log_text.lower()
    for kw, color in KEYWORD_OVERRIDES:
        if kw in log_lower:
            return PatternFill("solid", fgColor=color)
    return PatternFill("solid", fgColor=COLOR_FALLBACK) if _category_matchers else None


# ---------------------------------------------------------------------------
# Orbit helpers
# ---------------------------------------------------------------------------
def parse_iso8601(s: Any) -> Optional[datetime]:
    if not isinstance(s, str) or not s:
        return None
    s = s.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None

def to_jst_str(dt: Optional[datetime]) -> str:
    if dt is None:
        return ""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(JST).strftime(OUT_TIME_FMT)

def normalise_log_json(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, list):
        return [x for x in payload if isinstance(x, dict)]
    if isinstance(payload, dict):
        res = payload.get("resources")
        if isinstance(res, list):
            return [x for x in res if isinstance(x, dict)]
    return []

def mission_allowed(name: str) -> bool:
    if not name:
        return False
    return not MISSION_IGNORE_RE.search(name)

def extract_rows(run: Dict[str, Any], log_items: List[Dict[str, Any]]) -> List[List[Any]]:
    mission = str(run.get("missionName") or "").strip()
    serial  = str(run.get("robotSerial") or "Unknown")
    if not mission_allowed(mission):
        return []
    run_start_dt = parse_iso8601(run.get("startTime"))
    rows: List[List[Any]] = []
    for item in log_items:
        detail = str(item.get("detail") or "")
        if not detail:
            continue
        if INCLUDE_RE.search(detail) and not IGNORE_RE.search(detail):
            item_dt = parse_iso8601(item.get("start")) or run_start_dt
            rows.append([serial, to_jst_str(item_dt), mission, detail])
    return rows


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------
def get_last_timestamp(path: str) -> Optional[datetime]:
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except FileNotFoundError:
        return None
    ws   = wb.active
    last: Optional[datetime] = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        cell = row[1] if len(row) > 1 else None
        if not cell:
            continue
        try:
            dt = datetime.strptime(str(cell).strip(), OUT_TIME_FMT).replace(tzinfo=JST)
            if last is None or dt > last:
                last = dt
        except Exception:
            pass
    wb.close()
    return last

def append_to_xlsx(path: str, rows: List[List[Any]]) -> None:
    try:
        wb = load_workbook(path)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Orbit Logs"
        ws.append(HEADERS)
        ws.freeze_panes = "A2"
    for r in rows:
        ws.append(r)
        fill = get_row_fill(r[3] if len(r) > 3 else "")
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill
    for col_idx in range(1, len(HEADERS) + 1):
        col_letter = get_column_letter(col_idx)
        if ws.max_row > 1:
            max_len = max(len(str(c.value)) if c.value else 0 for c in ws[col_letter])
            ws.column_dimensions[col_letter].width = min(max(14, max_len + 2), 90)
    wb.save(path)


# ---------------------------------------------------------------------------
# Poll cycle
# ---------------------------------------------------------------------------
def poll_once(client: Client, start_dt: Optional[datetime], end_dt: datetime) -> List[List[Any]]:
    """Fetch log rows in [start_dt, end_dt). Returns the rows (does NOT write xlsx)."""
    if start_dt is None:
        start_dt = end_dt - timedelta(minutes=BOOTSTRAP_MINUTES)

    resp = client.get_runs(params={"orderBy": "newest", "limit": RUN_FETCH_LIMIT})
    resp.raise_for_status()
    runs = resp.json().get("resources", []) or []

    target_runs = [
        r for r in runs
        if (st := parse_iso8601(r.get("startTime"))) and start_dt <= st.astimezone(JST) < end_dt
    ]

    rows: List[List[Any]] = []
    for r in target_runs:
        run_id = r.get("uuid") or r.get("id")
        if not run_id:
            continue
        try:
            log_resp = client.get_run_log(run_id)
            log_resp.raise_for_status()
            rows.extend(extract_rows(r, normalise_log_json(log_resp.json())))
        except Exception:
            pass

    rows.sort(key=lambda x: x[1])

    if start_dt:
        cutoff = start_dt.strftime(OUT_TIME_FMT)
        rows   = [r for r in rows if r[1] > cutoff]

    return rows


# =============================================================================
# Setup dialogs
# =============================================================================
class _NewCollectionDialog(tk.Toplevel):
    """Sub-dialog to collect start date/time and save path for a new log file."""

    def __init__(self, parent: tk.Toplevel):
        super().__init__(parent)
        self._parent = parent
        self.title("New Collection  /  新規収集設定")
        self.resizable(False, False)
        self.grab_set()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._centre()

    def _centre(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - self.winfo_width()) // 2}+{(sh - self.winfo_height()) // 2}")

    def _build(self):
        pad = dict(padx=20, pady=8)

        ttk.Label(self, text="New Collection Setup  /  新規収集設定",
                  font=("Segoe UI", 11, "bold")).pack(**pad)

        # Start date/time
        f1 = ttk.LabelFrame(self, text="Start date / time (JST)  /  開始日時 (JST)", padding=8)
        f1.pack(fill="x", padx=20, pady=4)
        ttk.Label(f1, text="Format: YYYY/MM/DD HH:MM:SS",
                  foreground="grey", font=("Segoe UI", 8)).pack(anchor="w")
        self._dt_var = tk.StringVar(value=datetime.now(JST).strftime(OUT_TIME_FMT))
        ttk.Entry(f1, textvariable=self._dt_var, width=24).pack(anchor="w", pady=4)

        # Save location
        f2 = ttk.LabelFrame(self, text="Save location  /  保存先", padding=8)
        f2.pack(fill="x", padx=20, pady=4)
        row = ttk.Frame(f2)
        row.pack(fill="x")
        default = f"orbit_logs_{datetime.now(JST).strftime('%Y%m%d')}.xlsx"
        self._path_var = tk.StringVar(value=default)
        ttk.Entry(row, textvariable=self._path_var, width=32).pack(side="left")
        ttk.Button(row, text="Browse  /  参照", command=self._browse).pack(side="left", padx=6)

        # Buttons
        btn = ttk.Frame(self)
        btn.pack(pady=14)
        ttk.Button(btn, text="Start  /  開始", command=self._confirm).pack(side="left", padx=8)
        ttk.Button(btn, text="Cancel  /  キャンセル", command=self.destroy).pack(side="left")

    def _browse(self):
        path = fd.asksaveasfilename(
            title="Save log file  /  保存先を選択",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=self._path_var.get(),
            parent=self,
        )
        if path:
            self._path_var.set(path)

    def _confirm(self):
        try:
            start_dt = datetime.strptime(
                self._dt_var.get().strip(), OUT_TIME_FMT
            ).replace(tzinfo=JST)
        except ValueError:
            mb.showerror("Invalid date/time  /  日時エラー",
                         "Please use format: YYYY/MM/DD HH:MM:SS\n"
                         "形式: YYYY/MM/DD HH:MM:SS", parent=self)
            return
        path = self._path_var.get().strip()
        if not path:
            mb.showerror("Missing path  /  パスエラー",
                         "Please specify a save location.\n保存先を指定してください。",
                         parent=self)
            return
        self._parent._accept(path, start_dt)
        self.destroy()


class SetupDialog(tk.Toplevel):
    """
    Modal startup dialog.
    Sets self.result to {"xlsx_path": str, "start_dt": datetime|None}
    or leaves it as None if the user cancels.
    """

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.result: Optional[Dict] = None
        self.title("Orbit Log Collector  —  Setup  /  セットアップ")
        self.resizable(False, False)
        self.grab_set()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self._centre()

    def _centre(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - self.winfo_width()) // 2}+{(sh - self.winfo_height()) // 2}")

    def _build(self):
        pad = dict(padx=24, pady=10)

        ttk.Label(self, text="Orbit Log Collector",
                  font=("Segoe UI", 15, "bold")).pack(pady=(20, 4))
        ttk.Label(self, text="How would you like to start?\n起動方法を選択してください",
                  justify="center", foreground="#555555").pack(pady=(0, 16))

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20)

        ttk.Button(self,
                   text="Continue from existing file\n既存ファイルから続ける",
                   width=38, command=self._continue_existing).pack(**pad)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20)

        ttk.Button(self,
                   text="Start new collection\n新規収集を開始する",
                   width=38, command=self._start_new).pack(**pad)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20)

        ttk.Button(self, text="Cancel  /  キャンセル",
                   command=self._cancel).pack(pady=12)

    def _continue_existing(self):
        path = fd.askopenfilename(
            title="Select orbit log file  /  ログファイルを選択",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            parent=self,
        )
        if not path:
            return
        try:
            wb = load_workbook(path, read_only=True)
            if "Orbit Logs" not in wb.sheetnames:
                mb.showerror("Invalid file  /  ファイルエラー",
                             "No 'Orbit Logs' sheet found in the selected file.\n"
                             "選択したファイルに 'Orbit Logs' シートが見つかりません。",
                             parent=self)
                wb.close()
                return
            wb.close()
        except Exception as e:
            mb.showerror("Error  /  エラー", str(e), parent=self)
            return
        self._accept(path, None)

    def _start_new(self):
        _NewCollectionDialog(self)

    def _accept(self, xlsx_path: str, start_dt: Optional[datetime]):
        self.result = {"xlsx_path": xlsx_path, "start_dt": start_dt}
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# =============================================================================
# Main GUI application
# =============================================================================
class CollectorApp:
    """
    Wraps the root Tk window and manages the background polling thread.

    Thread safety:
      - All GUI updates go through root.after(0, fn) or are scheduled
        by _drain_log_queue() / _tick_countdown() which run on the main thread.
      - The poll thread communicates via self._log_queue (thread-safe queue).
      - next_dt is protected by self._next_dt_lock.
    """

    _COUNTDOWN_MS = 1_000
    _LOG_DRAIN_MS =   200

    def __init__(self, root: tk.Tk, config: Dict):
        self._root      = root
        self._xlsx_path = config["xlsx_path"]
        self._start_dt  = config.get("start_dt")   # explicit query-from time (first run only)

        # First poll in a few seconds so the user sees something quickly
        self._next_dt      = datetime.now(JST) + timedelta(seconds=8)
        self._next_dt_lock = threading.Lock()

        self._retrieve_event = threading.Event()
        self._stop_event     = threading.Event()
        self._log_queue: queue.Queue[str] = queue.Queue()

        self._poll_thread:   Optional[threading.Thread] = None
        self._restart_count  = 0

        load_category_matchers()
        self._build_ui()
        self._refresh_last_label()
        self._start_poll_thread()
        self._tick_countdown()
        self._drain_log_queue()

    # ── UI Construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        root = self._root
        root.title("Orbit Log Continuous Collector  /  Orbit Log 継続収集ツール")
        root.minsize(640, 500)

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"))

        main = ttk.Frame(root, padding=12)
        main.pack(fill="both", expand=True)

        # ── Info panel ──────────────────────────────────────────────────────
        info = ttk.LabelFrame(main,
                               text="Collection Info  /  収集情報", padding=8)
        info.pack(fill="x", pady=(0, 8))

        ttk.Label(info, text="File  /  ファイル:", anchor="w",
                  font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(info, text=self._xlsx_path, anchor="w",
                  wraplength=500).grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(info, text="Last retrieved  /  最終取得:", anchor="w",
                  font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._last_lbl = ttk.Label(info, text="—", anchor="w")
        self._last_lbl.grid(row=1, column=1, sticky="w", padx=6, pady=(4, 0))

        # ── Schedule panel ───────────────────────────────────────────────────
        sched = ttk.LabelFrame(main,
                                text="Schedule  /  スケジュール", padding=8)
        sched.pack(fill="x", pady=(0, 8))

        row0 = ttk.Frame(sched)
        row0.pack(fill="x")
        ttk.Label(row0, text="Next retrieval  /  次回取得:").pack(side="left")
        self._next_var = tk.StringVar(value=self._next_dt.strftime(OUT_TIME_FMT))
        ttk.Entry(row0, textvariable=self._next_var, width=22).pack(side="left", padx=6)
        ttk.Button(row0, text="Set  /  設定",
                   command=self._set_schedule).pack(side="left")

        row1 = ttk.Frame(sched)
        row1.pack(fill="x", pady=(6, 0))
        ttk.Label(row1, text="Countdown  /  カウントダウン:").pack(side="left")
        self._countdown_lbl = ttk.Label(row1, text="—",
                                         foreground="#1a6ba0",
                                         font=("Consolas", 9, "bold"))
        self._countdown_lbl.pack(side="left", padx=6)

        # ── Action buttons ───────────────────────────────────────────────────
        btns = ttk.Frame(main)
        btns.pack(fill="x", pady=(0, 8))

        ttk.Button(btns,
                   text="Retrieve Now  /  今すぐ取得",
                   style="Bold.TButton",
                   command=self._retrieve_now).pack(side="left", ipady=5, ipadx=10)

        ttk.Button(btns,
                   text="Stop Collector  /  収集停止",
                   command=self.on_close).pack(side="right", ipady=5, ipadx=10)

        # ── Activity log ─────────────────────────────────────────────────────
        log_frm = ttk.LabelFrame(main,
                                   text="Activity Log  /  活動ログ", padding=4)
        log_frm.pack(fill="both", expand=True)

        self._log_text = tk.Text(
            log_frm, height=12, state="disabled",
            font=("Consolas", 9), wrap="word",
            bg="#1e1e1e", fg="#d4d4d4",
        )
        vsb = ttk.Scrollbar(log_frm, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._log_text.pack(fill="both", expand=True)

    # ── Background threads ────────────────────────────────────────────────────
    def _start_poll_thread(self):
        t = threading.Thread(target=self._poll_worker, daemon=True, name="poll-worker")
        t.start()
        self._poll_thread = t
        # Watchdog runs once and restarts the poll thread if needed
        threading.Thread(target=self._watchdog, daemon=True, name="watchdog").start()

    def _poll_worker(self):
        try:
            client = Client(ORBIT_HOST, verify=TLS_VERIFY)
            client.authenticate_with_api_token(ORBIT_API_TOKEN)
            self._log(f"Connected to Orbit  /  Orbit に接続しました: {ORBIT_HOST}")
        except Exception as e:
            self._log(f"[ERROR] Could not connect to Orbit  /  Orbit への接続に失敗: {e}")
            return

        while not self._stop_event.is_set():
            self._wait_for_trigger()
            if self._stop_event.is_set():
                break

            end_dt = datetime.now(JST)

            # Determine start time: explicit on first run, then always from file
            if self._start_dt is not None:
                last_dt        = self._start_dt
                self._start_dt = None
            else:
                last_dt = get_last_timestamp(self._xlsx_path)

            if last_dt:
                self._log(f"Last entry  /  最終エントリ: {last_dt.strftime(OUT_TIME_FMT)}")
            else:
                self._log(f"No existing data — fetching last {BOOTSTRAP_MINUTES} min  /"
                          f"  既存データなし — 直近 {BOOTSTRAP_MINUTES} 分を取得します")

            try:
                rows = poll_once(client, last_dt, end_dt)
                if rows:
                    append_to_xlsx(self._xlsx_path, rows)
                    self._root.after(0, self._refresh_last_label)
                self._log(f"→ {len(rows)} new row(s) appended  /  {len(rows)} 件追加しました")
            except Exception as e:
                self._log(f"→ Fetch error  /  取得エラー: {e}")

            # Schedule next automatic retrieval
            next_dt = datetime.now(JST) + timedelta(seconds=POLL_INTERVAL_SECS)
            with self._next_dt_lock:
                self._next_dt = next_dt
            self._root.after(0, lambda dt=next_dt: self._next_var.set(dt.strftime(OUT_TIME_FMT)))
            self._log(f"→ Next check  /  次回確認: {next_dt.strftime(OUT_TIME_FMT)}")

    def _wait_for_trigger(self):
        """Block until next_dt is reached OR retrieve_event fires OR stop_event fires."""
        while not self._stop_event.is_set():
            with self._next_dt_lock:
                remaining = (self._next_dt - datetime.now(JST)).total_seconds()
            if remaining <= 0:
                break
            if self._retrieve_event.wait(timeout=min(remaining, 0.5)):
                self._retrieve_event.clear()
                break

    def _watchdog(self):
        """Restart the poll thread if it exits unexpectedly."""
        time.sleep(30)
        while not self._stop_event.is_set():
            time.sleep(10)
            if self._poll_thread and not self._poll_thread.is_alive() \
                    and not self._stop_event.is_set():
                self._restart_count += 1
                self._log(f"[WATCHDOG] Poll thread stopped — restarting (#{self._restart_count})"
                          f"  /  ポーリングスレッド停止 — 再起動中")
                time.sleep(RESTART_DELAY_SECS)
                if not self._stop_event.is_set():
                    self._start_poll_thread()

    # ── GUI event handlers ────────────────────────────────────────────────────
    def _retrieve_now(self):
        with self._next_dt_lock:
            self._next_dt = datetime.now(JST)
        self._next_var.set(self._next_dt.strftime(OUT_TIME_FMT))
        self._retrieve_event.set()
        self._log("Manual retrieval triggered  /  手動取得をトリガーしました")

    def _set_schedule(self):
        raw = self._next_var.get().strip()
        try:
            dt = datetime.strptime(raw, OUT_TIME_FMT).replace(tzinfo=JST)
        except ValueError:
            mb.showerror("Invalid date/time  /  日時エラー",
                         "Please use format: YYYY/MM/DD HH:MM:SS\n"
                         "形式: YYYY/MM/DD HH:MM:SS")
            return
        with self._next_dt_lock:
            self._next_dt = dt
        self._log(f"Schedule updated  /  スケジュール更新: {dt.strftime(OUT_TIME_FMT)}")

    def on_close(self):
        if mb.askyesno("Confirm Exit  /  終了の確認",
                        "Are you sure you want to stop the Orbit Log Collector?\n"
                        "Orbit Log 収集ツールを停止しますか？",
                        default="no"):
            self._stop_event.set()
            self._retrieve_event.set()   # unblock the wait loop immediately
            self._root.destroy()

    # ── Periodic GUI updates (main thread) ────────────────────────────────────
    def _refresh_last_label(self):
        last_dt = get_last_timestamp(self._xlsx_path)
        text = (last_dt.strftime(OUT_TIME_FMT) + " (JST)") if last_dt else "—"
        self._last_lbl.config(text=text)

    def _tick_countdown(self):
        with self._next_dt_lock:
            remaining = (self._next_dt - datetime.now(JST)).total_seconds()
        if remaining > 0:
            h, rem  = divmod(int(remaining), 3600)
            m, s    = divmod(rem, 60)
            text    = (f"{h}h {m:02d}m {s:02d}s" if h else f"{m:02d}m {s:02d}s")
        else:
            text = "retrieving…  /  取得中…"
        self._countdown_lbl.config(text=text)
        self._root.after(self._COUNTDOWN_MS, self._tick_countdown)

    def _drain_log_queue(self):
        try:
            while True:
                msg = self._log_queue.get_nowait()
                self._log_text.configure(state="normal")
                self._log_text.insert(tk.END, msg + "\n")
                self._log_text.see(tk.END)
                self._log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self._root.after(self._LOG_DRAIN_MS, self._drain_log_queue)

    def _log(self, msg: str):
        """Thread-safe: enqueue a timestamped message for the activity log."""
        ts = datetime.now(JST).strftime(OUT_TIME_FMT)
        self._log_queue.put(f"[{ts}]  {msg}")


# =============================================================================
# Entry point
# =============================================================================
def main():
    root = tk.Tk()
    root.withdraw()   # hide root while setup dialog is shown

    dialog = SetupDialog(root)
    root.wait_window(dialog)

    if not dialog.result:
        root.destroy()
        return

    root.deiconify()
    app = CollectorApp(root, dialog.result)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
