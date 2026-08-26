#!/usr/bin/env python3
"""
Continuous Orbit log collector — GUI edition.

Startup flow:
  1. Setup dialog: Continue from existing file  OR  Start new collection
  2. Main window: schedule panel, Retrieve Now, Settings, activity log.

Threading model:
  Main thread  → tkinter GUI
  poll-worker  → background thread; fetches and appends rows on schedule
  watchdog     → single background thread; restarts poll-worker if it dies

Config is persisted to orbit_config.json next to this script (gitignored —
it holds the API token). No code edits needed to change host, token, or
schedule; set them via Settings in the GUI on first run.

Dependencies:
  pip install tkcalendar   # optional — calendar popup in date pickers
"""
from __future__ import annotations

import json
import queue
import re
import threading
import time
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import tkinter.ttk as ttk
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from bosdyn.orbit.client import Client

try:
    from tkcalendar import DateEntry
    _HAS_TKCALENDAR = True
except ImportError:
    _HAS_TKCALENDAR = False


# =============================================================================
# Persistent configuration
# =============================================================================
CONFIG_FILE = Path(__file__).with_name("orbit_config.json")

# NOTE: host and token are intentionally BLANK here. They are secrets/site
# config and must never live in source. Set them once via Settings in the GUI
# (or drop them into orbit_config.json, which is gitignored).
DEFAULT_CONFIG: Dict[str, Any] = {
    "orbit_host":          "",
    "api_token":           "",
    "tls_verify":          False,
    "run_fetch_limit":     10000,
    "bootstrap_minutes":   30,
    "max_connect_retries": 5,
    "connect_retry_base":  10,
    "schedule_mode":       "interval",   # "interval" | "times"
    "poll_interval_mins":  30,
    "scheduled_times":     ["09:00", "13:00", "17:00"],
}

def load_config() -> Dict[str, Any]:
    cfg = dict(DEFAULT_CONFIG)
    try:
        stored = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        cfg.update({k: v for k, v in stored.items() if k in DEFAULT_CONFIG})
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"[WARN] Could not load config: {e}")
    return cfg

def save_config(cfg: Dict[str, Any]) -> None:
    try:
        CONFIG_FILE.write_text(
            json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except Exception as e:
        print(f"[WARN] Could not save config: {e}")


# =============================================================================
# Non-configurable constants
# =============================================================================
RESTART_DELAY_SECS = 15
LOG_MAX_LINES      = 500

HEADERS      = ["Robot Serial Number", "Date and time", "Route", "Orbit log"]
OUT_TIME_FMT = "%Y/%m/%d %H:%M:%S"
JST          = timezone(timedelta(hours=9))

INCLUDE_KEYWORDS = [""]
IGNORE_WORDS     = []
MISSION_IGNORE   = []

PATTERNS_XLSX  = "log_patterns_example.xlsx"
PATTERNS_SHEET = "Log Patterns"

CATEGORY_COLORS: Dict[str, str] = {
    "Software / System Issues":          "FF4C4C",
    "Navigation — Route Conflict":       "FF8042",
    "Navigation — Stuck (Goal Blocked)": "FF9900",
    "Navigation — Stuck":                "FFB347",
    "Navigation — Path Not Found":       "FFD966",
    "Entity Detection":                  "FFDC73",
    "Task Completion":                   "C6EFCE",
    "Docking / Undocking":               "BDD7EE",
    "Run Lifecycle":                     "DDEEFF",
    "Task Skipped":                      "F2F2F2",
    "Mission Interruption":              "F0EBFF",
    "Battery":                           "FEFCE8",
}
COLOR_FALLBACK = "F3F3F3"

PATTERN_OVERRIDES: List[tuple] = [
    ("after stuck",          "FFD966"),
    ("after route conflict", "FFD966"),
]
KEYWORD_OVERRIDES: List[tuple] = [
    ("battery temp",    "FFD966"),
    ("battery error",   "FFD966"),
    ("battery fault",   "FFD966"),
    ("battery failure", "FFD966"),
    ("thermal",         "FFD966"),
]


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

_category_matchers: List[tuple] = []

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

def append_to_xlsx(path: str, rows: List[List[Any]]) -> Optional[datetime]:
    """Append rows to xlsx. Returns the latest timestamp written, or None."""
    try:
        wb = load_workbook(path)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Orbit Logs"
        ws.append(HEADERS)
        ws.freeze_panes = "A2"

    latest_dt: Optional[datetime] = None
    for r in rows:
        ws.append(r)
        fill = get_row_fill(r[3] if len(r) > 3 else "")
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill
        try:
            dt = datetime.strptime(str(r[1]).strip(), OUT_TIME_FMT).replace(tzinfo=JST)
            if latest_dt is None or dt > latest_dt:
                latest_dt = dt
        except Exception:
            pass

    for col_idx in range(1, len(HEADERS) + 1):
        col_letter = get_column_letter(col_idx)
        if ws.max_row > 1:
            max_len = max(len(str(c.value)) if c.value else 0 for c in ws[col_letter])
            ws.column_dimensions[col_letter].width = min(max(14, max_len + 2), 90)
    wb.save(path)
    return latest_dt


# ---------------------------------------------------------------------------
# Poll cycle
# ---------------------------------------------------------------------------
def poll_once(
    client: Client,
    start_dt: Optional[datetime],
    end_dt: datetime,
    cfg: Dict[str, Any],
) -> List[List[Any]]:
    """Fetch log rows in (start_dt, end_dt]. Returns rows (does NOT write xlsx)."""
    if start_dt is None:
        start_dt = end_dt - timedelta(minutes=cfg["bootstrap_minutes"])

    resp = client.get_runs(params={"orderBy": "newest", "limit": cfg["run_fetch_limit"]})
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
        except Exception as exc:
            print(f"[WARN] Failed to fetch log for run {run_id}: {exc}")

    rows.sort(key=lambda x: x[1])

    if start_dt:
        cutoff = start_dt.strftime(OUT_TIME_FMT)
        rows   = [r for r in rows if r[1] > cutoff]

    return rows


# ---------------------------------------------------------------------------
# Schedule helpers
# ---------------------------------------------------------------------------
def compute_next_scheduled_dt(times: List[str]) -> datetime:
    """Given ['HH:MM', ...], return the nearest upcoming datetime in JST."""
    now = datetime.now(JST)
    candidates: List[datetime] = []
    for t in times:
        try:
            h, m = map(int, t.strip().split(":"))
            candidate = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if candidate <= now:
                candidate += timedelta(days=1)
            candidates.append(candidate)
        except Exception:
            pass
    return min(candidates) if candidates else now + timedelta(hours=1)


# =============================================================================
# Reusable widgets
# =============================================================================
class DateTimePicker(ttk.Frame):
    """
    Date + time picker.
    Uses tkcalendar.DateEntry (calendar popup) when available,
    falls back to a plain Entry for the date portion.
    """

    def __init__(self, parent, initial: Optional[datetime] = None, **kwargs):
        super().__init__(parent, **kwargs)
        if initial is None:
            initial = datetime.now(JST)

        if _HAS_TKCALENDAR:
            self._date_entry = DateEntry(self, width=12, date_pattern="yyyy/mm/dd",
                                          firstweekday="sunday")
            self._date_entry.set_date(initial.date())
            self._date_entry.pack(side="left", padx=(0, 4))
        else:
            self._date_var = tk.StringVar(value=initial.strftime("%Y/%m/%d"))
            ttk.Entry(self, textvariable=self._date_var, width=12).pack(side="left", padx=(0, 4))

        self._hour_var   = tk.StringVar(value=f"{initial.hour:02d}")
        self._minute_var = tk.StringVar(value=f"{initial.minute:02d}")
        self._second_var = tk.StringVar(value=f"{initial.second:02d}")

        ttk.Spinbox(self, from_=0, to=23, width=3, textvariable=self._hour_var,
                    format="%02.0f", wrap=True).pack(side="left")
        ttk.Label(self, text=":").pack(side="left")
        ttk.Spinbox(self, from_=0, to=59, width=3, textvariable=self._minute_var,
                    format="%02.0f", wrap=True).pack(side="left")
        ttk.Label(self, text=":").pack(side="left")
        ttk.Spinbox(self, from_=0, to=59, width=3, textvariable=self._second_var,
                    format="%02.0f", wrap=True).pack(side="left")

    def get(self) -> Optional[datetime]:
        try:
            if _HAS_TKCALENDAR:
                date_str = self._date_entry.get_date().strftime("%Y/%m/%d")
            else:
                date_str = self._date_var.get().strip()
            h = int(self._hour_var.get())
            m = int(self._minute_var.get())
            s = int(self._second_var.get())
            return datetime.strptime(
                f"{date_str} {h:02d}:{m:02d}:{s:02d}", OUT_TIME_FMT
            ).replace(tzinfo=JST)
        except Exception:
            return None

    def set(self, dt: datetime) -> None:
        if _HAS_TKCALENDAR:
            self._date_entry.set_date(dt.date())
        else:
            self._date_var.set(dt.strftime("%Y/%m/%d"))
        self._hour_var.set(f"{dt.hour:02d}")
        self._minute_var.set(f"{dt.minute:02d}")
        self._second_var.set(f"{dt.second:02d}")

    def set_state(self, state: str) -> None:
        for child in self.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass


class TimePicker(ttk.Frame):
    """HH:MM spinbox pair for scheduled-times entry."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._h = tk.StringVar(value="09")
        self._m = tk.StringVar(value="00")
        ttk.Spinbox(self, from_=0, to=23, width=3, textvariable=self._h,
                    format="%02.0f", wrap=True).pack(side="left")
        ttk.Label(self, text=":").pack(side="left")
        ttk.Spinbox(self, from_=0, to=59, width=3, textvariable=self._m,
                    format="%02.0f", wrap=True).pack(side="left")

    def get(self) -> str:
        return f"{int(self._h.get()):02d}:{int(self._m.get()):02d}"

    def set(self, hhmm: str) -> None:
        h, m = hhmm.split(":")
        self._h.set(h)
        self._m.set(m)


# =============================================================================
# Settings dialog
# =============================================================================
class SettingsDialog(tk.Toplevel):
    """
    Edit orbit_config.json from the UI.
    on_save(new_cfg) is called with the merged config when the user saves.
    """

    def __init__(self, parent, cfg: Dict[str, Any], on_save):
        super().__init__(parent)
        self._cfg     = dict(cfg)
        self._on_save = on_save
        self.title("Settings  /  設定")
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
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(12, 4))

        # ── Tab 1: Connection ────────────────────────────────────────────────
        conn = ttk.Frame(nb, padding=12)
        nb.add(conn, text="Connection  /  接続")
        conn.columnconfigure(1, weight=1)

        ttk.Label(conn, text="Orbit Host  /  Orbit ホスト:").grid(
            row=0, column=0, sticky="w", pady=6)
        self._host_var = tk.StringVar(value=self._cfg["orbit_host"])
        ttk.Entry(conn, textvariable=self._host_var, width=32).grid(
            row=0, column=1, sticky="w", padx=8)

        ttk.Label(conn, text="API Token:").grid(row=1, column=0, sticky="w", pady=6)
        tok_frame = ttk.Frame(conn)
        tok_frame.grid(row=1, column=1, sticky="w", padx=8)
        self._token_var = tk.StringVar(value=self._cfg["api_token"])
        self._token_entry = ttk.Entry(tok_frame, textvariable=self._token_var,
                                       width=30, show="*")
        self._token_entry.pack(side="left")
        self._show_tok = tk.BooleanVar(value=False)
        ttk.Checkbutton(tok_frame, text="Show  /  表示",
                        variable=self._show_tok,
                        command=self._toggle_token).pack(side="left", padx=6)

        self._tls_var = tk.BooleanVar(value=self._cfg["tls_verify"])
        ttk.Checkbutton(conn, text="Verify TLS certificate  /  TLS証明書を検証する",
                        variable=self._tls_var).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=6)

        ttk.Label(conn, text="Connection settings take effect after Reconnect.",
                  foreground="grey", font=("Segoe UI", 8)).grid(
            row=3, column=0, columnspan=2, sticky="w")

        # ── Tab 2: Advanced ──────────────────────────────────────────────────
        adv = ttk.Frame(nb, padding=12)
        nb.add(adv, text="Advanced  /  詳細")
        adv.columnconfigure(1, weight=1)

        adv_fields = [
            ("Run fetch limit  /  実行取得上限:",        "run_fetch_limit",     1, 100000),
            ("Bootstrap look-back (min)  /  初回遡及分数:", "bootstrap_minutes",   1, 10080),
            ("Max connect retries  /  最大再接続回数:",  "max_connect_retries", 1, 20),
            ("Connect retry base (s)  /  再接続待機基準秒:", "connect_retry_base",  1, 300),
        ]
        self._adv_vars: Dict[str, tk.IntVar] = {}
        for i, (label, key, lo, hi) in enumerate(adv_fields):
            ttk.Label(adv, text=label).grid(row=i, column=0, sticky="w", pady=5)
            var = tk.IntVar(value=int(self._cfg[key]))
            ttk.Spinbox(adv, textvariable=var, from_=lo, to=hi, width=10).grid(
                row=i, column=1, sticky="w", padx=8)
            self._adv_vars[key] = var

        # ── Buttons ──────────────────────────────────────────────────────────
        btn_row = ttk.Frame(self)
        btn_row.pack(pady=10)
        ttk.Button(btn_row, text="Save  /  保存",
                   command=self._save).pack(side="left", padx=8)
        ttk.Button(btn_row, text="Cancel  /  キャンセル",
                   command=self.destroy).pack(side="left")

    def _toggle_token(self):
        self._token_entry.configure(show="" if self._show_tok.get() else "*")

    def _save(self):
        host  = self._host_var.get().strip()
        token = self._token_var.get().strip()
        if not host:
            mb.showerror("Missing host  /  ホストエラー",
                         "Orbit host cannot be empty.\nOrbit ホストを入力してください。",
                         parent=self)
            return
        if not token:
            mb.showerror("Missing token  /  トークンエラー",
                         "API token cannot be empty.\nAPIトークンを入力してください。",
                         parent=self)
            return
        self._cfg["orbit_host"]          = host
        self._cfg["api_token"]           = token
        self._cfg["tls_verify"]          = self._tls_var.get()
        for key, var in self._adv_vars.items():
            try:
                self._cfg[key] = int(var.get())
            except Exception:
                pass
        save_config(self._cfg)
        self._on_save(self._cfg)
        self.destroy()


# =============================================================================
# Setup dialogs
# =============================================================================
class _NewCollectionDialog(tk.Toplevel):
    """Sub-dialog: collect start/end date-time and save path for a new log file."""

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
        now = datetime.now(JST)

        ttk.Label(self, text="New Collection Setup  /  新規収集設定",
                  font=("Segoe UI", 11, "bold")).pack(**pad)

        # ── Start date/time ──────────────────────────────────────────────────
        f1 = ttk.LabelFrame(self, text="Start date / time (JST)  /  開始日時 (JST)", padding=8)
        f1.pack(fill="x", padx=20, pady=4)
        self._start_picker = DateTimePicker(f1, initial=now)
        self._start_picker.pack(anchor="w")

        # ── End date/time ────────────────────────────────────────────────────
        f2 = ttk.LabelFrame(self, text="End date / time (JST)  /  終了日時 (JST)", padding=8)
        f2.pack(fill="x", padx=20, pady=4)
        self._use_now_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            f2,
            text="Use current time as end ('now')  /  終了時刻を「現在時刻」にする",
            variable=self._use_now_var,
            command=self._toggle_end,
        ).pack(anchor="w")
        self._end_picker = DateTimePicker(f2, initial=now)
        self._end_picker.pack(anchor="w", pady=(4, 0))
        self._toggle_end()

        # ── Save location ────────────────────────────────────────────────────
        f3 = ttk.LabelFrame(self, text="Save location  /  保存先", padding=8)
        f3.pack(fill="x", padx=20, pady=4)
        row = ttk.Frame(f3)
        row.pack(fill="x")
        self._path_var = tk.StringVar(value=f"orbit_logs_{now.strftime('%Y%m%d')}.xlsx")
        ttk.Entry(row, textvariable=self._path_var, width=32).pack(side="left")
        ttk.Button(row, text="Browse  /  参照", command=self._browse).pack(side="left", padx=6)

        # ── Buttons ──────────────────────────────────────────────────────────
        btn = ttk.Frame(self)
        btn.pack(pady=14)
        ttk.Button(btn, text="Start  /  開始", command=self._confirm).pack(side="left", padx=8)
        ttk.Button(btn, text="Cancel  /  キャンセル", command=self.destroy).pack(side="left")

    def _toggle_end(self):
        self._end_picker.set_state("disabled" if self._use_now_var.get() else "normal")

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
        start_dt = self._start_picker.get()
        if start_dt is None:
            mb.showerror("Invalid start date/time  /  開始日時エラー",
                         "Could not parse start date/time.\n開始日時を正しく入力してください。",
                         parent=self)
            return

        if self._use_now_var.get():
            end_dt = None
        else:
            end_dt = self._end_picker.get()
            if end_dt is None:
                mb.showerror("Invalid end date/time  /  終了日時エラー",
                             "Could not parse end date/time.\n終了日時を正しく入力してください。",
                             parent=self)
                return
            if end_dt <= start_dt:
                mb.showerror("Date range error  /  日時範囲エラー",
                             "End time must be after start time.\n"
                             "終了日時は開始日時より後にしてください。",
                             parent=self)
                return

        path = self._path_var.get().strip()
        if not path:
            mb.showerror("Missing path  /  パスエラー",
                         "Please specify a save location.\n保存先を指定してください。",
                         parent=self)
            return

        self._parent._accept(path, start_dt, end_dt)
        self.destroy()


class SetupDialog(tk.Toplevel):
    """
    Modal startup dialog.
    self.result → {"xlsx_path": str, "start_dt": datetime|None, "end_dt": datetime|None}
    or None if cancelled.
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
                             "No 'Orbit Logs' sheet found.\n"
                             "'Orbit Logs' シートが見つかりません。",
                             parent=self)
                wb.close()
                return
            wb.close()
        except Exception as e:
            mb.showerror("Error  /  エラー", str(e), parent=self)
            return
        self._accept(path, None, None)

    def _start_new(self):
        _NewCollectionDialog(self)

    def _accept(self, xlsx_path: str, start_dt: Optional[datetime], end_dt: Optional[datetime]):
        self.result = {"xlsx_path": xlsx_path, "start_dt": start_dt, "end_dt": end_dt}
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# =============================================================================
# Main GUI application
# =============================================================================
class CollectorApp:
    """
    Main window.  Thread-safety notes:
      - GUI updates only on main thread (root.after or periodic callbacks).
      - poll-worker communicates via _log_queue.
      - _next_dt protected by _next_dt_lock.
      - _last_known_dt protected by _last_dt_lock.
      - _cfg is a plain dict; structural changes (mode, times, interval)
        happen only on the main thread via GUI handlers — safe to read in
        poll-worker without locking because Python GIL protects dict reads.
    """

    _COUNTDOWN_MS = 1_000
    _LOG_DRAIN_MS =   200

    def __init__(self, root: tk.Tk, collection_cfg: Dict, app_cfg: Dict):
        self._root      = root
        self._xlsx_path = collection_cfg["xlsx_path"]
        self._start_dt  = collection_cfg.get("start_dt")
        self._end_dt    = collection_cfg.get("end_dt")
        self._cfg       = app_cfg   # live config; updated by SettingsDialog

        self._next_dt      = datetime.now(JST) + timedelta(seconds=8)
        self._next_dt_lock = threading.Lock()

        self._last_known_dt: Optional[datetime] = None
        self._last_dt_lock  = threading.Lock()

        self._retrieve_event  = threading.Event()
        self._stop_event      = threading.Event()
        self._reconnect_flag  = False   # set by Settings save to trigger reconnect
        self._log_queue: queue.Queue[str] = queue.Queue()

        self._poll_thread:     Optional[threading.Thread] = None
        self._restart_count    = 0
        self._watchdog_running = False

        load_category_matchers()
        self._build_ui()
        self._init_last_timestamp()
        self._start_poll_thread()
        self._tick_countdown()
        self._drain_log_queue()

    def _init_last_timestamp(self):
        dt = get_last_timestamp(self._xlsx_path)
        with self._last_dt_lock:
            self._last_known_dt = dt
        self._refresh_last_label()

    # ── UI Construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        root = self._root
        root.title("Orbit Log Continuous Collector  /  Orbit Log 継続収集ツール")
        root.minsize(700, 580)

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"))

        main = ttk.Frame(root, padding=12)
        main.pack(fill="both", expand=True)

        # ── Info panel ──────────────────────────────────────────────────────
        info = ttk.LabelFrame(main, text="Collection Info  /  収集情報", padding=8)
        info.pack(fill="x", pady=(0, 8))
        info.columnconfigure(1, weight=1)

        ttk.Label(info, text="Orbit host  /  Orbit ホスト:", anchor="w",
                  font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        self._host_lbl = ttk.Label(info, text=self._cfg["orbit_host"], anchor="w")
        self._host_lbl.grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(info, text="File  /  ファイル:", anchor="w",
                  font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(info, text=self._xlsx_path, anchor="w",
                  wraplength=520).grid(row=1, column=1, sticky="w", padx=6, pady=(4, 0))

        ttk.Label(info, text="Last retrieved  /  最終取得:", anchor="w",
                  font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 0))
        self._last_lbl = ttk.Label(info, text="—", anchor="w")
        self._last_lbl.grid(row=2, column=1, sticky="w", padx=6, pady=(4, 0))

        if self._end_dt:
            ttk.Label(info, text="Collection end  /  収集終了:", anchor="w",
                      font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky="w", pady=(4, 0))
            ttk.Label(info, text=self._end_dt.strftime(OUT_TIME_FMT) + " (JST)",
                      anchor="w").grid(row=3, column=1, sticky="w", padx=6, pady=(4, 0))

        # ── Schedule panel ───────────────────────────────────────────────────
        sched = ttk.LabelFrame(main, text="Schedule  /  スケジュール", padding=8)
        sched.pack(fill="x", pady=(0, 8))

        # Mode selector row
        mode_row = ttk.Frame(sched)
        mode_row.pack(fill="x", pady=(0, 6))
        ttk.Label(mode_row, text="Mode  /  モード:").pack(side="left")
        self._mode_var = tk.StringVar(value=self._cfg["schedule_mode"])
        mode_cb = ttk.Combobox(
            mode_row, textvariable=self._mode_var,
            values=["interval", "times"],
            state="readonly", width=10,
        )
        mode_cb.pack(side="left", padx=6)
        ttk.Label(mode_row,
                  text="interval = repeat every N min  |  times = daily at set times",
                  foreground="grey", font=("Segoe UI", 8)).pack(side="left", padx=6)
        mode_cb.bind("<<ComboboxSelected>>", self._on_mode_change)

        # ── Interval sub-panel ───────────────────────────────────────────────
        self._interval_frame = ttk.Frame(sched)

        int_row = ttk.Frame(self._interval_frame)
        int_row.pack(fill="x")
        ttk.Label(int_row, text="Interval  /  間隔 (min):").pack(side="left")
        self._interval_var = tk.IntVar(value=self._cfg["poll_interval_mins"])
        ttk.Spinbox(int_row, from_=1, to=1440, width=6,
                    textvariable=self._interval_var, wrap=False).pack(side="left", padx=6)
        ttk.Button(int_row, text="Apply  /  適用",
                   command=self._apply_interval).pack(side="left")

        # ── Daily-times sub-panel ────────────────────────────────────────────
        self._times_frame = ttk.Frame(sched)

        times_top = ttk.Frame(self._times_frame)
        times_top.pack(fill="x")

        # Listbox of scheduled times
        lb_frame = ttk.Frame(times_top)
        lb_frame.pack(side="left")
        self._times_lb = tk.Listbox(lb_frame, height=5, width=10,
                                     selectmode="single", font=("Consolas", 10))
        lb_sb = ttk.Scrollbar(lb_frame, command=self._times_lb.yview)
        self._times_lb.configure(yscrollcommand=lb_sb.set)
        lb_sb.pack(side="right", fill="y")
        self._times_lb.pack(side="left", fill="both")
        for t in self._cfg["scheduled_times"]:
            self._times_lb.insert(tk.END, t)

        # Add / remove controls
        ctrl = ttk.Frame(times_top)
        ctrl.pack(side="left", padx=12, anchor="n")
        ttk.Label(ctrl, text="Add time  /  時刻を追加:").pack(anchor="w")
        self._time_picker = TimePicker(ctrl)
        self._time_picker.pack(anchor="w", pady=4)
        ttk.Button(ctrl, text="Add  /  追加",
                   command=self._add_time).pack(anchor="w")
        ttk.Button(ctrl, text="Remove selected  /  選択を削除",
                   command=self._remove_time).pack(anchor="w", pady=(6, 0))

        # ── Shared: Next retrieval + countdown ──────────────────────────────
        shared = ttk.Frame(sched)
        shared.pack(fill="x", pady=(8, 0))

        next_row = ttk.Frame(shared)
        next_row.pack(fill="x")
        ttk.Label(next_row, text="Next retrieval  /  次回取得:").pack(side="left")
        self._next_var = tk.StringVar(value=self._next_dt.strftime(OUT_TIME_FMT))
        self._next_entry = ttk.Entry(next_row, textvariable=self._next_var, width=22)
        self._next_entry.pack(side="left", padx=6)
        self._set_btn = ttk.Button(next_row, text="Set  /  設定",
                                    command=self._set_schedule)
        self._set_btn.pack(side="left")

        cd_row = ttk.Frame(shared)
        cd_row.pack(fill="x", pady=(4, 0))
        ttk.Label(cd_row, text="Countdown  /  カウントダウン:").pack(side="left")
        self._countdown_lbl = ttk.Label(cd_row, text="—",
                                         foreground="#1a6ba0",
                                         font=("Consolas", 9, "bold"))
        self._countdown_lbl.pack(side="left", padx=6)

        # Show correct sub-panel immediately
        self._refresh_sched_panel()

        # ── Action buttons ───────────────────────────────────────────────────
        btns = ttk.Frame(main)
        btns.pack(fill="x", pady=(0, 8))

        ttk.Button(btns,
                   text="Retrieve Now  /  今すぐ取得",
                   style="Bold.TButton",
                   command=self._retrieve_now).pack(side="left", ipady=5, ipadx=10)

        ttk.Button(btns,
                   text="Reconnect  /  再接続",
                   command=self._reconnect).pack(side="left", padx=8, ipady=5, ipadx=6)

        ttk.Button(btns,
                   text="Settings  /  設定",
                   command=self._open_settings).pack(side="left", ipady=5, ipadx=6)

        ttk.Button(btns,
                   text="Stop Collector  /  収集停止",
                   command=self.on_close).pack(side="right", ipady=5, ipadx=10)

        # ── Activity log ─────────────────────────────────────────────────────
        log_frm = ttk.LabelFrame(main, text="Activity Log  /  活動ログ", padding=4)
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

    # ── Schedule panel helpers ────────────────────────────────────────────────
    def _refresh_sched_panel(self):
        mode = self._cfg["schedule_mode"]
        if mode == "interval":
            self._times_frame.pack_forget()
            self._interval_frame.pack(fill="x", pady=(0, 4))
            # "Set" button only meaningful in interval mode
            self._next_entry.configure(state="normal")
            self._set_btn.configure(state="normal")
        else:
            self._interval_frame.pack_forget()
            self._times_frame.pack(fill="x", pady=(0, 4))
            # Next retrieval is computed automatically; entry is read-only
            self._next_entry.configure(state="readonly")
            self._set_btn.configure(state="disabled")
            self._update_next_from_times()

    def _on_mode_change(self, _event=None):
        mode = self._mode_var.get()
        self._cfg["schedule_mode"] = mode
        save_config(self._cfg)
        self._refresh_sched_panel()
        self._log(f"Schedule mode → {mode}  /  スケジュールモード変更: {mode}")

    def _update_next_from_times(self):
        times = self._cfg["scheduled_times"]
        if times:
            next_dt = compute_next_scheduled_dt(times)
            with self._next_dt_lock:
                self._next_dt = next_dt
            self._next_var.set(next_dt.strftime(OUT_TIME_FMT))

    def _add_time(self):
        hhmm = self._time_picker.get()
        existing = list(self._times_lb.get(0, tk.END))
        if hhmm in existing:
            mb.showinfo("Duplicate  /  重複",
                        f"{hhmm} is already in the list.\n{hhmm} はすでに追加されています。",
                        parent=self._root)
            return
        existing.append(hhmm)
        existing.sort()
        self._times_lb.delete(0, tk.END)
        for t in existing:
            self._times_lb.insert(tk.END, t)
        self._cfg["scheduled_times"] = existing
        save_config(self._cfg)
        self._update_next_from_times()
        self._log(f"Scheduled time added  /  時刻追加: {hhmm}")

    def _remove_time(self):
        sel = self._times_lb.curselection()
        if not sel:
            return
        hhmm = self._times_lb.get(sel[0])
        self._times_lb.delete(sel[0])
        times = list(self._times_lb.get(0, tk.END))
        self._cfg["scheduled_times"] = times
        save_config(self._cfg)
        if times:
            self._update_next_from_times()
        self._log(f"Scheduled time removed  /  時刻削除: {hhmm}")

    # ── Settings / Reconnect ──────────────────────────────────────────────────
    def _open_settings(self):
        SettingsDialog(self._root, self._cfg, on_save=self._on_settings_saved)

    def _on_settings_saved(self, new_cfg: Dict):
        self._cfg = new_cfg
        self._host_lbl.config(text=new_cfg["orbit_host"])
        self._log(
            "Settings saved. Use 'Reconnect' to apply connection changes."
            "  /  設定を保存しました。接続設定の反映は「再接続」ボタンで。"
        )

    def _reconnect(self):
        """Restart the poll worker so it picks up any new connection settings."""
        self._log("Reconnecting…  /  再接続中…")
        self._reconnect_flag = True
        self._retrieve_event.set()   # unblock _wait_for_trigger so worker exits cleanly

    # ── Background threads ────────────────────────────────────────────────────
    def _start_poll_thread(self):
        t = threading.Thread(target=self._poll_worker, daemon=True, name="poll-worker")
        t.start()
        self._poll_thread = t
        if not self._watchdog_running:
            self._watchdog_running = True
            threading.Thread(target=self._watchdog, daemon=True, name="watchdog").start()

    def _poll_worker(self):
        cfg = self._cfg   # snapshot at connect time

        # ── Connect with exponential-backoff retry ────────────────────────────
        client = None
        for attempt in range(1, cfg["max_connect_retries"] + 1):
            try:
                client = Client(cfg["orbit_host"], verify=cfg["tls_verify"])
                client.authenticate_with_api_token(cfg["api_token"])
                self._log(f"Connected  /  接続しました: {cfg['orbit_host']}")
                break
            except Exception as e:
                if attempt < cfg["max_connect_retries"]:
                    delay = cfg["connect_retry_base"] * (2 ** (attempt - 1))
                    self._log(
                        f"[WARN] Attempt {attempt}/{cfg['max_connect_retries']} failed, "
                        f"retry in {delay}s  /  接続失敗、{delay}秒後に再試行: {e}"
                    )
                    for _ in range(delay * 2):
                        if self._stop_event.is_set() or self._reconnect_flag:
                            return
                        time.sleep(0.5)
                else:
                    self._log(
                        f"[ERROR] Could not connect after {cfg['max_connect_retries']} attempts  /"
                        f"  {cfg['max_connect_retries']} 回試行しましたが接続できません: {e}"
                    )
                    return

        while not self._stop_event.is_set():
            self._wait_for_trigger()

            # Reconnect requested — exit so watchdog (or explicit restart) relaunches
            if self._reconnect_flag:
                self._reconnect_flag = False
                self._log("Reconnect requested — restarting worker  /  再接続のためワーカーを再起動")
                self._root.after(0, self._start_poll_thread)
                return

            if self._stop_event.is_set():
                break

            live_cfg = self._cfg   # re-read live config each cycle
            end_dt   = self._end_dt if self._end_dt is not None else datetime.now(JST)

            if self._start_dt is not None:
                last_dt        = self._start_dt
                self._start_dt = None
            else:
                with self._last_dt_lock:
                    last_dt = self._last_known_dt
                if last_dt is None:
                    last_dt = get_last_timestamp(self._xlsx_path)
                    with self._last_dt_lock:
                        self._last_known_dt = last_dt

            if last_dt:
                self._log(
                    f"Fetching  /  取得: "
                    f"{last_dt.strftime(OUT_TIME_FMT)} → {end_dt.strftime(OUT_TIME_FMT)}"
                )
            else:
                self._log(
                    f"No existing data — bootstrap {live_cfg['bootstrap_minutes']} min  /"
                    f"  既存データなし — 直近 {live_cfg['bootstrap_minutes']} 分を取得"
                )

            try:
                rows = poll_once(client, last_dt, end_dt, live_cfg)
                if rows:
                    latest_dt = append_to_xlsx(self._xlsx_path, rows)
                    if latest_dt:
                        with self._last_dt_lock:
                            if self._last_known_dt is None or latest_dt > self._last_known_dt:
                                self._last_known_dt = latest_dt
                    self._root.after(0, self._refresh_last_label)
                self._log(f"→ {len(rows)} row(s) appended  /  {len(rows)} 件追加")
            except Exception as e:
                self._log(f"→ Fetch error  /  取得エラー: {e}")

            if self._end_dt is not None and datetime.now(JST) >= self._end_dt:
                self._log("End time reached — stopping  /  終了時刻に達しました")
                self._stop_event.set()
                self._root.after(0, self._root.destroy)
                return

            # Compute next retrieval based on current mode
            if live_cfg["schedule_mode"] == "interval":
                next_dt = datetime.now(JST) + timedelta(minutes=live_cfg["poll_interval_mins"])
            else:
                next_dt = compute_next_scheduled_dt(live_cfg["scheduled_times"])

            with self._next_dt_lock:
                self._next_dt = next_dt
            self._root.after(0, lambda dt=next_dt: self._next_var.set(dt.strftime(OUT_TIME_FMT)))
            self._log(f"→ Next check  /  次回: {next_dt.strftime(OUT_TIME_FMT)}")

    def _wait_for_trigger(self):
        while not self._stop_event.is_set() and not self._reconnect_flag:
            with self._next_dt_lock:
                remaining = (self._next_dt - datetime.now(JST)).total_seconds()
            if remaining <= 0:
                break
            if self._retrieve_event.wait(timeout=min(remaining, 0.5)):
                self._retrieve_event.clear()
                break

    def _watchdog(self):
        time.sleep(30)
        while not self._stop_event.is_set():
            time.sleep(10)
            if (self._poll_thread and not self._poll_thread.is_alive()
                    and not self._stop_event.is_set()
                    and not self._reconnect_flag):
                self._restart_count += 1
                self._log(
                    f"[WATCHDOG] Restarting poll worker (#{self._restart_count})"
                    f"  /  ポーリングスレッド再起動"
                )
                time.sleep(RESTART_DELAY_SECS)
                if not self._stop_event.is_set():
                    t = threading.Thread(target=self._poll_worker,
                                          daemon=True, name="poll-worker")
                    t.start()
                    self._poll_thread = t
        self._watchdog_running = False

    # ── GUI event handlers ────────────────────────────────────────────────────
    def _retrieve_now(self):
        with self._next_dt_lock:
            self._next_dt = datetime.now(JST)
        self._next_var.set(self._next_dt.strftime(OUT_TIME_FMT))
        self._retrieve_event.set()
        self._log("Manual retrieval triggered  /  手動取得")

    def _set_schedule(self):
        raw = self._next_var.get().strip()
        try:
            dt = datetime.strptime(raw, OUT_TIME_FMT).replace(tzinfo=JST)
        except ValueError:
            mb.showerror("Invalid date/time  /  日時エラー",
                         "Format: YYYY/MM/DD HH:MM:SS\n形式: YYYY/MM/DD HH:MM:SS")
            return
        with self._next_dt_lock:
            self._next_dt = dt
        self._log(f"Next retrieval set  /  次回取得設定: {dt.strftime(OUT_TIME_FMT)}")

    def _apply_interval(self):
        try:
            mins = int(self._interval_var.get())
            if mins < 1:
                raise ValueError
        except (ValueError, tk.TclError):
            mb.showerror("Invalid interval  /  間隔エラー",
                         "Enter a positive integer (minutes).\n正の整数を分単位で入力してください。")
            return
        self._cfg["poll_interval_mins"] = mins
        save_config(self._cfg)
        self._log(f"Interval updated  /  間隔変更: {mins} min")

    def on_close(self):
        if mb.askyesno("Confirm Exit  /  終了の確認",
                        "Stop the Orbit Log Collector?\nOrbit Log 収集を停止しますか？",
                        default="no"):
            self._stop_event.set()
            self._retrieve_event.set()
            self._root.destroy()

    # ── Periodic GUI updates (main thread) ────────────────────────────────────
    def _refresh_last_label(self):
        with self._last_dt_lock:
            last_dt = self._last_known_dt
        self._last_lbl.config(
            text=(last_dt.strftime(OUT_TIME_FMT) + " (JST)") if last_dt else "—"
        )

    def _tick_countdown(self):
        with self._next_dt_lock:
            remaining = (self._next_dt - datetime.now(JST)).total_seconds()
        if remaining > 0:
            h, rem = divmod(int(remaining), 3600)
            m, s   = divmod(rem, 60)
            text   = f"{h}h {m:02d}m {s:02d}s" if h else f"{m:02d}m {s:02d}s"
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
                line_count = int(self._log_text.index(tk.END).split(".")[0]) - 1
                if line_count > LOG_MAX_LINES:
                    self._log_text.delete("1.0", f"{line_count - LOG_MAX_LINES + 1}.0")
                self._log_text.see(tk.END)
                self._log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self._root.after(self._LOG_DRAIN_MS, self._drain_log_queue)

    def _log(self, msg: str):
        ts = datetime.now(JST).strftime(OUT_TIME_FMT)
        self._log_queue.put(f"[{ts}]  {msg}")


# =============================================================================
# Entry point
# =============================================================================
def main():
    app_cfg = load_config()

    root = tk.Tk()
    root.withdraw()

    dialog = SetupDialog(root)
    root.wait_window(dialog)

    if not dialog.result:
        root.destroy()
        return

    root.deiconify()
    app = CollectorApp(root, dialog.result, app_cfg)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
