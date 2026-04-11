# fetch_orbit_logs_to_excel

A GUI-based continuous log collector for **Boston Dynamics Spot** robots that fetches activity logs from a [Spot Orbit](https://www.bostondynamics.com/orbit) management system and saves them to color-coded Excel (`.xlsx`) files.

---

## Features

- **Continuous polling** — Automatically retrieves new logs on a configurable schedule (default: every 30 minutes)
- **Color-coded Excel output** — Rows are highlighted by severity (critical → red, warning → yellow, informational → green/blue)
- **Resume support** — Continue appending to an existing `.xlsx` file or start a fresh collection from a custom date/time
- **Manual retrieval** — "Retrieve Now" button for on-demand log fetching outside the schedule
- **Watchdog thread** — Automatically restarts the poll worker if it crashes
- **Custom pattern matching** — Load regex-based log categorization rules from an optional `log_patterns_example.xlsx` file
- **JST timestamps** — All timestamps are displayed in Japan Standard Time (UTC+9)

---

## Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.7+ |
| Boston Dynamics Orbit SDK | latest |
| openpyxl | 3.0+ |
| tkinter | (bundled with Python) |

You must have network access to a running **Spot Orbit** server.

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/SpaceGrab/fetch_orbit_logs_to_excel.git
cd fetch_orbit_logs_to_excel

# 2. (Recommended) Create and activate a virtual environment
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

---

## Configuration

Open `fetch_orbit_logs_continuous.py` and update the constants at the top of the file:

```python
# ── Required ──────────────────────────────────────────────────
ORBIT_HOST      = "172.24.43.164"           # IP / hostname of your Orbit server
ORBIT_API_TOKEN = "<your-api-token-here>"   # Orbit API token

# ── Optional ──────────────────────────────────────────────────
TLS_VERIFY         = False          # Set True to verify the server's TLS certificate
POLL_INTERVAL_SECS = 30 * 60       # Polling interval in seconds (default: 30 min)
BOOTSTRAP_MINUTES  = 30            # Look-back window when no prior data exists
```

> **Security note:** Never commit a real API token to source control. See [SECURITY.md](SECURITY.md) for best practices.

### Optional: Custom Log Pattern File

Place a file named `log_patterns_example.xlsx` in the working directory. It must contain a sheet called **"Log Patterns"** with regex patterns and their assigned categories. If the file is absent, the tool falls back to built-in pattern matching.

---

## Usage

```bash
python fetch_orbit_logs_continuous.py
```

### Startup Dialog

On first launch a setup dialog appears with two options:

| Option | Description |
|---|---|
| **Continue from existing file** | Select an existing `.xlsx` file that already has an "Orbit Logs" sheet. New rows are appended after the last entry. |
| **Start new collection** | Choose a start date/time and a save path for a brand-new `.xlsx` file. |

### Main Window

Once running, the main window shows:

- **Excel file path** currently being written to
- **Last retrieved** timestamp
- **Next retrieval** countdown timer
- **Retrieve Now** — triggers an immediate fetch
- **Set Schedule** — override the next scheduled retrieval time
- **Activity log** — live feed of connection events, fetch results, and errors
- **Stop Collector** — graceful shutdown with confirmation prompt

---

## Excel Output Format

| Column | Description |
|---|---|
| Robot Serial Number | Spot robot identifier |
| Date and time | Timestamp in JST (`YYYY/MM/DD HH:MM:SS`) |
| Route | Mission / route name |
| Orbit log | Raw log message |

### Color Coding

| Color | Category |
|---|---|
| Red `#FF4C4C` | Software / System Issues |
| Orange `#FF8042` | Navigation — Route Conflict |
| Dark Orange `#FF9900` | Navigation — Stuck (Goal Blocked) |
| Light Orange `#FFB347` | Navigation — Stuck |
| Yellow `#FFD966` | Navigation — Path Not Found |
| Light Yellow `#FFDC73` | Entity Detection |
| Green `#C6EFCE` | Task Completion |
| Blue `#BDD7EE` | Docking / Undocking |
| Light Blue `#DDEEFF` | Run Lifecycle |
| Light Grey `#F2F2F2` | Task Skipped |
| Lavender `#F0EBFF` | Mission Interruption |
| Cream `#FEFCE8` | Battery |

---

## Project Structure

```
fetch_orbit_logs_to_excel/
├── fetch_orbit_logs_continuous.py   # Main application
├── log_patterns_example.xlsx        # (Optional) Custom pattern rules
├── requirements.txt
├── CHANGELOG.md
├── CONTRIBUTING.md
├── LICENSE
├── SECURITY.md
└── README.md
```

---

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) before submitting a pull request.

---

## Security

This project contains network credentials (API token, host address). Please review [SECURITY.md](SECURITY.md) for responsible handling of credentials and for reporting security vulnerabilities.

---

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.

---

## Acknowledgements

- [Boston Dynamics Spot SDK](https://github.com/boston-dynamics/spot-sdk) for the Orbit client library
- [openpyxl](https://openpyxl.readthedocs.io/) for Excel file handling
