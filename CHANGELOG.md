# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [1.0.0] - 2026-04-11

### Added
- GUI-based continuous log collector using tkinter
- Setup dialog: continue from existing `.xlsx` file or start new collection
- Main window with countdown timer, "Retrieve Now" button, and live activity log
- Automatic polling every 30 minutes (configurable via `POLL_INTERVAL_SECS`)
- Color-coded Excel rows by log category (critical → red, warning → yellow, info → green/blue)
- Watchdog thread that automatically restarts the poll worker on failure
- JST (UTC+9) timestamp conversion for all log entries
- Optional custom pattern matching via `log_patterns_example.xlsx`
- "Set Schedule" panel to override the next retrieval time manually
- Auto-resizing Excel columns (14–90 character width)
- Frozen header row in Excel output
- Thread-safe logging via `queue.Queue`
- Graceful shutdown with confirmation dialog

[Unreleased]: https://github.com/SpaceGrab/fetch_orbit_logs_to_excel/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/SpaceGrab/fetch_orbit_logs_to_excel/releases/tag/v1.0.0
