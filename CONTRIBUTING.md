# Contributing to fetch_orbit_logs_to_excel

Thank you for your interest in contributing! This document outlines the process for reporting bugs, requesting features, and submitting pull requests.

---

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Reporting Bugs](#reporting-bugs)
- [Requesting Features](#requesting-features)
- [Development Setup](#development-setup)
- [Pull Request Process](#pull-request-process)
- [Coding Standards](#coding-standards)

---

## Code of Conduct

Please be respectful and constructive in all interactions. We follow the [Contributor Covenant](https://www.contributor-covenant.org/) code of conduct.

---

## Reporting Bugs

Before filing a bug report, please search the [existing issues](https://github.com/SpaceGrab/fetch_orbit_logs_to_excel/issues) to avoid duplicates.

When filing a new bug, use the **Bug Report** issue template and include:

- Your OS and Python version
- Steps to reproduce the issue
- Expected vs. actual behavior
- Relevant log output from the activity log panel

> **Important:** Never include API tokens, IP addresses, or other credentials in bug reports. Redact them before posting.

---

## Requesting Features

Use the **Feature Request** issue template and describe:

- The problem you are trying to solve
- Your proposed solution
- Any alternatives you have considered

---

## Development Setup

```bash
# Clone your fork
git clone https://github.com/<your-username>/fetch_orbit_logs_to_excel.git
cd fetch_orbit_logs_to_excel

# Create a virtual environment
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Install development tools (optional but recommended)
pip install ruff mypy
```

---

## Pull Request Process

1. **Fork** the repository and create a branch from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. Make your changes. Keep commits focused and atomic.

3. **Do not commit credentials.** Ensure `ORBIT_HOST` and `ORBIT_API_TOKEN` contain only placeholder values before committing.

4. Run a quick lint check before opening a PR:
   ```bash
   ruff check fetch_orbit_logs_continuous.py
   ```

5. Open a Pull Request against `main`. Fill in the PR template, describing:
   - What the change does
   - How you tested it
   - Any relevant screenshots or output

6. A maintainer will review your PR. Address any requested changes and push new commits to the same branch.

---

## Coding Standards

- Follow [PEP 8](https://peps.python.org/pep-0008/) style guidelines.
- Use type hints where practical.
- Keep the single-file architecture unless there is a compelling reason to split.
- All user-visible strings should be in English. Japanese translations are welcome as inline comments.
- Do not add new hardcoded credentials or IP addresses.
