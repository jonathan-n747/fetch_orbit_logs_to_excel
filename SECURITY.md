# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| 1.x     | Yes       |

---

## Credential Handling

This tool requires an **Orbit API token** and a **host IP address** to connect to a Boston Dynamics Spot Orbit server. The source file ships with placeholder/example values.

### Before committing to any repository

Replace or remove all real credentials from the configuration section:

```python
ORBIT_HOST      = "172.24.43.164"           # <-- replace with your host or use env var
ORBIT_API_TOKEN = "98bbd976-..."            # <-- replace with your token or use env var
```

### Recommended: use environment variables

Instead of hardcoding credentials, read them from environment variables:

```python
import os

ORBIT_HOST      = os.environ["ORBIT_HOST"]
ORBIT_API_TOKEN = os.environ["ORBIT_API_TOKEN"]
```

Set these before running:

```bash
# Windows (PowerShell)
$env:ORBIT_HOST      = "your-orbit-host"
$env:ORBIT_API_TOKEN = "your-api-token"

# macOS / Linux
export ORBIT_HOST="your-orbit-host"
export ORBIT_API_TOKEN="your-api-token"
```

### TLS verification

`TLS_VERIFY = False` disables certificate verification. This is acceptable on isolated internal networks but should be set to `True` (or a path to a CA bundle) in environments where TLS authenticity matters.

---

## Reporting a Vulnerability

If you discover a security vulnerability in this project, please **do not open a public GitHub issue**.

Instead, report it privately by emailing the repository owner via the contact information on their [GitHub profile](https://github.com/SpaceGrab) or by using GitHub's private vulnerability reporting feature:

1. Go to the repository on GitHub
2. Click **Security** → **Report a vulnerability**
3. Fill in the details

We aim to respond within **7 business days** and will coordinate a fix and disclosure timeline with you.

---

## Scope

The following are considered in-scope for security reports:

- Credential leakage or insecure storage
- Remote code execution via crafted log data
- Dependency vulnerabilities with a realistic exploit path

The following are out-of-scope:

- Vulnerabilities that require physical access to the Orbit server
- Issues already tracked in public CVE databases with no project-specific impact
