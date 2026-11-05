"""PowerShell/Outlook-COM bridge for the hermes-outlook plugin.

The plugin shells out to Windows PowerShell (``powershell.exe``), which talks to
classic Outlook through its COM automation interface (``Outlook.Application``).
A JSON request is written to a temp file, the PowerShell script is run with
``-File``, and it writes a JSON result file back. File-based I/O avoids every
encoding problem that redirected PowerShell stdout would otherwise cause
(PowerShell 5.1 uses the OEM code page on redirected output).
"""

from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path
from typing import Any, Dict, Optional

_SCRIPT_DIR = Path(__file__).resolve().parent.parent / "scripts"

# Module-level settings, set once at register() time from plugin config.
_POWERSHELL: str = "powershell.exe"
_TIMEOUT: int = 120


def configure(powershell_path: str = "", timeout_seconds: int = 120) -> None:
    """Set the PowerShell executable and per-call timeout (called at register time)."""
    global _POWERSHELL, _TIMEOUT
    if powershell_path:
        _POWERSHELL = powershell_path
    if timeout_seconds:
        _TIMEOUT = int(timeout_seconds)


def run(action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    """Run one Outlook action and return the parsed JSON result dict."""
    script = _SCRIPT_DIR / "outlook_com.ps1"
    if not script.exists():
        return {"ok": False, "error": f"PowerShell bridge missing: {script}"}

    tmp = Path(tempfile.gettempdir()) / f"hermes-outlook-{uuid.uuid4().hex}"
    tmp.mkdir(parents=True, exist_ok=True)
    req_path = tmp / "request.json"
    res_path = tmp / "result.json"
    req_path.write_text(
        json.dumps({"action": action, "params": params or {}}, ensure_ascii=False),
        encoding="utf-8",
    )

    try:
        proc = subprocess.run(
            [
                _POWERSHELL,
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-RequestPath",
                str(req_path),
                "-ResultPath",
                str(res_path),
            ],
            capture_output=True,
            text=True,
            timeout=_TIMEOUT,
            encoding="utf-8",
            errors="replace",
        )

        if not res_path.exists():
            return {
                "ok": False,
                "error": (
                    f"PowerShell produced no result (exit {proc.returncode}). "
                    f"stdout={proc.stdout[-400:].strip()!r} stderr={proc.stderr[-400:].strip()!r}"
                ),
            }

        raw = res_path.read_text(encoding="utf-8-sig")
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            return {"ok": False, "error": f"Could not parse PowerShell output: {exc}. raw={raw[:400]!r}"}

        if not isinstance(data, dict):
            data = {"ok": True, "data": data}
        return data

    except subprocess.TimeoutExpired:
        return {"ok": False, "error": f"PowerShell timed out after {_TIMEOUT}s (scan too large?)."}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


__all__ = ["configure", "run"]
