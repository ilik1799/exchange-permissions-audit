"""
Запуск PowerShell через subprocess, парсинг JSON.

Особенности:
  - ExcludedOUs/Groups через временный JSON-файл
  - Декодирование: utf-8 / cp1251 / cp866
  - Прогресс из stderr в реальном времени
"""

import json
import logging
import subprocess
import tempfile
from pathlib import Path

log = logging.getLogger(__name__)

# PS-скрипт лежит рядом: scripts/check_permissions.ps1
SCRIPT_PATH = Path(__file__).parent / "scripts" / "check_permissions.ps1"


def _decode(raw_bytes: bytes) -> str:
    """Декодирование байтов из PS."""
    for enc in ("utf-8", "cp1251", "cp866"):
        try:
            return raw_bytes.decode(enc)
        except (UnicodeDecodeError, ValueError):
            continue
    return raw_bytes.decode("utf-8", errors="replace")


def run(
    target_user: str,
    login: str,
    password: str,
    exchange_server: str,
    max_threads: int,
    excluded_ous: list[str],
    excluded_groups: list[str],
    send_as_guid: str,
    debug: bool = False,
) -> dict:
    """Запуск PS, возврат dict с результатами."""
    if not SCRIPT_PATH.exists():
        return {"error": f"PS script not found: {SCRIPT_PATH}", "report": []}

    # Экранирование пароля: ' → ''
    password_escaped = password.replace("'", "''")

    # Исключения через временный JSON-файл
    # (кавычки в аргументах ломают парсинг)
    config_data = {
        "ExcludedOUs": excluded_ous,
        "ExcludedGroups": excluded_groups,
    }

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False, encoding="utf-8"
    ) as f:
        json.dump(config_data, f, ensure_ascii=False)
        config_path = f.name

    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_PATH),
        "-TargetUser",
        target_user,
        "-ExchangeServer",
        exchange_server,
        "-Login",
        login,
        "-Password",
        password_escaped,
        "-MaxThreads",
        str(max_threads),
        "-ConfigPath",
        config_path,
        "-SendAsGuid",
        send_as_guid,
    ]
    if debug:
        cmd.append("-EnableDebug")

    log.info("Запуск PowerShell...")

    import sys
    import threading

    stdout_bytes = b""

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        # Стримим stderr в реальном времени
        def _stream_stderr(pipe):
            for raw_line in pipe:
                line = _decode(raw_line).strip()
                if line:
                    msg = f"\r  {line}" + " " * 20
                    sys.stderr.write(msg)
                    sys.stderr.flush()
            sys.stderr.write("\r" + " " * 60 + "\r")
            sys.stderr.flush()

        t = threading.Thread(
            target=_stream_stderr,
            args=(proc.stderr,),
            daemon=True,
        )
        t.start()

        if proc.stdout:
            stdout_bytes = proc.stdout.read()
        proc.wait()
        t.join(timeout=5)
    finally:
        # Удаляем временный JSON-файл с конфигом
        Path(config_path).unlink(missing_ok=True)

    # Декодируем stdout и парсим JSON
    stdout_text = _decode(stdout_bytes).strip()
    if not stdout_text:
        return {
            "error": f"Empty output (exit code: {proc.returncode})",
            "report": [],
        }

    try:
        # Ищем начало JSON (PS может вывести мусор)
        json_start = stdout_text.find("{")
        if json_start >= 0:
            stdout_text = stdout_text[json_start:]
        return json.loads(stdout_text)
    except json.JSONDecodeError as e:
        log.error("JSON parse error: %s", e)
        return {"error": str(e), "report": [], "raw": stdout_text[:500]}
