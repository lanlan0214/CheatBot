import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib import request
import json


def _normalize_version(version: str) -> tuple[int, ...]:
    text = (version or "").strip().lstrip("vV")
    parts: list[int] = []
    for token in text.split("."):
        try:
            parts.append(int(token))
        except Exception:
            parts.append(0)
    return tuple(parts) if parts else (0,)


def _is_newer(latest: str, current: str) -> bool:
    latest_parts = _normalize_version(latest)
    current_parts = _normalize_version(current)

    size = max(len(latest_parts), len(current_parts))
    latest_padded = latest_parts + (0,) * (size - len(latest_parts))
    current_padded = current_parts + (0,) * (size - len(current_parts))
    return latest_padded > current_padded


def _http_get_json(url: str, timeout: int = 8) -> dict:
    req = request.Request(
        url,
        headers={"User-Agent": "desktop-automation-updater/1.0"},
        method="GET",
    )
    with request.urlopen(req, timeout=timeout) as resp:
        if resp.status != 200:
            raise RuntimeError(f"HTTP {resp.status}")
        data = resp.read().decode("utf-8", errors="replace")
    payload = json.loads(data)
    if not isinstance(payload, dict):
        raise RuntimeError("更新資訊格式錯誤")
    return payload


def _pick_github_release_asset(assets: list, preferred_name: str = "") -> str:
    if not isinstance(assets, list):
        return ""

    preferred_name = (preferred_name or "").lower().strip()
    exe_assets: list[dict] = []

    for item in assets:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "")).strip()
        url = str(item.get("browser_download_url", "")).strip()
        if not name or not url:
            continue
        if name.lower().endswith(".exe"):
            exe_assets.append(item)

    if not exe_assets:
        return ""

    if preferred_name:
        for item in exe_assets:
            name = str(item.get("name", "")).strip().lower()
            if preferred_name in name:
                return str(item.get("browser_download_url", "")).strip()

    return str(exe_assets[0].get("browser_download_url", "")).strip()


def _resolve_update_payload(payload: dict, current_version: str) -> dict:
    # 模式 A：自訂 manifest
    # {
    #   "latest_version": "1.0.1",
    #   "download_url": "https://.../app.exe",
    #   "notes": "..."
    # }
    if "latest_version" in payload:
        latest_version = str(payload.get("latest_version", "")).strip()
        download_url = str(payload.get("download_url", "")).strip()
        notes = str(payload.get("notes", "")).strip()

        if not latest_version:
            raise RuntimeError("manifest 缺少 latest_version")
        if _is_newer(latest_version, current_version) and not download_url:
            raise RuntimeError("manifest 缺少 download_url")

        return {
            "has_update": _is_newer(latest_version, current_version),
            "latest_version": latest_version,
            "download_url": download_url,
            "notes": notes,
        }

    # 模式 B：GitHub Releases latest API
    # https://api.github.com/repos/<owner>/<repo>/releases/latest
    if "tag_name" in payload:
        latest_version = str(payload.get("tag_name", "")).strip().lstrip("vV")
        notes = str(payload.get("body", "")).strip()

        preferred_name = ""
        if getattr(sys, "frozen", False):
            preferred_name = Path(sys.executable).stem

        download_url = _pick_github_release_asset(payload.get("assets", []), preferred_name)
        if not latest_version:
            raise RuntimeError("GitHub release 缺少 tag_name")
        if _is_newer(latest_version, current_version) and not download_url:
            raise RuntimeError("GitHub release 找不到可下載的 .exe 資產")

        return {
            "has_update": _is_newer(latest_version, current_version),
            "latest_version": latest_version,
            "download_url": download_url,
            "notes": notes,
        }

    raise RuntimeError("不支援的更新格式：請提供 manifest 或 GitHub release latest API")


def check_for_update(current_version: str, manifest_url: str, timeout: int = 8) -> dict:
    payload = _http_get_json(manifest_url, timeout=timeout)
    return _resolve_update_payload(payload, current_version)


def _download_file(url: str, output_path: Path, timeout: int = 20) -> None:
    req = request.Request(
        url,
        headers={"User-Agent": "desktop-automation-updater/1.0"},
        method="GET",
    )
    with request.urlopen(req, timeout=timeout) as resp, open(output_path, "wb") as out:
        if resp.status != 200:
            raise RuntimeError(f"下載更新檔失敗: HTTP {resp.status}")
        shutil.copyfileobj(resp, out)


def _build_replace_script(script_path: Path, new_exe: Path, target_exe: Path) -> None:
    # 目前執行中的 exe 無法直接覆蓋，交給外部 bat 在程式關閉後替換
    script = f"""@echo off
setlocal
set "NEW_EXE={new_exe}"
set "TARGET_EXE={target_exe}"

ping 127.0.0.1 -n 3 >nul

:retry
copy /Y "%NEW_EXE%" "%TARGET_EXE%" >nul
if errorlevel 1 (
  ping 127.0.0.1 -n 2 >nul
  goto retry
)

start "" "%TARGET_EXE%"
del /F /Q "%NEW_EXE%" >nul 2>nul
del /F /Q "%~f0" >nul 2>nul
endlocal
"""
    script_path.write_text(script, encoding="utf-8", newline="\r\n")


def prepare_and_launch_update(download_url: str) -> None:
    if not getattr(sys, "frozen", False):
        raise RuntimeError("目前不是 exe 模式，無法自動覆蓋更新")

    target_exe = Path(sys.executable)
    temp_dir = Path(tempfile.mkdtemp(prefix="desktop_auto_update_"))
    new_exe = temp_dir / "update_new.exe"
    bat_file = temp_dir / "apply_update.bat"

    _download_file(download_url, new_exe)
    _build_replace_script(bat_file, new_exe, target_exe)

    creation_flags = 0
    if hasattr(subprocess, "DETACHED_PROCESS"):
        creation_flags |= subprocess.DETACHED_PROCESS
    if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
        creation_flags |= subprocess.CREATE_NEW_PROCESS_GROUP

    subprocess.Popen(
        ["cmd", "/c", str(bat_file)],
        creationflags=creation_flags,
        close_fds=True,
    )
