from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


class EmfConvertError(RuntimeError):
    """SVG 转 EMF 失败。"""


def _prefer_windows_exe(path_value: str | os.PathLike[str] | None) -> str | None:
    """Windows 下优先使用 inkscape.exe，避免 inkscape.com 弹出控制台窗口。"""
    if not path_value:
        return None

    path = Path(path_value)
    if sys.platform.startswith("win") and path.suffix.lower() == ".com":
        exe_path = path.with_suffix(".exe")
        if exe_path.exists():
            return str(exe_path)
    return str(path) if path.exists() else None


def find_inkscape() -> str | None:
    """查找 Inkscape。Windows 下始终优先选择 GUI 版 inkscape.exe。"""
    env_path = os.environ.get("INKSCAPE_PATH") or os.environ.get("LATEX_SVG_INKSCAPE")
    env_result = _prefer_windows_exe(env_path)
    if env_result:
        return env_result

    if sys.platform.startswith("win"):
        # 不先查找裸名称 inkscape，因为 PATHEXT 通常把 .COM 排在 .EXE 前面，
        # shutil.which("inkscape") 可能返回会弹黑框的 inkscape.com。
        found_exe = shutil.which("inkscape.exe")
        if found_exe:
            return found_exe
    else:
        found = shutil.which("inkscape")
        if found:
            return found

    candidates = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Inkscape" / "bin" / "inkscape.exe",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Inkscape" / "inkscape.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Inkscape" / "bin" / "inkscape.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Inkscape" / "bin" / "inkscape.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    # 极少数环境可能只有控制台启动器；作为最后回退仍可使用，
    # subprocess 参数会再尝试隐藏它的控制台窗口。
    if sys.platform.startswith("win"):
        found_com = shutil.which("inkscape.com")
        if found_com:
            return found_com
    return None


def _hidden_subprocess_kwargs() -> dict[str, Any]:
    """构造不显示外部控制台窗口的 subprocess 参数。"""
    kwargs: dict[str, Any] = {
        "capture_output": True,
        "text": True,
        "encoding": "utf-8",
        "errors": "replace",
        "stdin": subprocess.DEVNULL,
        "timeout": 60,
    }

    if sys.platform.startswith("win"):
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs["startupinfo"] = startupinfo
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    return kwargs


def _run_inkscape(command: list[str]) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(command, **_hidden_subprocess_kwargs())
    except subprocess.TimeoutExpired as exc:
        raise EmfConvertError("Inkscape 转换超时（超过 60 秒）。") from exc
    except OSError as exc:
        raise EmfConvertError(f"无法启动 Inkscape：{exc}") from exc


def svg_to_emf(svg_path: str | Path, emf_path: str | Path) -> Path:
    svg_path = Path(svg_path).resolve()
    emf_path = Path(emf_path).resolve()
    if not svg_path.exists():
        raise EmfConvertError(f"SVG 文件不存在：{svg_path}")

    inkscape = find_inkscape()
    if not inkscape:
        raise EmfConvertError(
            "未找到 Inkscape。请安装 Inkscape，或设置环境变量 INKSCAPE_PATH 指向 inkscape.exe。"
        )

    emf_path.parent.mkdir(parents=True, exist_ok=True)
    if emf_path.exists():
        try:
            emf_path.unlink()
        except Exception:
            pass

    # Inkscape 1.x 推荐参数。
    cmd = [
        inkscape,
        str(svg_path),
        "--export-type=emf",
        f"--export-filename={emf_path}",
    ]
    proc = _run_inkscape(cmd)
    if proc.returncode != 0 or not emf_path.exists():
        # 兼容较老版本 Inkscape。
        legacy_cmd = [inkscape, str(svg_path), f"--export-emf={emf_path}"]
        legacy = _run_inkscape(legacy_cmd)
        if legacy.returncode != 0 or not emf_path.exists():
            detail = (proc.stderr or proc.stdout or legacy.stderr or legacy.stdout or "").strip()
            raise EmfConvertError("SVG 转 EMF 失败。" + (f"\n{detail}" if detail else ""))

    return emf_path
