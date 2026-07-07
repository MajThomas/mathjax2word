from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path


class EmfConvertError(RuntimeError):
    """SVG 转 EMF 失败。"""


def find_inkscape() -> str | None:
    """查找 Inkscape 可执行文件。优先级：环境变量 → PATH → 常见安装目录。"""
    env_path = os.environ.get("INKSCAPE_PATH") or os.environ.get("LATEX_SVG_INKSCAPE")
    if env_path and Path(env_path).exists():
        return env_path

    found = shutil.which("inkscape") or shutil.which("inkscape.exe")
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
    return None


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
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    if proc.returncode != 0 or not emf_path.exists():
        # 兼容较老版本 Inkscape。
        legacy_cmd = [inkscape, str(svg_path), f"--export-emf={emf_path}"]
        legacy = subprocess.run(legacy_cmd, capture_output=True, text=True, timeout=60)
        if legacy.returncode != 0 or not emf_path.exists():
            detail = (proc.stderr or proc.stdout or legacy.stderr or legacy.stdout or "").strip()
            raise EmfConvertError("SVG 转 EMF 失败。" + (f"\n{detail}" if detail else ""))

    return emf_path
