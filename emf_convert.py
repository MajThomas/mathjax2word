from __future__ import annotations

import os
import shutil
import struct
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence


class EmfConvertError(RuntimeError):
    """SVG 直接转换为 EMF 失败。"""


ProgressCallback = Callable[[int, int], None]


def _prefer_windows_exe(path_value: str | os.PathLike[str] | None) -> str | None:
    if not path_value:
        return None
    path = Path(path_value)
    if not path.exists():
        return None
    if sys.platform.startswith("win") and path.suffix.lower() == ".exe":
        com_path = path.with_suffix(".com")
        if com_path.exists():
            return str(com_path)
    return str(path)


def find_inkscape() -> str | None:
    env_path = os.environ.get("INKSCAPE_PATH") or os.environ.get("LATEX_SVG_INKSCAPE")
    env_result = _prefer_windows_exe(env_path)
    if env_result:
        return env_result

    names = ("inkscape.com", "inkscape.exe") if sys.platform.startswith("win") else ("inkscape",)
    for name in names:
        found = shutil.which(name)
        if found:
            return found

    install_dirs = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Inkscape" / "bin",
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Inkscape",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Inkscape" / "bin",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Inkscape" / "bin",
    ]
    for directory in install_dirs:
        for name in names:
            candidate = directory / name
            if candidate.exists():
                return str(candidate)
    return None


def _hidden_popen_kwargs() -> dict[str, Any]:
    kwargs: dict[str, Any] = {
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
        "text": True,
        "encoding": "utf-8",
        "errors": "replace",
        "stdin": subprocess.DEVNULL,
    }
    if sys.platform.startswith("win"):
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs["startupinfo"] = startupinfo
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return kwargs


def _kill_process_tree(proc: subprocess.Popen[str]) -> None:
    if proc.poll() is not None:
        return
    if sys.platform.startswith("win"):
        try:
            subprocess.run(
                ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            pass
    try:
        proc.kill()
    except Exception:
        pass


def _run_inkscape(command: list[str], *, timeout: float) -> subprocess.CompletedProcess[str]:
    try:
        proc = subprocess.Popen(command, **_hidden_popen_kwargs())
    except OSError as exc:
        raise EmfConvertError(f"无法启动 Inkscape：{exc}") from exc
    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _kill_process_tree(proc)
        raise EmfConvertError(f"Inkscape 转换超时（超过 {timeout:g} 秒），已终止进程。") from exc
    return subprocess.CompletedProcess(command, int(proc.returncode or 0), stdout, stderr)


def _single_timeout_seconds() -> float:
    try:
        value = float(os.environ.get("LATEX_SVG_INKSCAPE_TIMEOUT", "30"))
    except Exception:
        value = 30.0
    return max(5.0, min(300.0, value))




def _batch_timeout_seconds(count: int) -> float:
    """单个 Inkscape shell 进程的批量超时。"""
    raw = os.environ.get("LATEX_SVG_INKSCAPE_BATCH_TIMEOUT")
    if raw:
        try:
            return max(15.0, min(900.0, float(raw)))
        except Exception:
            pass
    return max(30.0, min(600.0, 15.0 + max(1, count) * 5.0))


def _batch_chunks(
    pairs: Sequence[tuple[Path, Path]],
    *,
    max_files: int = 80,
) -> list[list[tuple[Path, Path]]]:
    """限制单个 shell 会话的文件数，避免超大批次长期占用一个进程。"""
    return [list(pairs[index:index + max_files]) for index in range(0, len(pairs), max_files)]


def _run_inkscape_shell(
    inkscape: str,
    *,
    workdir: Path,
    commands: str,
    timeout: float,
) -> subprocess.CompletedProcess[str]:
    """启动一次 Inkscape shell，并在该进程内连续转换多个文件。"""
    kwargs = _hidden_popen_kwargs()
    kwargs["stdin"] = subprocess.PIPE
    kwargs["cwd"] = str(workdir)
    try:
        proc = subprocess.Popen([inkscape, "--shell"], **kwargs)
    except OSError as exc:
        raise EmfConvertError(f"无法启动 Inkscape 批处理：{exc}") from exc
    try:
        stdout, stderr = proc.communicate(commands, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _kill_process_tree(proc)
        raise EmfConvertError(
            f"Inkscape 批量转换超时（超过 {timeout:g} 秒），已终止进程。"
        ) from exc
    return subprocess.CompletedProcess(
        [inkscape, "--shell"],
        int(proc.returncode or 0),
        stdout,
        stderr,
    )


def _shell_action_line(svg_name: str, emf_name: str) -> str:
    # 临时目录中只使用短 ASCII 文件名，因此不需要处理 shell action 的
    # 分号、冒号、引号和 Unicode 路径转义。
    return (
        f"file-open:{svg_name}; "
        "export-area-page; "
        "export-type:emf; "
        f"export-filename:{emf_name}; "
        "export-overwrite; export-do; file-close"
    )


def _valid_emf(path: Path) -> bool:
    try:
        if not path.exists() or path.stat().st_size < 88:
            return False
        data = path.read_bytes()[:44]
        return len(data) >= 44 and struct.unpack_from("<I", data, 40)[0] == 0x464D4520
    except Exception:
        return False


def _clean(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except Exception:
        pass


def _export_one(svg: Path, emf: Path, *, index: int | None = None, total: int | None = None) -> Path:
    inkscape = find_inkscape()
    if not inkscape:
        raise EmfConvertError("未找到 Inkscape。请安装 Inkscape，或设置环境变量 INKSCAPE_PATH。")

    emf.parent.mkdir(parents=True, exist_ok=True)
    part = emf.parent / f"direct_{os.getpid()}_{index or 0:04d}.part.emf"
    _clean(part)
    _clean(emf)

    command = [
        inkscape,
        "--export-area-page",
        "--export-overwrite",
        "--export-type=emf",
        f"--export-filename={part}",
        str(svg),
    ]
    result = _run_inkscape(command, timeout=_single_timeout_seconds())
    if result.returncode == 0 and _valid_emf(part):
        part.replace(emf)
        return emf

    detail = str(result.stderr or result.stdout or "").strip()
    _clean(part)
    prefix = f"第 {index}/{total} 个公式" if index is not None and total is not None else "SVG"
    raise EmfConvertError(
        f"{prefix}直接转换 EMF 失败：{svg.name}" + (f"\n{detail}" if detail else "")
    )


def svg_to_emf(svg_path: str | Path, emf_path: str | Path) -> Path:
    svg = Path(svg_path).resolve()
    emf = Path(emf_path).resolve()
    if not svg.exists():
        raise EmfConvertError(f"SVG 文件不存在：{svg}")
    return _export_one(svg, emf)


def svg_batch_to_emf(
    pairs: Iterable[tuple[str | Path, str | Path]],
    *,
    progress_callback: ProgressCallback | None = None,
) -> list[Path]:
    """
    在尽量少的 Inkscape 进程中批量转换 SVG。

    每个批次先复制到独立临时目录，并改为短 ASCII 文件名；随后通过
    Inkscape ``--shell`` 在同一进程内逐个显式指定输出文件名。这样既避免
    多次启动 Inkscape，也不依赖多文件模式的派生文件名。仅对缺失项执行
    单文件回退。
    """
    normalized: list[tuple[Path, Path]] = []
    for svg_value, emf_value in pairs:
        svg = Path(svg_value).resolve()
        emf = Path(emf_value).resolve()
        if not svg.exists():
            raise EmfConvertError(f"SVG 文件不存在：{svg}")
        emf.parent.mkdir(parents=True, exist_ok=True)
        _clean(emf)
        normalized.append((svg, emf))

    total = len(normalized)
    if total == 0:
        return []

    inkscape = find_inkscape()
    if not inkscape:
        raise EmfConvertError("未找到 Inkscape。请安装 Inkscape，或设置环境变量 INKSCAPE_PATH。")

    completed = 0
    failures: list[str] = []
    for chunk_index, chunk in enumerate(_batch_chunks(normalized), start=1):
        stage_dir = Path(tempfile.mkdtemp(prefix=f"latex_svg_emf_{os.getpid()}_{chunk_index}_"))
        stage_items: list[tuple[Path, Path, Path, Path]] = []
        try:
            action_lines: list[str] = []
            for local_index, (source_svg, desired_emf) in enumerate(chunk, start=1):
                stem = f"f{local_index:04d}"
                staged_svg = stage_dir / f"{stem}.svg"
                staged_emf = stage_dir / f"{stem}.emf"
                shutil.copyfile(source_svg, staged_svg)
                stage_items.append((source_svg, desired_emf, staged_svg, staged_emf))
                action_lines.append(_shell_action_line(staged_svg.name, staged_emf.name))

            # quit 必须单独一行；每个公式一行 action，便于 Inkscape shell
            # 完成当前导出后再打开下一个文件。
            shell_input = "\n".join(action_lines + ["quit", ""])
            try:
                result = _run_inkscape_shell(
                    inkscape,
                    workdir=stage_dir,
                    commands=shell_input,
                    timeout=_batch_timeout_seconds(len(chunk)),
                )
                if result.returncode != 0:
                    detail = str(result.stderr or result.stdout or "").strip()
                    failures.append(
                        f"第 {chunk_index} 批 Inkscape shell 退出码 {result.returncode}"
                        + (f"：{detail[-1200:]}" if detail else "")
                    )
            except Exception as exc:
                failures.append(f"第 {chunk_index} 批：{exc}")

            missing: list[tuple[Path, Path]] = []
            for source_svg, desired_emf, _staged_svg, staged_emf in stage_items:
                if _valid_emf(staged_emf):
                    _clean(desired_emf)
                    shutil.move(str(staged_emf), str(desired_emf))
                    completed += 1
                    if progress_callback:
                        progress_callback(completed, total)
                else:
                    missing.append((source_svg, desired_emf))

            # shell 会话中仅有异常项才重新启动 Inkscape。
            for source_svg, desired_emf in missing:
                try:
                    _export_one(source_svg, desired_emf, index=completed + 1, total=total)
                except Exception as exc:
                    failures.append(f"{source_svg.name}：{exc}")
                    continue
                completed += 1
                if progress_callback:
                    progress_callback(completed, total)
        finally:
            shutil.rmtree(stage_dir, ignore_errors=True)

    invalid = [str(emf) for _svg, emf in normalized if not _valid_emf(emf)]
    if invalid:
        detail = "\n".join(failures[-8:])
        raise EmfConvertError(
            f"批量 SVG 转 EMF 未完成，缺失 {len(invalid)} 个输出：{', '.join(invalid[:3])}"
            + (f"\n{detail}" if detail else "")
        )
    return [emf for _svg, emf in normalized]

