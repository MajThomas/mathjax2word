from __future__ import annotations

import atexit
import os
import socket
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Callable, Optional

ROOT = Path(__file__).resolve().parent
PORT = int(os.environ.get("LATEX_SVG_PORT", "8000"))
HOST = "127.0.0.1"
URL = f"http://localhost:{PORT}/latex-svg-clipboard.html?mini=1"
MINI_WIDTH = int(os.environ.get("LATEX_SVG_MINI_WIDTH", "1295"))
MINI_COLLAPSED_WIDTH = int(os.environ.get("LATEX_SVG_MINI_COLLAPSED_WIDTH", "730"))
MINI_HEIGHT = int(os.environ.get("LATEX_SVG_MINI_HEIGHT", "40"))
LOG_FILE = ROOT / "mini_window.log"
SINGLE_INSTANCE_MUTEX_NAME = r"Local\LatexSvgClipboardMiniWindow"
_SINGLE_INSTANCE_MUTEX_HANDLE = None

# 注意：绝不能把 pywebview Window 对象保存到 js_api 实例的公开属性中。
# pywebview 注册 JS API 时会检查对象属性；若属性引用 Window，就可能继续遍历
# window.native（WinForms/WebView2 对象树），造成循环引用、递归溢出和非 UI 线程 COM 异常。
# 因此窗口与抽屉状态只保存在模块私有状态中，MiniApi 本身保持完全无状态。
_MINI_RUNTIME: dict[str, object] = {
    "window": None,
    "drawer_collapsed": False,
    "current_width": MINI_WIDTH,
}


def _release_single_instance_lock() -> None:
    """释放 Windows 命名互斥量句柄；进程异常退出时系统也会自动释放。"""
    global _SINGLE_INSTANCE_MUTEX_HANDLE
    handle = _SINGLE_INSTANCE_MUTEX_HANDLE
    _SINGLE_INSTANCE_MUTEX_HANDLE = None
    if not handle or not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.kernel32.CloseHandle(handle)
    except Exception:
        pass


def _acquire_single_instance_lock() -> bool:
    """
    获取迷你窗口的单实例锁。

    返回 False 表示已有实例正在运行。调用方应直接结束当前（第二个）进程，
    不唤出、不隐藏，也不关闭已经运行的第一个实例。
    """
    global _SINGLE_INSTANCE_MUTEX_HANDLE

    if not sys.platform.startswith("win"):
        # 该工具的迷你窗口主要面向 Windows；其他平台不额外限制。
        return True

    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        create_mutex = kernel32.CreateMutexW
        create_mutex.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
        create_mutex.restype = wintypes.HANDLE

        handle = create_mutex(None, False, SINGLE_INSTANCE_MUTEX_NAME)
        if not handle:
            # 互斥量创建失败时不阻断程序启动，但记录错误便于排查。
            _log(f"创建单实例互斥量失败，WinError={ctypes.get_last_error()}")
            return True

        ERROR_ALREADY_EXISTS = 183
        if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
            kernel32.CloseHandle(handle)
            return False

        _SINGLE_INSTANCE_MUTEX_HANDLE = handle
        atexit.register(_release_single_instance_lock)
        return True
    except Exception:
        _log("创建单实例锁失败：\n" + traceback.format_exc())
        # 锁机制异常时仍允许启动，避免工具完全不可用。
        return True


def _log(message: str) -> None:
    """pythonw.exe 无控制台运行时，把错误写入日志，便于排查。"""
    try:
        LOG_FILE.write_text("", encoding="utf-8") if LOG_FILE.exists() and LOG_FILE.stat().st_size > 1_000_000 else None
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(time.strftime("[%Y-%m-%d %H:%M:%S] ") + message + "\n")
    except Exception:
        pass


def _safe_print(message: str) -> None:
    try:
        if sys.stdout:
            print(message)
        else:
            _log(message)
    except Exception:
        _log(message)


def _port_is_open(host: str, port: int) -> bool:
    try:
        with socket.create_connection((host, port), timeout=0.4):
            return True
    except OSError:
        return False


def _start_server_if_needed() -> None:
    if _port_is_open(HOST, PORT):
        return
    from run_local_server import Handler, ReusableTCPServer

    os.chdir(ROOT)
    httpd = ReusableTCPServer((HOST, PORT), Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()

    for _ in range(40):
        if _port_is_open(HOST, PORT):
            return
        time.sleep(0.05)
    raise RuntimeError("本地服务启动失败。")


class WinTrayIcon:
    """Windows 系统托盘图标：左键唤出窗口，右键菜单可退出后台。"""

    WM_TRAYICON = 0x0400 + 20
    ID_SHOW = 1001
    ID_EXIT = 1002

    def __init__(self, on_show: Callable[[], None], on_exit: Callable[[], None]) -> None:
        self.on_show = on_show
        self.on_exit = on_exit
        self.hwnd: Optional[int] = None
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()
        self._class_name = f"LatexSvgMiniTrayWindow_{os.getpid()}"

    def start(self) -> bool:
        if not sys.platform.startswith("win"):
            return False
        self._thread = threading.Thread(target=self._message_loop, name="LatexSvgMiniTray", daemon=True)
        self._thread.start()
        return self._ready.wait(timeout=2.0)

    def stop(self) -> None:
        try:
            if self.hwnd:
                import win32con  # type: ignore
                import win32gui  # type: ignore
                win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass

    def _message_loop(self) -> None:
        try:
            import win32api  # type: ignore
            import win32con  # type: ignore
            import win32gui  # type: ignore

            hinst = win32api.GetModuleHandle(None)
            wc = win32gui.WNDCLASS()
            wc.hInstance = hinst
            wc.lpszClassName = self._class_name
            wc.lpfnWndProc = self._wnd_proc
            try:
                win32gui.RegisterClass(wc)
            except Exception:
                pass

            self.hwnd = win32gui.CreateWindow(
                self._class_name,
                "LaTeX 公式小工具托盘",
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                hinst,
                None,
            )
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
            flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
            nid = (self.hwnd, 0, flags, self.WM_TRAYICON, hicon, "LaTeX 公式小工具")
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
            self._ready.set()
            win32gui.PumpMessages()
        except Exception:
            _log("系统托盘启动失败：\n" + traceback.format_exc())
            self._ready.set()

    def _show_popup_menu(self) -> None:
        try:
            import win32api  # type: ignore
            import win32con  # type: ignore
            import win32gui  # type: ignore

            menu = win32gui.CreatePopupMenu()
            win32gui.AppendMenu(menu, win32con.MF_STRING, self.ID_SHOW, "显示迷你窗口")
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
            win32gui.AppendMenu(menu, win32con.MF_STRING, self.ID_EXIT, "退出后台")
            x, y = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(self.hwnd)
            win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, x, y, 0, self.hwnd, None)
            win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
        except Exception:
            _log("显示托盘菜单失败：\n" + traceback.format_exc())

    def _delete_icon(self) -> None:
        try:
            if self.hwnd:
                import win32gui  # type: ignore
                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self.hwnd, 0))
        except Exception:
            pass

    def _wnd_proc(self, hwnd, msg, wparam, lparam):  # noqa: ANN001
        import win32con  # type: ignore
        import win32gui  # type: ignore

        if msg == self.WM_TRAYICON:
            if lparam in (win32con.WM_LBUTTONUP, win32con.WM_LBUTTONDBLCLK):
                self.on_show()
                return 0
            if lparam == win32con.WM_RBUTTONUP:
                self._show_popup_menu()
                return 0
        elif msg == win32con.WM_COMMAND:
            command_id = int(wparam) & 0xFFFF
            if command_id == self.ID_SHOW:
                self.on_show()
                return 0
            if command_id == self.ID_EXIT:
                self.on_exit()
                return 0
        elif msg == win32con.WM_CLOSE:
            self._delete_icon()
            try:
                win32gui.DestroyWindow(hwnd)
            except Exception:
                pass
            return 0
        elif msg == win32con.WM_DESTROY:
            self._delete_icon()
            try:
                win32gui.PostQuitMessage(0)
            except Exception:
                pass
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)


class MiniApi:
    """仅向 JavaScript 暴露基础类型接口；实例本身不持有 Window/COM/.NET 对象。"""

    __slots__ = ()

    def show(self) -> bool:
        window = _MINI_RUNTIME.get("window")
        width = int(_MINI_RUNTIME.get("current_width") or MINI_WIDTH)
        return _show_window(window, width)

    def set_drawer_collapsed(self, collapsed=False) -> dict[str, object]:
        if isinstance(collapsed, str):
            collapsed = collapsed.strip().lower() in ("1", "true", "yes", "on")

        is_collapsed = bool(collapsed)
        width = MINI_COLLAPSED_WIDTH if is_collapsed else MINI_WIDTH
        _MINI_RUNTIME["drawer_collapsed"] = is_collapsed
        _MINI_RUNTIME["current_width"] = width

        try:
            window = _MINI_RUNTIME.get("window")
            if window is not None:
                window.resize(width, MINI_HEIGHT)
            return {"ok": True, "collapsed": is_collapsed, "width": width}
        except Exception:
            _log("调整抽屉窗口宽度失败：\n" + traceback.format_exc())
            return {"ok": False, "collapsed": is_collapsed, "width": width}

    def hide(self) -> bool:
        return _hide_window(_MINI_RUNTIME.get("window"))

    def close(self) -> bool:
        # 关闭按钮只隐藏窗口，不结束后台。需要完全退出时，从托盘图标右键选择“退出后台”。
        return self.hide()


def _hide_window(window) -> bool:  # noqa: ANN001
    if window is None:
        return False
    for method_name in ("hide", "minimize"):
        try:
            method = getattr(window, method_name, None)
            if callable(method):
                method()
                return True
        except Exception:
            pass
    return False


def _show_window(window, width: int = MINI_WIDTH) -> bool:  # noqa: ANN001
    if window is None:
        return False
    for method_name in ("show", "restore"):
        try:
            method = getattr(window, method_name, None)
            if callable(method):
                method()
                break
        except Exception:
            pass
    try:
        window.on_top = True
    except Exception:
        pass
    try:
        window.resize(int(width), MINI_HEIGHT)
    except Exception:
        pass
    try:
        window.evaluate_js(
            "(document.body.classList.contains('mini-drawer-collapsed')"
            " ? document.getElementById('miniDrawerToggle')"
            " : document.getElementById('miniTex'))?.focus();"
        )
    except Exception:
        pass
    return True


def _exit_process(window, tray: Optional[WinTrayIcon]) -> None:  # noqa: ANN001
    try:
        if tray:
            tray.stop()
    except Exception:
        pass
    try:
        if window is not None:
            window.destroy()
    except Exception:
        pass
    os._exit(0)


def main() -> int:
    # 必须在导入 pywebview、启动本地服务和创建托盘之前获取锁。
    # 检测到已有实例时，仅结束本次启动的第二个进程，不影响首个实例。
    if not _acquire_single_instance_lock():
        _safe_print("LaTeX 公式迷你窗口已经在运行，本次启动已退出。")
        return 0

    try:
        import webview  # type: ignore
    except Exception:
        _safe_print("缺少 pywebview。请先运行：python -m pip install -r requirements.txt")
        _safe_print("或单独运行：python -m pip install pywebview")
        return 1

    try:
        _start_server_if_needed()
    except Exception as exc:
        _log("启动本地服务失败：\n" + traceback.format_exc())
        _safe_print(f"启动本地服务失败：{exc}")
        return 1

    api = MiniApi()
    tray_ref: dict[str, Optional[WinTrayIcon]] = {"tray": None}

    def show_from_tray() -> None:
        api.show()

    def exit_from_tray() -> None:
        _exit_process(_MINI_RUNTIME.get("window"), tray_ref["tray"])

    kwargs = dict(
        title="LaTeX 公式",
        url=URL,
        width=MINI_WIDTH,
        height=MINI_HEIGHT,
        min_size=(MINI_COLLAPSED_WIDTH, MINI_HEIGHT),
        resizable=False,
        frameless=True,
        on_top=True,
        js_api=api,
        background_color="#111827",
    )
    # easy_drag=False 禁止整窗任意位置拖动；HTML 中仅左右把手使用
    # pywebview-drag-region + -webkit-app-region: drag。
    try:
        window = webview.create_window(easy_drag=False, **kwargs)
    except TypeError:
        # 兼容旧版本 pywebview：去掉不支持的 min_size/easy_drag 参数。
        kwargs.pop("min_size", None)
        try:
            window = webview.create_window(easy_drag=False, **kwargs)
        except TypeError:
            window = webview.create_window(**kwargs)

    # Window 仅保存到模块私有状态，禁止挂到 js_api 实例上。
    _MINI_RUNTIME["window"] = window
    _MINI_RUNTIME["drawer_collapsed"] = False
    _MINI_RUNTIME["current_width"] = MINI_WIDTH

    tray = WinTrayIcon(on_show=show_from_tray, on_exit=exit_from_tray)
    tray_ref["tray"] = tray

    def on_started() -> None:
        try:
            if sys.platform.startswith("win"):
                tray.start()
        except Exception:
            _log("系统托盘初始化失败：\n" + traceback.format_exc())


    _safe_print(f"迷你窗口已启动：{URL}")
    try:
        webview.start(on_started)
    except TypeError:
        webview.start()
    finally:
        try:
            tray.stop()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:
        _log("迷你窗口异常退出：\n" + traceback.format_exc())
        raise
