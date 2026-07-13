from __future__ import annotations

import http.server
import json
import os
import socketserver
import webbrowser
from pathlib import Path
from typing import Any

PORT = int(os.environ.get("LATEX_SVG_PORT", "8000"))
ROOT = Path(__file__).resolve().parent
URL = f"http://localhost:{PORT}/latex-svg-clipboard.html"

FONT_LABELS = {
    "mathjax-newcm": "MathJax New Computer Modern（内置）",
    "mathjax-stix2": "MathJax STIX Two",
    "mathjax-tex": "MathJax TeX",
    "mathjax-bonum": "MathJax Bonum",
    "mathjax-pagella": "MathJax Pagella",
    "mathjax-termes": "MathJax Termes",
    "mathjax-dejavu": "MathJax DejaVu",
    "mathjax-schola": "MathJax Schola",
    "mathjax-asana": "MathJax Asana",
}


def load_manifest_fonts() -> list[dict]:
    manifest = ROOT / "fonts-manifest.json"
    if not manifest.exists():
        return []
    try:
        data = json.loads(manifest.read_text(encoding="utf-8"))
    except Exception:
        return []
    return data if isinstance(data, list) else []


def scan_vendor_fonts() -> list[dict]:
    """扫描 ./vendor/fonts 下额外放入的 @mathjax/mathjax-*-font 包。"""
    fonts: list[dict] = []
    font_root = ROOT / "vendor" / "fonts"
    if not font_root.exists():
        return fonts
    for path in sorted(font_root.iterdir()):
        if not path.is_dir():
            continue
        name = path.name
        if name.endswith("-font") and name.startswith("mathjax-"):
            font_id = name[:-5]
        else:
            font_id = name
        package_json = path / "package.json"
        label = FONT_LABELS.get(font_id, font_id)
        if package_json.exists():
            try:
                pkg = json.loads(package_json.read_text(encoding="utf-8"))
                label = pkg.get("name", label).replace("@mathjax/", "")
                label = FONT_LABELS.get(font_id, label)
            except Exception:
                pass
        fonts.append({
            "id": font_id,
            "label": label,
            "path": f"./vendor/fonts/{path.name}",
            "builtin": False,
        })
    return fonts


def available_fonts() -> list[dict]:
    seen = set()
    result: list[dict] = []
    for font in load_manifest_fonts() + scan_vendor_fonts():
        font_id = font.get("id")
        if not font_id or font_id in seen:
            continue
        seen.add(font_id)
        result.append(font)
    if "mathjax-newcm" not in seen:
        result.insert(0, {
            "id": "mathjax-newcm",
            "label": "MathJax New Computer Modern（内置）",
            "builtin": True,
        })
    return result


def _json_response(handler: http.server.BaseHTTPRequestHandler, status: int, payload: dict[str, Any]) -> None:
    data = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    handler.wfile.write(data)


class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # 页面/API 始终读取最新版本。MathJax 主程序允许缓存；字体动态分片
        # 使用 no-store，避免此前错误路径产生的 404 被 WebView2 缓存。
        path = self.path.split("?", 1)[0]
        if path.startswith("/vendor/mathjax/"):
            self.send_header("Cache-Control", "public, max-age=86400")
        else:
            self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self):  # noqa: N802 - stdlib method name
        path = self.path.split("?", 1)[0]
        if path == "/local-fonts.js":
            payload = "window.LATEX_SVG_FONTS = " + json.dumps(available_fonts(), ensure_ascii=False, indent=2) + ";\n"
            data = payload.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/javascript; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if path == "/api/check-environment":
            try:
                from formula_word import check_environment
                _json_response(self, 200, check_environment())
            except Exception as exc:
                _json_response(self, 500, {"ok": False, "error": str(exc)})
            return
        super().do_GET()

    def do_POST(self):  # noqa: N802 - stdlib method name
        path = self.path.split("?", 1)[0]
        if not path.startswith("/api/"):
            _json_response(self, 404, {"ok": False, "error": "接口不存在。"})
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length) if length > 0 else b"{}"
            body = json.loads(raw.decode("utf-8")) if raw else {}
            if not isinstance(body, dict):
                raise ValueError("请求体必须是 JSON 对象。")
        except Exception as exc:
            _json_response(self, 400, {"ok": False, "error": f"JSON 解析失败：{exc}"})
            return

        try:
            from formula_word import (
                convert_word_markers,
                find_word_conversion_targets,
                insert_formula,
                insert_numbered_formula,
                read_selected_formula,
                update_selected_formula,
                writeback_selected_formulas,
            )
            if path == "/api/insert-word":
                result = insert_formula(body)
            elif path == "/api/insert-numbered-word":
                result = insert_numbered_formula(body)
            elif path == "/api/read-word":
                result = read_selected_formula()
            elif path == "/api/find-word-conversion-targets":
                result = find_word_conversion_targets()
            elif path == "/api/convert-word-markers":
                result = convert_word_markers(body)
            elif path == "/api/writeback-word-formulas":
                result = writeback_selected_formulas()
            elif path == "/api/update-word":
                result = update_selected_formula(body)
            else:
                _json_response(self, 404, {"ok": False, "error": "接口不存在。"})
                return
            _json_response(self, 200, result)
        except Exception as exc:
            _json_response(self, 500, {"ok": False, "error": str(exc)})


class ReusableTCPServer(socketserver.ThreadingTCPServer):
    # MathJax 会并行请求主组件、扩展和字体。线程服务器可避免单个请求占住
    # 服务端后，其余组件长期排队等待。
    allow_reuse_address = True
    daemon_threads = True
    request_queue_size = 64


if __name__ == "__main__":
    os.chdir(ROOT)
    with ReusableTCPServer(("127.0.0.1", PORT), Handler) as httpd:
        print(f"服务已启动：{URL}")
        print("Word 增强功能需要 Windows + Word 桌面版 + pywin32 + Inkscape。")
        print("按 Ctrl+C 退出。")
        if os.environ.get("LATEX_SVG_NO_OPEN") != "1":
            try:
                webbrowser.open(URL)
            except Exception:
                pass
        httpd.serve_forever()
