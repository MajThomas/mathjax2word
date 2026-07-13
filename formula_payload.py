from __future__ import annotations

import base64
import json
import zlib
from typing import Any

# Word 与 WPS 统一的新格式：不压缩，便于跨端读取和排查。
PREFIX = "LATEX_SVG_FORMULA_JSON_V1:"
# 兼容旧 Word 版已插入公式。
LEGACY_PREFIX = "LATEX_SVG_FORMULA_V1:"
TYPE_NAME = "latex-svg-formula"
VERSION = 1


class FormulaPayloadError(ValueError):
    """公式隐藏数据编码/解码错误。"""


def _json_default(obj: Any) -> str:
    return str(obj)


def _b64_encode(text: str) -> str:
    return base64.urlsafe_b64encode(text.encode("utf-8")).decode("ascii")


def _b64_decode(text: str) -> str:
    return base64.urlsafe_b64decode(text.encode("ascii")).decode("utf-8")


def normalize_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """
    统一并精简 Word/WPS 可编辑公式 payload。

    AlternativeText 只保存重建公式需要的信息，不保存 SVG/EMF、渲染平台、
    字体显示名称或编号域的冗余数据。
    """
    data = dict(payload or {})
    latex = str(data.get("latex") or data.get("tex") or "")
    mode = str(data.get("mode") or ("display" if data.get("display") else "inline")).lower()
    if mode not in {"inline", "display"}:
        mode = "display" if mode in {"1", "true", "yes"} else "inline"

    font_size = data.get("fontSize", data.get("font_size", data.get("font-size", 12)))
    try:
        font_size = float(font_size)
    except Exception:
        font_size = 12.0
    font_size = min(96.0, max(6.0, font_size))
    if font_size.is_integer():
        font_size = int(font_size)

    color = str(data.get("color") or "#000000").strip().upper()
    if not (len(color) == 7 and color.startswith("#")):
        color = "#000000"

    normalized: dict[str, Any] = {
        "type": TYPE_NAME,
        "version": VERSION,
        "latex": latex,
        "mode": mode,
        "fontSize": font_size,
        "mathFont": str(data.get("mathFont") or data.get("math_font") or "mathjax-newcm"),
        "color": color,
    }

    if bool(data.get("numbered")):
        normalized["numbered"] = True
        normalized["includeChapterNumber"] = bool(data.get("includeChapterNumber"))

    return normalized


def encode_payload(payload: dict[str, Any], *, max_svg_chars: int = 120_000) -> str:
    """
    把精简公式 JSON 编码成 AlternativeText。

    max_svg_chars 参数仅为兼容旧调用；新格式始终不写入 SVG 图片数据。
    """
    del max_svg_chars
    data = normalize_payload(payload)
    raw = json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=_json_default)
    return PREFIX + _b64_encode(raw)


def _decode_new_payload(text: str) -> dict[str, Any]:
    encoded = text[len(PREFIX):].strip()
    # 兼容从 HTML 属性中截取出来的情况。
    encoded = encoded.split()[0].strip('"\'<>])')
    data = json.loads(_b64_decode(encoded))
    if not isinstance(data, dict) or data.get("type") != TYPE_NAME:
        raise FormulaPayloadError("公式数据格式不正确。")
    return normalize_payload(data)


def _decode_legacy_payload(text: str) -> dict[str, Any]:
    encoded = text[len(LEGACY_PREFIX):].strip()
    try:
        compressed = base64.urlsafe_b64decode(encoded.encode("ascii"))
        raw = zlib.decompress(compressed)
        data = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise FormulaPayloadError(f"旧版公式数据解码失败：{exc}") from exc
    if not isinstance(data, dict) or data.get("type") != TYPE_NAME:
        raise FormulaPayloadError("旧版公式数据格式不正确。")
    return normalize_payload(data)


def decode_payload(text: str | None) -> dict[str, Any]:
    """从 AlternativeText 恢复公式 JSON，兼容新旧两种格式。"""
    if not text or not isinstance(text, str):
        raise FormulaPayloadError("选中对象没有保存公式数据。")
    if text.startswith(PREFIX):
        return _decode_new_payload(text)
    if text.startswith(LEGACY_PREFIX):
        return _decode_legacy_payload(text)

    # 如果 payload 混在 Title/HTML 中，也尝试查找前缀。
    idx = text.find(PREFIX)
    if idx >= 0:
        return _decode_new_payload(text[idx:])
    idx = text.find(LEGACY_PREFIX)
    if idx >= 0:
        return _decode_legacy_payload(text[idx:])
    raise FormulaPayloadError("选中对象不是由本工具插入的可编辑公式。")


def is_formula_payload(text: str | None) -> bool:
    return isinstance(text, str) and (PREFIX in text or LEGACY_PREFIX in text)
