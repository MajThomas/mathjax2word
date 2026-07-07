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
    """统一 Word/WPS 可编辑公式 payload 字段。"""
    data = dict(payload or {})
    latex = str(data.get("latex") or data.get("tex") or "")
    mode = str(data.get("mode") or ("display" if data.get("display") else "inline"))
    if mode not in {"inline", "display"}:
        mode = "display" if str(mode).lower() in {"1", "true", "yes"} else "inline"
    font_size = data.get("fontSize", data.get("font_size", data.get("font-size", 12)))
    try:
        font_size = float(font_size)
    except Exception:
        font_size = 12.0
    font_size = min(96.0, max(6.0, font_size))

    color = str(data.get("color") or "#000000").strip()
    if not (len(color) == 7 and color.startswith("#")):
        color = "#000000"

    svg_text = data.get("svgText")
    if not isinstance(svg_text, str) or not svg_text.strip():
        svg_text = data.get("svg")

    normalized: dict[str, Any] = {
        "type": TYPE_NAME,
        "version": VERSION,
        "latex": latex,
        "display": mode == "display",
        "mode": mode,
        "fontSize": font_size,
        "mathFont": str(data.get("mathFont") or data.get("math_font") or "mathjax-newcm"),
        "mathFontLabel": str(data.get("mathFontLabel") or data.get("math_font_label") or ""),
        "color": color.upper(),
        "renderer": str(data.get("renderer") or "mathjax-svg"),
        "displayImageFormat": str(data.get("displayImageFormat") or data.get("imageFormat") or "emf"),
        "platform": str(data.get("platform") or "word-routeB"),
    }

    passthrough_keys = (
        "numbered",
        "numberType",
        "sequenceName",
        "sequenceResetByHeadingLevel",
        "includeChapterNumber",
        "chapterNumber",
        "equationIndex",
        "sequenceNumber",
        "equationLabel",
    )
    for key in passthrough_keys:
        if key in data:
            normalized[key] = data[key]

    if isinstance(svg_text, str) and svg_text.strip():
        normalized["svgText"] = svg_text
    return normalized


def encode_payload(payload: dict[str, Any], *, max_svg_chars: int = 120_000) -> str:
    """把公式 JSON 编码成 Word/WPS 通用 AlternativeText 字符串。"""
    data = normalize_payload(payload)
    svg_text = data.get("svgText")
    if isinstance(svg_text, str) and len(svg_text) > max_svg_chars:
        # AlternativeText 不适合塞超大 SVG；LaTeX 和配置才是主数据。
        data.pop("svgText", None)
        data["svgOmitted"] = True
        data["svgOriginalChars"] = len(svg_text)

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
    """从 AlternativeText 字符串恢复公式 JSON。兼容新旧两种格式。"""
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
