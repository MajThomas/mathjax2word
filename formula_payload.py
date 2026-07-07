from __future__ import annotations

import base64
import json
import zlib
from dataclasses import dataclass
from typing import Any

PREFIX = "LATEX_SVG_FORMULA_V1:"
TYPE_NAME = "latex-svg-formula"
VERSION = 1


class FormulaPayloadError(ValueError):
    """公式隐藏数据编码/解码错误。"""


def _json_default(obj: Any) -> str:
    return str(obj)


def normalize_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """保留 Word 回读编辑所需字段，避免把不可序列化对象写入 Word。"""
    data = dict(payload or {})
    latex = str(data.get("latex") or data.get("tex") or "")
    mode = str(data.get("mode") or ("display" if data.get("display") else "inline"))
    if mode not in {"inline", "display"}:
        mode = "display" if str(mode).lower() in {"1", "true", "yes"} else "inline"
    font_size = data.get("fontSize", data.get("font_size", data.get("font-size", 18)))
    try:
        font_size = float(font_size)
    except Exception:
        font_size = 18.0
    font_size = min(96.0, max(6.0, font_size))

    color = str(data.get("color") or "#000000")
    if not (len(color) == 7 and color.startswith("#")):
        color = "#000000"

    normalized: dict[str, Any] = {
        "type": TYPE_NAME,
        "version": VERSION,
        "latex": latex,
        "mode": mode,
        "display": mode == "display",
        "fontSize": font_size,
        "color": color.upper(),
        "mathFont": str(data.get("mathFont") or data.get("math_font") or "mathjax-newcm"),
        "mathFontLabel": str(data.get("mathFontLabel") or data.get("math_font_label") or ""),
    }

    # 编号公式元数据。普通公式没有这些字段；编号公式回读/更新时需要保留。
    passthrough_keys = (
        "numbered",
        "includeChapterNumber",
        "chapterNumber",
        "equationIndex",
        "sequenceNumber",
        "equationLabel",
    )
    for key in passthrough_keys:
        if key in data:
            normalized[key] = data[key]
    # SVG 是备份字段。AlternativeText 有长度风险，所以只在前端传入时保留；过大时后端会自动裁剪。
    svg = data.get("svg")
    if isinstance(svg, str) and svg.strip():
        normalized["svg"] = svg
    return normalized


def encode_payload(payload: dict[str, Any], *, max_svg_chars: int = 120_000) -> str:
    """把公式 JSON 压缩成可写入 Word AlternativeText 的字符串。"""
    data = normalize_payload(payload)
    svg = data.get("svg")
    if isinstance(svg, str) and len(svg) > max_svg_chars:
        # Word 的 AlternativeText 不适合塞超大 SVG；LaTeX 和配置才是回读编辑的主数据。
        data.pop("svg", None)
        data["svgOmitted"] = True
        data["svgOriginalChars"] = len(svg)

    raw = json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=_json_default).encode("utf-8")
    compressed = zlib.compress(raw, level=9)
    encoded = base64.urlsafe_b64encode(compressed).decode("ascii")
    return PREFIX + encoded


def decode_payload(text: str | None) -> dict[str, Any]:
    """从 AlternativeText 字符串恢复公式 JSON。"""
    if not text or not isinstance(text, str):
        raise FormulaPayloadError("选中对象没有保存公式数据。")
    if not text.startswith(PREFIX):
        raise FormulaPayloadError("选中对象不是由本工具插入的可编辑公式。")
    encoded = text[len(PREFIX):].strip()
    try:
        compressed = base64.urlsafe_b64decode(encoded.encode("ascii"))
        raw = zlib.decompress(compressed)
        data = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise FormulaPayloadError(f"公式数据解码失败：{exc}") from exc
    if not isinstance(data, dict) or data.get("type") != TYPE_NAME:
        raise FormulaPayloadError("公式数据格式不正确。")
    return data


def is_formula_payload(text: str | None) -> bool:
    return isinstance(text, str) and text.startswith(PREFIX)
