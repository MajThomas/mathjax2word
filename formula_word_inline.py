from __future__ import annotations

import json
import math
import re
import sys
import time
from pathlib import Path
from typing import Any, Callable

from formula_payload import decode_payload, encode_payload, is_formula_payload
from emf_convert import svg_batch_to_emf, svg_to_emf

ROOT = Path(__file__).resolve().parent
TEMP_DIR = ROOT / "temp"
LAST_JSON = TEMP_DIR / "last_formula.json"
LAST_SVG = TEMP_DIR / "last_formula.svg"
LAST_EMF = TEMP_DIR / "last_formula.emf"

FORMULA_WORD_BUILD = "2026-07-17-conversion-section-fixed-v21"
# Word COM 常量，直接使用数字避免依赖 win32com.constants 初始化。
WD_FIELD_EMPTY = -1
WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
PLAIN_EQUATION_SEQUENCE_NAME = "公式"
CHAPTER_EQUATION_SEQUENCE_NAME = "章节公式"
EQUATION_SEQUENCE_NAME = PLAIN_EQUATION_SEQUENCE_NAME

# 四类 Word 文本定界符。必须按“长定界符优先”识别，避免 $ 与 $$、# 与 ## 互相串扰。
FORMULA_MARKER_SPECS = (
    {"marker": "##", "formulaType": "chapter-numbered", "mode": "display", "includeChapterNumber": True},
    {"marker": "$$", "formulaType": "display", "mode": "display", "includeChapterNumber": False},
    {"marker": "#", "formulaType": "numbered", "mode": "display", "includeChapterNumber": False},
    {"marker": "$", "formulaType": "inline", "mode": "inline", "includeChapterNumber": False},
)


# Word 的字符高级位置按 0.5 pt 写入。公式本身不缩放；SVG/EMF 的
# 空白只通过 PictureFormat.Crop* 裁剪，统一字体偏移由前端配置传入。
BASELINE_WORD_POSITION_STEP_PT = 0.5

# 行内公式统一图框。均可由前端 JSON 覆盖：
# inlineFrameHeightPt：统一图框高度
# inlineFrameBaselineDepthPt：数学基线到图框底边的距离
# inlineFrameSafetyPt：图元与图框边界的最小安全距离
# inlineTextOffsetDownPt：相对正文基线的最终下调量
INLINE_FRAME_HEIGHT_FACTOR = 1.60
INLINE_FRAME_BASELINE_DEPTH_FACTOR = 0.35
INLINE_FRAME_SAFETY_PT = 0.25

# 大型运算符的视觉校正。可由 JSON 的 largeOperatorOffsetDownPt 覆盖。
# 正值向下，负值向上。默认 0，不做经验补偿。
LARGE_OPERATOR_OFFSET_DOWN_PT = 0.0

# 初次扩大行内公式图框时，上下基本对称，下侧额外多留一部分空间。
# 可由 JSON 的 inlineFrameExtraBottomPt 覆盖。
INLINE_FRAME_EXTRA_BOTTOM_FACTOR = 0.12
LARGE_OPERATOR_PATTERN = re.compile(
    r"\\(?:int|iint|iiint|iiiint|oint|oiint|oiiint|sum|prod|coprod|bigcup|bigcap|bigsqcup|bigvee|bigwedge|bigodot|bigotimes|bigoplus|biguplus)\b"
)


class WordFormulaError(RuntimeError):
    """Word 公式插入/读取/更新错误。"""


def _require_windows() -> None:
    if not sys.platform.startswith("win"):
        raise WordFormulaError("Word 自动插入功能只能在 Windows + Microsoft Word 桌面版环境使用。")


def _word_app():
    _require_windows()
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise WordFormulaError("缺少 pywin32。请先运行：python -m pip install pywin32") from exc

    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception as exc:
        raise WordFormulaError("没有检测到正在运行的 Word。请先打开 Word 文档，并把光标放到要插入公式的位置。") from exc


def _document_identity(doc: Any) -> str:
    for attr in ("FullName", "Name"):
        try:
            value = str(getattr(doc, attr) or "").strip()
            if value:
                return value
        except Exception:
            pass
    return ""


def _utf16_length(text: str) -> int:
    """Word Range 使用 UTF-16 字符位置；这里把 Python 字符下标转换为 Word 偏移。"""
    return len(str(text).encode("utf-16-le")) // 2


def _is_escaped(text: str, index: int) -> bool:
    """判断 index 位置的定界符前是否有奇数个反斜杠。"""
    slash_count = 0
    pos = index - 1
    while pos >= 0 and text[pos] == "\\":
        slash_count += 1
        pos -= 1
    return slash_count % 2 == 1


def _double_marker_at(text: str, index: int, marker: str) -> bool:
    return index >= 0 and text.startswith(marker, index) and not _is_escaped(text, index)


def _single_marker_at(text: str, index: int, marker: str) -> bool:
    return index >= 0 and index < len(text) and text[index] == marker and not _is_escaped(text, index)


def _marker_spec_at(text: str, index: int) -> dict[str, Any] | None:
    for spec in FORMULA_MARKER_SPECS:
        marker = str(spec["marker"])
        if len(marker) == 2:
            matched = _double_marker_at(text, index, marker)
        else:
            matched = _single_marker_at(text, index, marker)
        if matched:
            return spec
    return None


def _find_closing_marker(text: str, start: int, marker: str) -> int:
    pos = start
    while pos < len(text):
        if len(marker) == 2:
            matched = _double_marker_at(text, pos, marker)
        else:
            matched = _single_marker_at(text, pos, marker)
        if matched:
            return pos
        pos += 1
    return -1


def _scan_formula_markers(text: str, *, base_start: int = 0) -> list[dict[str, Any]]:
    """扫描互不重叠的 LaTeX 定界符，返回 Word 绝对范围。"""
    matches: list[dict[str, Any]] = []
    pos = 0
    while pos < len(text):
        spec = _marker_spec_at(text, pos)
        if spec is None:
            pos += 1
            continue

        marker = str(spec["marker"])
        content_start = pos + len(marker)
        close_pos = _find_closing_marker(text, content_start, marker)
        if close_pos < 0:
            # 双字符起始符未闭合时，整段跳过，不能退化成单字符定界符。
            pos += len(marker)
            continue

        latex = text[content_start:close_pos]
        end_pos = close_pos + len(marker)
        if latex.strip():
            start_word = base_start + _utf16_length(text[:pos])
            content_start_word = base_start + _utf16_length(text[:content_start])
            content_end_word = base_start + _utf16_length(text[:close_pos])
            end_word = base_start + _utf16_length(text[:end_pos])
            matches.append({
                "start": start_word,
                "end": end_word,
                "contentStart": content_start_word,
                "contentEnd": content_end_word,
                "latex": latex,
                "sourceText": text[pos:end_pos],
                "marker": marker,
                "formulaType": spec["formulaType"],
                "mode": spec["mode"],
                "includeChapterNumber": bool(spec["includeChapterNumber"]),
            })
        pos = end_pos
    return matches


def _clean_conversion_payload(payload: dict[str, Any]) -> dict[str, Any]:
    data = dict(payload)
    for key in ("rangeStart", "rangeEnd", "start", "end", "sourceText", "marker", "markerType", "formulaType"):
        data.pop(key, None)
    return data


def _compact_formula_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """只保留回读、更新和回写公式真正需要的隐藏数据。"""
    mode = str(payload.get("mode") or "").lower()
    if mode not in ("inline", "display"):
        mode = "display" if bool(payload.get("display")) else "inline"

    data: dict[str, Any] = {
        "type": "latex-svg-formula",
        "version": 1,
        "latex": str(payload.get("latex") or ""),
        "mode": mode,
    }

    color = str(payload.get("color") or "").strip()
    if color:
        data["color"] = color

    font_size = payload.get("fontSize")
    if font_size is not None and font_size != "":
        try:
            numeric_size = float(font_size)
            data["fontSize"] = int(numeric_size) if numeric_size.is_integer() else numeric_size
        except Exception:
            data["fontSize"] = font_size

    math_font = str(payload.get("mathFont") or "").strip()
    if math_font:
        data["mathFont"] = math_font

    if bool(payload.get("numbered")):
        data["numbered"] = True
        data["includeChapterNumber"] = bool(payload.get("includeChapterNumber"))

    return data


def _encode_formula_alt(payload: dict[str, Any]) -> str:
    return encode_payload(_compact_formula_payload(payload))


def _quantize_half_point(value: float) -> float:
    """按文件顶部配置的 Word 字符位置步长四舍五入。"""
    value = max(-512.0, min(512.0, float(value)))
    step = max(0.1, float(BASELINE_WORD_POSITION_STEP_PT))
    sign = -1.0 if value < 0 else 1.0
    # 不使用 Python 的银行家舍入，确保步长边界按常规方式处理。
    return sign * (math.floor(abs(value) / step + 0.5) * step)


def _set_word_position_half_point(inline_shape: Any, position: float) -> None:
    """
    给图片字符设置半点级垂直位置。

    Word COM 的 Font.Position 声明为 Long，只能可靠写入整数点；Word 本身的
    字体高级设置支持 0.5 pt，因此小数半点通过 WordBasic.FormatFont 写入。
    """
    position = _quantize_half_point(position)
    if abs(position) < 0.001:
        try:
            inline_shape.Range.Font.Position = 0
        except Exception:
            pass
        return

    # 整数点优先走普通 COM 属性，速度更快，也不会改变当前 Selection。
    if abs(position - round(position)) < 0.001:
        try:
            inline_shape.Range.Font.Position = int(round(position))
            return
        except Exception:
            pass

    # 半点位置需要旧式 WordBasic 接口。该接口作用于当前选区，因此先保存并恢复。
    word = None
    current = None
    try:
        word = inline_shape.Application
        current = word.Selection.Range.Duplicate
        inline_shape.Range.Select()

        # WordBasic 在不同 Office/区域设置下对参数类型的接受方式略有差异：
        # 先按字符串写入，再按浮点数重试，避免 -0.5 被 COM 提前转成 Long。
        errors = []
        for candidate in (f"{position:.1f}", float(position)):
            try:
                word.WordBasic.FormatFont(Position=candidate)
                return
            except Exception as exc:
                errors.append(exc)
    except Exception:
        pass
    finally:
        if current is not None:
            try:
                current.Select()
            except Exception:
                pass

    # WordBasic 不可用时，不能使用 Python round(-0.5)==0 的银行家舍入，
    # 否则 0.5 pt 粗调会完全消失。回退时按远离零方向取整，至少保持方向正确。
    try:
        integer_position = math.floor(position) if position < 0 else math.ceil(position)
        inline_shape.Range.Font.Position = int(integer_position)
    except Exception:
        pass


def _safe_float(payload: dict[str, Any], key: str, default: float = 0.0) -> float:
    try:
        value = float(payload.get(key) or default)
    except Exception:
        return float(default)
    return value if math.isfinite(value) else float(default)


def _set_picture_crop_values(
    picture_format: Any,
    *,
    left: float,
    right: float,
    top: float,
    bottom: float,
) -> None:
    """一次性写入四边裁剪值。Crop* 只改变可见窗口，不缩放图片。"""
    picture_format.CropLeft = float(left)
    picture_format.CropRight = float(right)
    picture_format.CropTop = float(top)
    picture_format.CropBottom = float(bottom)


def _read_picture_crop_values(picture_format: Any) -> dict[str, float]:
    """读取 Word 实际保存的裁剪值，用于消除 COM/Office 内部舍入误差。"""
    result: dict[str, float] = {}
    for key, attr in (
        ("left", "CropLeft"),
        ("right", "CropRight"),
        ("top", "CropTop"),
        ("bottom", "CropBottom"),
    ):
        try:
            value = float(getattr(picture_format, attr))
        except Exception:
            value = 0.0
        result[key] = value if math.isfinite(value) else 0.0
    return result


def _apply_word_picture_crop(inline_shape: Any, payload: dict[str, Any]) -> dict[str, Any]:
    """设置 Word 图片基础裁剪。

    所有公式模式均完全取消横向裁剪：
    CropLeft = 0，CropRight = 0。

    行内公式和非编号行间公式上下也不裁剪，使用统一图框逻辑。
    编号公式仅保留原有上下裁剪，横向仍保持 EMF 原始宽度。
    """
    page_width_pt = max(0.0, _safe_float(payload, "pageWidthPt"))
    page_height_pt = max(0.0, _safe_float(payload, "pageHeightPt"))
    svg_top = max(0.0, _safe_float(payload, "cropTopPt"))
    svg_bottom = max(0.0, _safe_float(payload, "cropBottomPt"))

    try:
        actual_width_pt = max(0.0, float(inline_shape.Width))
        actual_height_pt = max(0.0, float(inline_shape.Height))
    except Exception as exc:
        raise WordFormulaError(f"读取 Word 中公式图片尺寸失败：{exc}") from exc

    extra_width_pt = (
        max(0.0, actual_width_pt - page_width_pt)
        if page_width_pt > 0
        else 0.0
    )
    extra_height_pt = (
        max(0.0, actual_height_pt - page_height_pt)
        if page_height_pt > 0
        else 0.0
    )

    tight_crop_top = svg_top
    tight_crop_bottom = svg_bottom + extra_height_pt

    mode = str(
        payload.get("mode")
        or ("display" if payload.get("display") else "inline")
    ).lower()
    numbered = bool(payload.get("numbered"))

    use_unified_frame = mode in ("inline", "display") and not numbered

    if not use_unified_frame and actual_height_pt > 0:
        max_vertical = max(0.0, actual_height_pt - 0.05)
        total_vertical = tight_crop_top + tight_crop_bottom
        if total_vertical > max_vertical and total_vertical > 0:
            scale = max_vertical / total_vertical
            tight_crop_top *= scale
            tight_crop_bottom *= scale

    try:
        picture_format = inline_shape.PictureFormat

        _set_picture_crop_values(
            picture_format,
            left=0.0,
            right=0.0,
            top=0.0,
            bottom=0.0,
        )

        try:
            crop_object = picture_format.Crop
            crop_object.PictureOffsetY = 0.0
            crop_object.PictureHeight = float(actual_height_pt)
        except Exception:
            pass

        if use_unified_frame:
            _set_picture_crop_values(
                picture_format,
                left=0.0,
                right=0.0,
                top=0.0,
                bottom=0.0,
            )
        else:
            _set_picture_crop_values(
                picture_format,
                left=0.0,
                right=0.0,
                top=tight_crop_top,
                bottom=tight_crop_bottom,
            )

        accepted = _read_picture_crop_values(picture_format)
    except Exception as exc:
        raise WordFormulaError(f"Word 原生裁剪公式图片失败：{exc}") from exc

    baseline_depth_pt = max(
        0.0,
        _safe_float(payload, "baselineDepthAfterCropPt"),
    )
    content_height_pt = max(
        0.0,
        actual_height_pt - tight_crop_top - tight_crop_bottom,
    )
    baseline_y_in_picture_pt = (
        actual_height_pt
        - tight_crop_bottom
        - baseline_depth_pt
    )

    return {
        "pictureFormat": picture_format,
        "actualWidthPt": actual_width_pt,
        "actualHeightPt": actual_height_pt,
        "extraWidthPt": extra_width_pt,
        "extraHeightPt": extra_height_pt,
        "cropLeftPt": 0.0,
        "cropRightPt": 0.0,
        "cropTopPt": accepted["top"],
        "cropBottomPt": accepted["bottom"],
        "tightCropTopPt": tight_crop_top,
        "tightCropBottomPt": tight_crop_bottom,
        "contentHeightPt": content_height_pt,
        "baselineDepthAfterCropPt": baseline_depth_pt,
        "baselineYInPicturePt": baseline_y_in_picture_pt,
        "useUnifiedFrame": use_unified_frame,
        "isDisplayFrame": mode == "display" and not numbered,
    }

def _first_positive_payload_float(
    payload: dict[str, Any],
    keys: tuple[str, ...],
    default: float,
) -> float:
    for key in keys:
        value = _safe_float(payload, key)
        if value > 0.0:
            return value
    return max(0.0, float(default))


def _resolve_inline_frame_geometry(
    payload: dict[str, Any],
    *,
    picture_height_pt: float,
    content_height_pt: float,
    formula_baseline_depth_pt: float,
) -> tuple[float, float, float, float]:
    """计算统一图框高度、固定基线深度、安全边和下侧额外余量。

    初次扩图框时不再按照公式自身下伸深度扩大上侧空间，而是：
    - 上下基本对称；
    - 下侧额外多留 inlineFrameExtraBottomPt；
    - 所有公式仍使用固定的图框基线深度。
    """
    del content_height_pt, formula_baseline_depth_pt

    font_size_pt = _first_positive_payload_float(
        payload,
        ("fontSize",),
        12.0,
    )
    safety_pt = _first_positive_payload_float(
        payload,
        ("inlineFrameSafetyPt",),
        INLINE_FRAME_SAFETY_PT,
    )
    configured_height_pt = _first_positive_payload_float(
        payload,
        ("inlineFrameHeightPt", "maxLineHeightPt", "lineHeightPt"),
        font_size_pt * INLINE_FRAME_HEIGHT_FACTOR,
    )
    fixed_depth_pt = _first_positive_payload_float(
        payload,
        (
            "inlineFrameBaselineDepthPt",
            "frameBaselineDepthPt",
            "lineDescentPt",
        ),
        font_size_pt * INLINE_FRAME_BASELINE_DEPTH_FACTOR,
    )
    extra_bottom_pt = _first_positive_payload_float(
        payload,
        ("inlineFrameExtraBottomPt",),
        font_size_pt * INLINE_FRAME_EXTRA_BOTTOM_FACTOR,
    )

    # 图框至少比原始 EMF 高：上侧 safety，下侧 safety + extra_bottom。
    minimum_height_pt = (
        float(picture_height_pt)
        + 2.0 * safety_pt
        + extra_bottom_pt
    )
    frame_height_pt = max(
        configured_height_pt,
        minimum_height_pt,
    )

    return (
        frame_height_pt,
        fixed_depth_pt,
        safety_pt,
        extra_bottom_pt,
    )

def _large_operator_visual_offset_down_pt(
    payload: dict[str, Any],
) -> float:
    """返回积分、求和等大型运算符的独立视觉微调量。"""
    latex = str(payload.get("latex") or "")
    if not LARGE_OPERATOR_PATTERN.search(latex):
        return 0.0
    return _safe_float(
        payload,
        "largeOperatorOffsetDownPt",
        LARGE_OPERATOR_OFFSET_DOWN_PT,
    )

def _apply_inline_formula_frame(
    inline_shape: Any,
    crop_state: dict[str, Any],
    payload: dict[str, Any],
    *,
    picture_micro_down_pt: float,
) -> dict[str, float]:
    """固定图框高度，将上侧多余空白转移到下侧。

    图框高度只设置一次，且不因积分式越界而扩大。

    若最终图元下侧空间不足：
    1. 计算上侧可用空白；
    2. 将图元向上移动，消耗上侧空白；
    3. 返回相同大小的整体下移补偿量；
    4. 由 _apply_word_baseline 把整个行内对象向下移动同样距离。

    这样公式数学基线在文档中的最终位置不变，但图框内的空白从上侧
    转移到了下侧，不会裁掉积分式下半部分。
    """
    picture_format = crop_state["pictureFormat"]

    try:
        crop_object = picture_format.Crop
    except Exception as exc:
        raise WordFormulaError(
            f"当前 Word 版本无法取得图片裁剪框对象：{exc}"
        ) from exc

    picture_height_pt = max(
        0.05,
        float(crop_state["actualHeightPt"]),
    )
    content_height_pt = max(
        0.0,
        float(crop_state["contentHeightPt"]),
    )
    formula_depth_pt = max(
        0.0,
        float(crop_state["baselineDepthAfterCropPt"]),
    )
    baseline_y_picture_pt = float(crop_state["baselineYInPicturePt"])

    (
        frame_height_pt,
        fixed_frame_depth_pt,
        safety_pt,
        extra_bottom_pt,
    ) = _resolve_inline_frame_geometry(
        payload,
        picture_height_pt=picture_height_pt,
        content_height_pt=content_height_pt,
        formula_baseline_depth_pt=formula_depth_pt,
    )

    _set_picture_crop_values(
        picture_format,
        left=float(crop_state["cropLeftPt"]),
        right=float(crop_state["cropRightPt"]),
        top=0.0,
        bottom=0.0,
    )

    try:
        crop_object.PictureHeight = float(picture_height_pt)
        actual_picture_height_pt = float(crop_object.PictureHeight)

        # ShapeHeight 仅设置一次。
        crop_object.ShapeHeight = float(frame_height_pt)
        actual_frame_height_pt = float(crop_object.ShapeHeight)
    except Exception as exc:
        raise WordFormulaError(
            f"当前 Word 版本无法建立统一行内公式图框：{exc}"
        ) from exc

    if actual_frame_height_pt <= actual_picture_height_pt + 0.01:
        raise WordFormulaError(
            "Word 未接受统一公式图框高度设置。"
            f"图元高度={actual_picture_height_pt:.3f} pt，"
            f"图框高度={actual_frame_height_pt:.3f} pt。"
        )

    large_operator_offset_pt = _large_operator_visual_offset_down_pt(payload)
    requested_micro_pt = (
        float(picture_micro_down_pt)
        + large_operator_offset_pt
    )

    # 先按统一数学基线计算原始图元位置。
    target_baseline_y_frame_pt = (
        actual_frame_height_pt - fixed_frame_depth_pt
    )
    base_offset_y_pt = (
        target_baseline_y_frame_pt
        - actual_frame_height_pt / 2.0
        - (
            baseline_y_picture_pt
            - actual_picture_height_pt / 2.0
        )
    )
    original_offset_y_pt = base_offset_y_pt + requested_micro_pt

    original_picture_top_pt = (
        actual_frame_height_pt / 2.0
        + original_offset_y_pt
        - actual_picture_height_pt / 2.0
    )
    original_picture_bottom_pt = (
        original_picture_top_pt + actual_picture_height_pt
    )

    top_available_pt = max(
        0.0,
        original_picture_top_pt - safety_pt,
    )
    bottom_missing_pt = max(
        0.0,
        original_picture_bottom_pt
        - (actual_frame_height_pt - safety_pt),
    )

    # 只利用现有上侧空白，不改变 ShapeHeight。
    transfer_up_pt = min(
        top_available_pt,
        bottom_missing_pt,
    )

    final_offset_y_pt = original_offset_y_pt - transfer_up_pt

    final_picture_top_pt = (
        actual_frame_height_pt / 2.0
        + final_offset_y_pt
        - actual_picture_height_pt / 2.0
    )
    final_picture_bottom_pt = (
        final_picture_top_pt + actual_picture_height_pt
    )

    # 如果上侧空白不足以完全解决下侧越界，则停止插入，避免静默裁切。
    if final_picture_bottom_pt > actual_frame_height_pt - safety_pt + 0.01:
        raise WordFormulaError(
            "统一图框内上侧空白不足，无法完全转移到下侧。"
            f"上侧可转移={top_available_pt:.3f} pt，"
            f"下侧缺口={bottom_missing_pt:.3f} pt。"
            "请适当增大 inlineFrameHeightPt 或减小大型运算符下移量。"
        )

    try:
        crop_object.PictureOffsetY = float(final_offset_y_pt)
        actual_offset_y_pt = float(crop_object.PictureOffsetY)
    except Exception as exc:
        raise WordFormulaError(
            f"当前 Word 版本无法移动公式图元：{exc}"
        ) from exc

    actual_picture_top_pt = (
        actual_frame_height_pt / 2.0
        + actual_offset_y_pt
        - actual_picture_height_pt / 2.0
    )
    actual_picture_bottom_pt = (
        actual_picture_top_pt + actual_picture_height_pt
    )

    if (
        actual_picture_top_pt < safety_pt - 0.02
        or actual_picture_bottom_pt
        > actual_frame_height_pt - safety_pt + 0.02
    ):
        raise WordFormulaError(
            "Word 实际保存的图元偏移仍会导致公式超出固定图框。"
        )

    return {
        "frameHeightPt": actual_frame_height_pt,
        "pictureHeightPt": actual_picture_height_pt,
        "frameBaselineDepthPt": fixed_frame_depth_pt,
        "frameSafetyPt": safety_pt,
        "frameExtraBottomPt": extra_bottom_pt,
        "formulaBaselineDepthPt": formula_depth_pt,
        "basePictureOffsetYPt": base_offset_y_pt,
        "largeOperatorOffsetDownPt": large_operator_offset_pt,
        "originalPictureOffsetYPt": original_offset_y_pt,
        "pictureOffsetYPt": actual_offset_y_pt,
        "pictureTopPt": actual_picture_top_pt,
        "pictureBottomPt": actual_picture_bottom_pt,
        "topSpaceTransferredPt": transfer_up_pt,
        # 图元向上多少，整个行内对象就需要向下补偿多少。
        "wordCompensationDownPt": transfer_up_pt,
        "actualMicroDownPt": (
            actual_offset_y_pt
            - base_offset_y_pt
            - large_operator_offset_pt
            + transfer_up_pt
        ),
    }

def _half_point_candidates(target_down_pt: float) -> list[float]:
    """生成目标值附近的半点档位，优先选择距离最近的档位。"""
    step = max(0.1, float(BASELINE_WORD_POSITION_STEP_PT))
    center_index = int(math.floor(target_down_pt / step + 0.5))
    values = [(center_index + delta) * step for delta in range(-4, 5)]
    return sorted(set(values), key=lambda value: (abs(value - target_down_pt), abs(value)))


def _split_baseline_between_word_and_crop(
    target_down_pt: float,
    crop_top_pt: float,
    crop_bottom_pt: float,
    retained_top_safety_pt: float,
    retained_bottom_safety_pt: float,
) -> tuple[float, float]:
    """把正文基线校准量拆成 Word 0.5 pt 粗调和图元微调。"""
    del crop_top_pt, crop_bottom_pt
    del retained_top_safety_pt, retained_bottom_safety_pt

    target = max(-512.0, min(512.0, float(target_down_pt)))
    word_down_pt = _quantize_half_point(target)
    picture_micro_down_pt = target - word_down_pt
    return word_down_pt, picture_micro_down_pt

def _apply_word_baseline(inline_shape: Any, payload: dict[str, Any]) -> None:
    """行内与非编号行间公式共用同一图框和基线校准逻辑。"""
    crop_state = _apply_word_picture_crop(inline_shape, payload)

    mode = str(
        payload.get("mode")
        or ("display" if payload.get("display") else "inline")
    ).lower()
    numbered = bool(payload.get("numbered"))

    if numbered or mode not in ("inline", "display"):
        _set_word_position_half_point(inline_shape, 0.0)
        return

    # $...$ 与 $$...$$ 完全复用同一套校准参数。
    text_offset_down_pt = _safe_float(
        payload,
        "inlineTextOffsetDownPt",
        _safe_float(payload, "fontOffsetDownPt"),
    )
    text_offset_down_pt = max(
        -512.0,
        min(512.0, text_offset_down_pt),
    )

    word_down_pt, picture_micro_down_pt = _split_baseline_between_word_and_crop(
        text_offset_down_pt,
        0.0,
        0.0,
        0.0,
        0.0,
    )

    frame_state = _apply_inline_formula_frame(
        inline_shape,
        crop_state,
        payload,
        picture_micro_down_pt=picture_micro_down_pt,
    )

    transfer_compensation_pt = float(
        frame_state.get("wordCompensationDownPt", 0.0)
    )
    total_word_down_pt = word_down_pt + transfer_compensation_pt

    accepted_word_down_pt = _quantize_half_point(total_word_down_pt)
    _set_word_position_half_point(
        inline_shape,
        -accepted_word_down_pt,
    )

    word_rounding_error_pt = (
        total_word_down_pt - accepted_word_down_pt
    )

    if abs(word_rounding_error_pt) > 0.005:
        try:
            crop_object = crop_state["pictureFormat"].Crop
            crop_object.PictureOffsetY = float(
                float(crop_object.PictureOffsetY)
                + word_rounding_error_pt
            )
        except Exception:
            pass

def _visible_paragraph_text(text: str) -> str:
    return str(text or "").replace("\r", "").replace("\x07", "").strip()


def _prepare_block_insertion_range(doc: Any, target_range: Any) -> Any:
    """
    把目标定界符替换为独立段落插槽，供行间公式或编号表格使用。

    若定界符已经独占段落，只清空定界符；若与正文同行，则按前后正文情况插入
    一个或两个段落标记，保证公式成为独立块且不改变相邻正文段落的对齐方式。
    """
    start = int(target_range.Start)
    end = int(target_range.End)
    try:
        para = target_range.Paragraphs.Item(1).Range
        para_start = int(para.Start)
        para_end = int(para.End)
        before = str(doc.Range(para_start, start).Text or "")
        after = str(doc.Range(end, para_end).Text or "")
        has_before = bool(_visible_paragraph_text(before))
        has_after = bool(_visible_paragraph_text(after))
        spans_multiple_paragraphs = int(target_range.Paragraphs.Count) > 1
    except Exception:
        has_before = True
        has_after = True
        spans_multiple_paragraphs = True

    if not spans_multiple_paragraphs and not has_before and not has_after:
        target_range.Text = ""
        return doc.Range(start, start)

    if has_before and has_after:
        target_range.Text = "\r\r"
        return doc.Range(start + 1, start + 1)
    if has_before:
        target_range.Text = "\r"
        return doc.Range(start + 1, start + 1)
    if has_after:
        target_range.Text = "\r"
        return doc.Range(start, start)

    target_range.Text = ""
    return doc.Range(start, start)


def _write_formula_source_files(payload: dict[str, Any], *, stem: str | None = None) -> tuple[Path, Path, Path]:
    """只写 JSON/SVG，不启动 Inkscape；批量转换时可先准备全部源文件。"""
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    svg = payload.get("svgText") or payload.get("svg")
    if not isinstance(svg, str) or "<svg" not in svg:
        raise WordFormulaError("前端没有传入有效 SVG，无法生成 EMF。")

    if stem:
        safe_stem = re.sub(r"[^0-9A-Za-z_.-]+", "_", stem).strip("._") or "formula"
        json_path = TEMP_DIR / f"{safe_stem}.json"
        svg_path = TEMP_DIR / f"{safe_stem}.svg"
        emf_path = TEMP_DIR / f"{safe_stem}.emf"
    else:
        json_path, svg_path, emf_path = LAST_JSON, LAST_SVG, LAST_EMF

    # 紧凑 SVG 原样交给 Inkscape 直接转换；Word 端只做原生图片裁剪。
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    svg_path.write_text(svg, encoding="utf-8")
    try:
        emf_path.unlink(missing_ok=True)
    except Exception:
        pass
    return json_path, svg_path, emf_path


def _save_formula_files(payload: dict[str, Any], *, stem: str | None = None) -> tuple[Path, Path, Path]:
    """紧凑 SVG 直接经 Inkscape 转换为 EMF。"""
    json_path, svg_path, emf_path = _write_formula_source_files(payload, stem=stem)
    svg_to_emf(svg_path, emf_path)
    return json_path, svg_path, emf_path


def _set_alt_text(obj: Any, alt_text: str) -> None:
    try:
        obj.AlternativeText = alt_text
    except Exception:
        try:
            obj.Title = "LaTeX Formula"
            obj.AlternativeText = alt_text
        except Exception as exc:
            raise WordFormulaError("公式已插入，但隐藏数据写入失败。") from exc


def _get_alt_text(obj: Any) -> str:
    for attr in ("AlternativeText", "Title"):
        try:
            value = getattr(obj, attr)
            if isinstance(value, str) and value:
                return value
        except Exception:
            pass
    return ""


def _selection_has_inline_shape(selection: Any) -> bool:
    try:
        return int(selection.InlineShapes.Count) >= 1
    except Exception:
        return False


def _selection_has_shape(selection: Any) -> bool:
    try:
        return int(selection.ShapeRange.Count) >= 1
    except Exception:
        return False


def _selected_object(selection: Any) -> tuple[str, Any]:
    if _selection_has_inline_shape(selection):
        return "inline", selection.InlineShapes.Item(1)
    if _selection_has_shape(selection):
        return "shape", selection.ShapeRange.Item(1)
    raise WordFormulaError("请先在 Word 中选中一个由本工具插入的公式图片。")


def _cell_content_range(cell: Any) -> Any:
    """返回单元格正文范围，排除末尾单元格标记。"""
    rng = cell.Range
    try:
        rng.End = rng.End - 1
    except Exception:
        pass
    return rng


def _style_name(paragraph: Any) -> str:
    try:
        style = paragraph.Style
        name_local = getattr(style, "NameLocal", "")
        name = getattr(style, "Name", "")
        return f"{name_local} {name}".strip()
    except Exception:
        return ""


def _is_heading1(paragraph: Any) -> bool:
    """尽量识别 Word 的一级标题。兼容英文 Heading 1 与中文 标题 1。"""
    try:
        # Word 的正文一般为 10，Heading 1 通常为 1。
        if int(paragraph.OutlineLevel) == 1:
            return True
    except Exception:
        pass
    name = _style_name(paragraph).lower().replace(" ", "")
    return any(token in name for token in ("heading1", "标题1", "標題1"))


def _paragraph_text(paragraph: Any) -> str:
    try:
        return str(paragraph.Range.Text or "").strip().replace("\r", "").replace("\x07", "")
    except Exception:
        return ""


def _clean_chapter_number(value: str | None) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    # 优先提取 1、1.2、2-3 这类编号；若为“第1章”，也会得到 1。
    match = re.search(r"\d+(?:[.\-．·]\d+)*", text)
    if match:
        return match.group(0).replace("．", ".").replace("·", ".").strip(".-")
    # 没有阿拉伯数字时，保留去掉常见标点后的原始编号，例如“一”。
    return text.strip(".．、，,：: ）)（(")


def _current_chapter_number(word: Any, selection: Any) -> str:
    """获取当前光标之前最近一个一级标题的编号。"""
    try:
        doc = word.ActiveDocument
        selection_start = int(selection.Range.Start)
        fallback_count = 0
        current_number = ""
        for idx in range(1, int(doc.Paragraphs.Count) + 1):
            paragraph = doc.Paragraphs.Item(idx)
            try:
                if int(paragraph.Range.Start) > selection_start:
                    break
            except Exception:
                pass
            if not _is_heading1(paragraph):
                continue
            fallback_count += 1
            list_string = ""
            try:
                list_string = str(paragraph.Range.ListFormat.ListString or "")
            except Exception:
                list_string = ""
            current_number = _clean_chapter_number(list_string) or _clean_chapter_number(_paragraph_text(paragraph)) or str(fallback_count)
        return current_number
    except Exception:
        return ""


def _iter_formula_payloads(doc: Any):
    """遍历文档中由本工具插入的公式隐藏数据。"""
    try:
        for idx in range(1, int(doc.InlineShapes.Count) + 1):
            obj = doc.InlineShapes.Item(idx)
            alt = _get_alt_text(obj)
            if is_formula_payload(alt):
                try:
                    yield decode_payload(alt), obj
                except Exception:
                    continue
    except Exception:
        pass
    try:
        for idx in range(1, int(doc.Shapes.Count) + 1):
            obj = doc.Shapes.Item(idx)
            alt = _get_alt_text(obj)
            if is_formula_payload(alt):
                try:
                    yield decode_payload(alt), obj
                except Exception:
                    continue
    except Exception:
        pass


def _next_equation_index(doc: Any, *, include_chapter: bool, chapter_number: str) -> int:
    """根据文档中既有编号公式数量生成下一个编号。删除或移动公式后不会自动重排。"""
    max_index = 0
    for data, _obj in _iter_formula_payloads(doc):
        if not data.get("numbered"):
            continue
        if include_chapter:
            if str(data.get("chapterNumber") or "") != str(chapter_number or ""):
                continue
        try:
            idx = int(data.get("equationIndex") or data.get("sequenceNumber") or 0)
        except Exception:
            idx = 0
        if idx > max_index:
            max_index = idx
    return max_index + 1


def _make_equation_label(index: int, *, include_chapter: bool, chapter_number: str) -> str:
    if include_chapter and chapter_number:
        return f"({chapter_number}.{index})"
    return f"({index})"


def _format_equation_table(table: Any) -> None:
    """设置编号公式表格为无边框、整行宽度、左右窄中间宽。"""
    try:
        table.Borders.Enable = False
    except Exception:
        pass
    try:
        table.AllowAutoFit = False
    except Exception:
        pass
    try:
        table.PreferredWidthType = 2  # wdPreferredWidthPercent
        table.PreferredWidth = 100
    except Exception:
        pass
    try:
        # 左右列相等，中间列居中，保证公式中心尽量落在页面中心。
        widths = [15, 70, 15]
        for col_idx, width in enumerate(widths, start=1):
            col = table.Columns.Item(col_idx)
            col.PreferredWidthType = 2  # wdPreferredWidthPercent
            col.PreferredWidth = width
    except Exception:
        pass
    try:
        for cell_idx in range(1, 4):
            table.Cell(1, cell_idx).VerticalAlignment = 1  # wdCellAlignVerticalCenter
    except Exception:
        pass
    try:
        table.Rows.Item(1).Range.ParagraphFormat.SpaceBefore = 0
        table.Rows.Item(1).Range.ParagraphFormat.SpaceAfter = 0
    except Exception:
        pass


def _field_result_text(field: Any) -> str:
    try:
        return str(field.Result.Text or "").strip()
    except Exception:
        return ""


def _insert_field_at_selection(selection: Any, field_code: str) -> Any:
    """在当前 Word Selection 位置插入域，并把光标移动到域结果后。"""
    field = selection.Fields.Add(
        Range=selection.Range,
        Type=WD_FIELD_EMPTY,
        Text=field_code,
        PreserveFormatting=False,
    )
    try:
        field.Update()
    except Exception:
        pass
    try:
        field.Result.Select()
        selection.Collapse(WD_COLLAPSE_END)
    except Exception:
        pass
    return field


def _ensure_formula_caption_labels(word: Any) -> None:
    """首次插入编号公式时自动建立 Word 交叉引用题注标签。"""
    labels = (
        PLAIN_EQUATION_SEQUENCE_NAME,
        CHAPTER_EQUATION_SEQUENCE_NAME,
    )
    for name in labels:
        exists = False
        try:
            count = int(word.CaptionLabels.Count)
            for index in range(1, count + 1):
                try:
                    if str(word.CaptionLabels.Item(index).Name) == name:
                        exists = True
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if exists:
            continue
        try:
            word.CaptionLabels.Add(Name=name)
        except Exception:
            # 已存在、版本差异或语言环境差异时不阻断公式插入。
            pass



_CHINESE_DIGITS = {
    "零": 0,
    "〇": 0,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
}


def _parse_chinese_number_1_to_99(text: str) -> int | None:
    """
    从标题编号结果中解析 1～99 的常见中文数字：
    一～九、十、十一～十九、二十～九十九。
    """
    value = str(text or "").strip()
    if not value:
        return None

    # 只保留中文数字字符，兼容“第十一章”“十一、”等结果。
    chars = [
        char for char in value
        if char in _CHINESE_DIGITS or char == "十"
    ]
    if not chars:
        return None

    token = "".join(chars)

    if token == "十":
        return 10

    if "十" in token:
        left, right = token.split("十", 1)
        tens = 1 if left == "" else _CHINESE_DIGITS.get(left)
        ones = 0 if right == "" else _CHINESE_DIGITS.get(right)
        if tens is None or ones is None:
            return None
        number = tens * 10 + ones
    elif len(token) == 1:
        number = _CHINESE_DIGITS.get(token)
    else:
        return None

    if number is None or not 1 <= number <= 99:
        return None
    return number


def _field_result_is_invalid(text: str) -> bool:
    value = str(text or "").strip()
    if not value:
        return True
    upper = value.upper()
    return (
        "ERROR!" in upper
        or "错误" in value
        or "未找到" in value
        or "NO TEXT OF SPECIFIED STYLE" in upper
    )


def _insert_heading_number_field(
    selection: Any,
    *,
    style_level: int,
) -> Any:
    """
    先插入普通 STYLEREF 域并更新，读取当前标题编号。

    - 阿拉伯数字标题：保留普通 STYLEREF；
    - 中文数字 1～30：改用“日”日期转换域；
    - 中文数字 31～99：改用“年”日期转换域。

    转换后仍是 Word 域，可通过 F9 更新。
    """
    style_level = 1 if int(style_level) <= 1 else 2
    # 题注“包含章节号”实际使用的是 STYLEREF <级别> \s。
    # 直接以同一开关读取当前层级编号，避免 Heading 2 的 \n 返回完整上下文。
    normal_code = f"STYLEREF {style_level} \\s"
    field = _insert_field_at_selection(selection, normal_code)

    try:
        field.Update()
    except Exception:
        pass

    result_text = _field_result_text(field)
    number = _parse_chinese_number_1_to_99(result_text)
    if number is None:
        return field

    try:
        field_range = field.Result.Duplicate
        field.Delete()
        field_range.Select()
    except Exception:
        try:
            field.Result.Select()
            field.Delete()
        except Exception:
            pass

    if number <= 30:
        # 1～30：使用十一月的“日”位转换，避免 31 日不存在。
        quote_code = (
            'QUOTE "二零二五年十一月'
            f'{{ STYLEREF {style_level} \\s }}'
            '日" \\@ "D"'
        )
    else:
        # 31～99：使用年份转换。
        # “一九”作为固定前缀，使 31～99 形成 1931～1999，
        # 再通过 \@ "Y" 仅输出年份；随后由外层公式编号读取其后两位。
        quote_code = (
            'QUOTE "一九'
            f'{{ STYLEREF {style_level} \\s }}'
            '年十一月一日" \\@ "Y"'
        )

    converted = _insert_field_at_selection(selection, quote_code)
    try:
        converted.Update()
    except Exception:
        pass
    return converted


# Word 域说明：
# 标题编号为中文时按需转换：
# 1～30 使用 QUOTE 日期“日”位；31～99 使用 QUOTE 日期“年”位。
# 不扫描全文，只处理当前公式所属标题；所有结果均保持为 Word 域并支持 F9 更新。
def _insert_equation_number_fields(
    word: Any,
    cell: Any,
    *,
    include_chapter: bool,
    include_section: bool = False,
) -> str:
    _ensure_formula_caption_labels(word)

    right_rng = _cell_content_range(cell)
    try:
        right_rng.Text = ""
        right_rng.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
    except Exception:
        pass

    right_rng.Select()
    selection = word.Selection
    selection.TypeText("(")
    parts: list[str] = []

    if include_chapter:
        chapter_field = _insert_heading_number_field(
            selection,
            style_level=1,
        )
        parts.append(_field_result_text(chapter_field))
        selection.TypeText(".")

        if include_section:
            section_field = _insert_heading_number_field(
                selection,
                style_level=2,
            )
            section_text = _field_result_text(section_field)
            if _field_result_is_invalid(section_text):
                raise WordFormulaError(
                    "已选择“章节编号含节号”，但当前位置之前未能读取有效的"
                    "二级标题编号。请确认当前公式位于使用“标题 2/Heading 2”"
                    "样式的二级标题之后。程序不会再自动降级为章.编号。"
                )
            parts.append(section_text)
            selection.TypeText(".")
            seq_code = (
                f"SEQ {CHAPTER_EQUATION_SEQUENCE_NAME} "
                "\\* ARABIC \\s 2"
            )
        else:
            seq_code = (
                f"SEQ {CHAPTER_EQUATION_SEQUENCE_NAME} "
                "\\* ARABIC \\s 1"
            )
    else:
        seq_code = (
            f"SEQ {PLAIN_EQUATION_SEQUENCE_NAME} \\* ARABIC"
        )

    seq_field = _insert_field_at_selection(selection, seq_code)
    parts.append(_field_result_text(seq_field))
    selection.TypeText(")")

    try:
        cell.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
        cell.Range.Fields.Update()
    except Exception:
        pass

    visible = [part for part in parts if part]
    if include_section and len(visible) < 3:
        raise WordFormulaError(
            "章.节.编号域未完整生成，已中止插入，避免错误地插入章.编号。"
        )
    return f"({'.'.join(visible)})" if visible else "Word 域编号"


def insert_formula(payload: dict[str, Any]) -> dict[str, Any]:
    """在当前 Word 光标处插入 EMF，并把公式 JSON 写入 AlternativeText。"""
    word = _word_app()
    selection = word.Selection
    _, _, emf_path = _save_formula_files(payload)
    alt_text = _encode_formula_alt(payload)
    is_display = bool(payload.get("display")) or str(payload.get("mode") or "").lower() == "display"

    try:
        if is_display:
            # 行间公式按独立居中段落插入；这与编号公式无关。
            rng = selection.Range
            try:
                if rng.Start != rng.End:
                    rng.Delete()
            except Exception:
                pass
            try:
                selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
                selection.ParagraphFormat.SpaceBefore = 0
                selection.ParagraphFormat.SpaceAfter = 0
            except Exception:
                pass

        inline_shape = selection.InlineShapes.AddPicture(
            FileName=str(emf_path),
            LinkToFile=False,
            SaveWithDocument=True,
        )
        _set_alt_text(inline_shape, alt_text)
        _apply_word_baseline(inline_shape, payload)

        if not is_display:
            try:
                selection.Collapse(WD_COLLAPSE_END)
                selection.Font.Position = 0
            except Exception:
                pass

        if is_display:
            try:
                # 光标移动到公式后方新段落，并恢复左对齐，方便继续输入正文。
                selection.Collapse(WD_COLLAPSE_END)
                selection.TypeParagraph()
                selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_LEFT
            except Exception:
                pass

        return {
            "ok": True,
            "message": "已插入 Word 行间公式。EMF 用于显示，公式数据已隐藏保存。" if is_display else "已插入 Word 行内公式。EMF 用于显示，公式数据已隐藏保存。",
            "emf": str(emf_path),
        }
    except Exception as exc:
        raise WordFormulaError(f"插入 Word 失败：{exc}") from exc


def _payload_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    return str(value).strip().lower() in {
        "1", "true", "yes", "on", "chapter-section"
    }


def insert_numbered_formula(payload: dict[str, Any]) -> dict[str, Any]:
    """以“无边框三列表格”形式插入居中公式和右侧 Word SEQ 域编号。"""
    word = _word_app()
    selection = word.Selection
    doc = word.ActiveDocument

    include_chapter = _payload_bool(payload.get("includeChapterNumber"))
    chapter_number_mode = str(
        payload.get("chapterNumberMode") or "chapter"
    ).strip().lower()
    include_section = include_chapter and (
        chapter_number_mode == "chapter-section"
        or _payload_bool(payload.get("includeSectionNumber"))
    )

    payload = dict(payload)
    payload["mode"] = "display"
    payload["display"] = True
    payload["numbered"] = True
    payload["numberType"] = "word-seq"
    payload["sequenceName"] = (
        CHAPTER_EQUATION_SEQUENCE_NAME
        if include_chapter
        else PLAIN_EQUATION_SEQUENCE_NAME
    )
    payload["sequenceResetByHeadingLevel"] = (
        2 if include_section else (1 if include_chapter else None)
    )
    payload["includeSectionNumber"] = include_section
    payload["chapterNumberMode"] = (
        "chapter-section" if include_section else "chapter"
    )
    payload["includeChapterNumber"] = include_chapter

    _, _, emf_path = _save_formula_files(payload)
    alt_text = _encode_formula_alt(payload)

    try:
        rng = selection.Range
        # 如果有选区，先删除选区内容。
        try:
            if rng.Start != rng.End:
                rng.Delete()
        except Exception:
            pass

        table = doc.Tables.Add(Range=rng, NumRows=1, NumColumns=3)
        _format_equation_table(table)

        left_rng = _cell_content_range(table.Cell(1, 1))
        center_rng = _cell_content_range(table.Cell(1, 2))

        try:
            left_rng.Text = ""
        except Exception:
            pass

        try:
            center_rng.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
        except Exception:
            pass
        formula_obj = center_rng.InlineShapes.AddPicture(
            FileName=str(emf_path),
            LinkToFile=False,
            SaveWithDocument=True,
        )
        _set_alt_text(formula_obj, alt_text)
        _apply_word_baseline(formula_obj, payload)

        equation_label = _insert_equation_number_fields(word, table.Cell(1, 3), include_chapter=include_chapter, include_section=include_section)

        # 更新全文域，保证前面插入/移动公式后编号可刷新。
        try:
            doc.Fields.Update()
        except Exception:
            pass

        # 将光标移到表格后方，方便继续输入正文。
        try:
            after_rng = table.Range
            after_rng.Collapse(WD_COLLAPSE_END)
            after_rng.InsertParagraphAfter()
            after_rng.Collapse(WD_COLLAPSE_END)
            after_rng.Select()
        except Exception:
            pass

        if include_section:
            msg = (
                f"已插入章节公式 {equation_label}。题注标签“章节公式”已自动建立；"
                "编号按二级标题重置，章号和节号强制显示为阿拉伯数字，按 F9 可更新。"
            )
        elif include_chapter:
            msg = (
                f"已插入章节公式 {equation_label}。题注标签“章节公式”已自动建立；"
                "编号按一级标题重置，章号强制显示为阿拉伯数字，按 F9 可更新。"
            )
        else:
            msg = (
                f"已插入公式 {equation_label}。题注标签“公式”已自动建立；"
                "普通编号与章节公式编号独立计数，按 F9 可更新。"
            )
        return {
            "ok": True,
            "message": msg,
            "emf": str(emf_path),
            "equationLabel": equation_label,
            "numberType": "word-seq",
            "sequenceName": payload["sequenceName"],
        }
    except Exception as exc:
        raise WordFormulaError(f"插入编号公式失败：{exc}") from exc


def read_selected_formula() -> dict[str, Any]:
    """读取 Word 当前选中公式图片的隐藏 JSON。"""
    word = _word_app()
    selection = word.Selection
    _, obj = _selected_object(selection)
    alt_text = _get_alt_text(obj)
    data = decode_payload(alt_text)
    data.pop("svg", None)
    data.pop("svgText", None)  # 前端会根据 LaTeX 重新渲染，避免传超大 SVG。
    return {
        "ok": True,
        "message": "已从 Word 选中公式读取源码和设置。",
        "formula": data,
    }


def update_selected_formula(payload: dict[str, Any]) -> dict[str, Any]:
    """替换 Word 当前选中的旧公式图片。"""
    word = _word_app()
    selection = word.Selection
    kind, obj = _selected_object(selection)

    old_alt = _get_alt_text(obj)
    if not is_formula_payload(old_alt):
        raise WordFormulaError("当前选中对象不是由本工具插入的可编辑公式，不能直接更新。")

    old_data = {}
    try:
        old_data = decode_payload(old_alt)
    except Exception:
        old_data = {}

    payload = dict(payload)
    # 如果旧对象是编号公式，更新时保留编号元数据；右侧 Word SEQ 域不随图片替换而改变。
    for key in (
        "numbered",
        "numberType",
        "sequenceName",
        "sequenceResetByHeadingLevel",
        "includeChapterNumber",
        "includeSectionNumber",
        "chapterNumberMode",
        "chapterNumber",
        "equationIndex",
        "equationLabel",
    ):
        if key in old_data and key not in payload:
            payload[key] = old_data[key]
    # 编号布局不会在“更新图片”时变化，因此编号类型必须以旧对象为准，
    # 避免用户未先点击“读取”时把章节编号误写成普通编号。
    if bool(old_data.get("numbered")):
        payload["numbered"] = True
        payload["includeChapterNumber"] = bool(old_data.get("includeChapterNumber"))
        payload["includeSectionNumber"] = bool(old_data.get("includeSectionNumber"))
        payload["chapterNumberMode"] = str(
            old_data.get("chapterNumberMode") or
            ("chapter-section" if old_data.get("includeSectionNumber") else "chapter")
        )

    _, _, emf_path = _save_formula_files(payload)
    alt_text = _encode_formula_alt(payload)

    try:
        if kind == "inline":
            rng = obj.Range
            obj.Delete()
            rng.Select()
            new_obj = word.Selection.InlineShapes.AddPicture(
                FileName=str(emf_path),
                LinkToFile=False,
                SaveWithDocument=True,
            )
            _set_alt_text(new_obj, alt_text)
            _apply_word_baseline(new_obj, payload)
        else:
            # 浮动图片也支持读取；更新时为避免复杂定位，先删除再在当前位置插入为行内图。
            obj.Delete()
            new_obj = word.Selection.InlineShapes.AddPicture(
                FileName=str(emf_path),
                LinkToFile=False,
                SaveWithDocument=True,
            )
            _set_alt_text(new_obj, alt_text)
            _apply_word_baseline(new_obj, payload)

        return {
            "ok": True,
            "message": "已更新 Word 选中公式。",
            "emf": str(emf_path),
        }
    except Exception as exc:
        raise WordFormulaError(f"更新 Word 公式失败：{exc}") from exc


def find_word_conversion_targets() -> dict[str, Any]:
    """
    读取 Word 当前选区中的 LaTeX 定界符。

    - 非空选区：识别选区内全部完整定界符；
    - 空选区：只识别当前段落中包围光标的一个完整定界符；
    - 无匹配时返回空列表，不修改 Word。
    """
    word = _word_app()
    doc = word.ActiveDocument
    selection = word.Selection
    rng = selection.Range
    selection_start = int(rng.Start)
    selection_end = int(rng.End)
    selection_empty = selection_start == selection_end

    if not selection_empty:
        text = str(rng.Text or "")
        matches = _scan_formula_markers(text, base_start=selection_start)
    else:
        try:
            paragraph_range = rng.Paragraphs.Item(1).Range
            paragraph_start = int(paragraph_range.Start)
            paragraph_text = str(paragraph_range.Text or "")
        except Exception:
            paragraph_start = selection_start
            paragraph_text = ""

        all_matches = _scan_formula_markers(paragraph_text, base_start=paragraph_start)
        enclosing = [
            item for item in all_matches
            if int(item["start"]) <= selection_start <= int(item["end"])
        ]
        # 理论上扫描结果互不重叠；仍按最短范围优先，避免异常文档中的歧义。
        enclosing.sort(key=lambda item: (int(item["end"]) - int(item["start"]), abs(selection_start - int(item["contentStart"]))))
        matches = enclosing[:1]

    if not matches:
        return {
            "ok": True,
            "message": "未找到可转换的 LaTeX 定界符，Word 内容未修改。",
            "matches": [],
            "selectionEmpty": selection_empty,
            "documentIdentity": _document_identity(doc),
        }

    return {
        "ok": True,
        "message": f"已识别 {len(matches)} 个待转换公式。",
        "matches": matches,
        "selectionEmpty": selection_empty,
        "documentIdentity": _document_identity(doc),
    }


def _validate_conversion_item(doc: Any, item: dict[str, Any]) -> tuple[int, int, dict[str, Any], str]:
    try:
        start = int(item.get("rangeStart", item.get("start")))
        end = int(item.get("rangeEnd", item.get("end")))
    except Exception as exc:
        raise WordFormulaError("转换请求缺少有效的 Word 范围。") from exc
    if start < 0 or end <= start:
        raise WordFormulaError("转换请求中的 Word 范围无效。")

    expected_source = str(item.get("sourceText") or "")
    try:
        current_source = str(doc.Range(start, end).Text or "")
    except Exception as exc:
        raise WordFormulaError("无法读取待转换公式在 Word 中的当前位置。") from exc
    if not expected_source or current_source != expected_source:
        raise WordFormulaError("Word 选区内容在渲染期间发生了变化，已取消转换，文档未修改。")

    parsed = _scan_formula_markers(current_source, base_start=start)
    if len(parsed) != 1 or int(parsed[0]["start"]) != start or int(parsed[0]["end"]) != end:
        raise WordFormulaError("待转换文本不再是一个完整且独立的公式定界符，已取消转换。")

    expected_type = str(item.get("formulaType") or item.get("markerType") or "")
    if expected_type and parsed[0]["formulaType"] != expected_type:
        raise WordFormulaError("公式定界符类型发生变化，已取消转换。")
    if str(item.get("latex") or "") != str(parsed[0]["latex"]):
        raise WordFormulaError("公式源码发生变化，已取消转换。")

    payload = _clean_conversion_payload(item)
    payload["latex"] = parsed[0]["latex"]
    payload["mode"] = parsed[0]["mode"]
    payload["display"] = parsed[0]["mode"] == "display"
    payload["includeChapterNumber"] = bool(parsed[0]["includeChapterNumber"])
    return start, end, payload, str(parsed[0]["formulaType"])


def _insert_converted_inline(doc: Any, start: int, end: int, payload: dict[str, Any], emf_path: Path) -> None:
    rng = doc.Range(start, end)
    rng.Text = ""
    insertion = doc.Range(start, start)
    formula_obj = insertion.InlineShapes.AddPicture(
        FileName=str(emf_path),
        LinkToFile=False,
        SaveWithDocument=True,
    )
    _set_alt_text(formula_obj, _encode_formula_alt(payload))
    _apply_word_baseline(formula_obj, payload)


def _insert_converted_display(doc: Any, start: int, end: int, payload: dict[str, Any], emf_path: Path) -> None:
    slot = _prepare_block_insertion_range(doc, doc.Range(start, end))
    try:
        slot.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
        slot.ParagraphFormat.SpaceBefore = 0
        slot.ParagraphFormat.SpaceAfter = 0
    except Exception:
        pass
    formula_obj = slot.InlineShapes.AddPicture(
        FileName=str(emf_path),
        LinkToFile=False,
        SaveWithDocument=True,
    )
    _set_alt_text(formula_obj, _encode_formula_alt(payload))
    _apply_word_baseline(formula_obj, payload)
    try:
        formula_obj.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
    except Exception:
        pass


def _insert_converted_numbered(
    word: Any,
    doc: Any,
    start: int,
    end: int,
    payload: dict[str, Any],
    emf_path: Path,
    *,
    include_chapter: bool,
) -> None:
    payload = dict(payload)
    payload["mode"] = "display"
    payload["display"] = True
    payload["numbered"] = True
    payload["numberType"] = "word-seq"
    include_section = include_chapter and (
        str(payload.get("chapterNumberMode") or "chapter").strip().lower()
        == "chapter-section"
        or _payload_bool(payload.get("includeSectionNumber"))
    )
    payload["sequenceName"] = (
        CHAPTER_EQUATION_SEQUENCE_NAME
        if include_chapter
        else PLAIN_EQUATION_SEQUENCE_NAME
    )
    payload["sequenceResetByHeadingLevel"] = (
        2 if include_section else (1 if include_chapter else None)
    )
    payload["includeSectionNumber"] = include_section
    payload["chapterNumberMode"] = (
        "chapter-section" if include_section else "chapter"
    )
    payload["includeChapterNumber"] = include_chapter

    slot = _prepare_block_insertion_range(doc, doc.Range(start, end))
    table = doc.Tables.Add(Range=slot, NumRows=1, NumColumns=3)
    _format_equation_table(table)

    left_rng = _cell_content_range(table.Cell(1, 1))
    center_rng = _cell_content_range(table.Cell(1, 2))
    try:
        left_rng.Text = ""
    except Exception:
        pass
    try:
        center_rng.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
    except Exception:
        pass

    formula_obj = center_rng.InlineShapes.AddPicture(
        FileName=str(emf_path),
        LinkToFile=False,
        SaveWithDocument=True,
    )
    _set_alt_text(formula_obj, _encode_formula_alt(payload))
    _apply_word_baseline(formula_obj, payload)
    _insert_equation_number_fields(word, table.Cell(1, 3), include_chapter=include_chapter, include_section=include_section)


def convert_word_markers(
    request: dict[str, Any],
    progress_callback: Callable[[str, int, int], None] | None = None,
) -> dict[str, Any]:
    """把前端已渲染的多个公式按 Word 绝对范围逆序原位替换。"""
    raw_items = request.get("items")
    if not isinstance(raw_items, list) or not raw_items:
        return {"ok": True, "message": "没有待转换公式，Word 内容未修改。", "converted": 0}

    word = _word_app()
    doc = word.ActiveDocument
    expected_document = str(request.get("documentIdentity") or "")
    if expected_document and _document_identity(doc) != expected_document:
        raise WordFormulaError("当前活动 Word 文档已变化，已取消转换，文档未修改。")

    validated: list[tuple[int, int, dict[str, Any], str]] = []
    seen_ranges: set[tuple[int, int]] = set()
    for raw_item in raw_items:
        if not isinstance(raw_item, dict):
            raise WordFormulaError("转换请求中的公式项目格式无效。")
        start, end, payload, formula_type = _validate_conversion_item(doc, raw_item)
        key = (start, end)
        if key in seen_ranges:
            raise WordFormulaError("转换请求中存在重复公式范围。")
        seen_ranges.add(key)
        validated.append((start, end, payload, formula_type))

    validated.sort(key=lambda value: value[0], reverse=True)
    for idx in range(len(validated) - 1):
        current_start, _current_end, _current_payload, _current_type = validated[idx]
        _next_start, next_end, _next_payload, _next_type = validated[idx + 1]
        if next_end > current_start:
            raise WordFormulaError("待转换公式范围发生重叠，已取消转换。")

    # 先完成全部 SVG→EMF 转换，再修改 Word，避免中途转换失败造成部分替换。
    prepared: list[tuple[int, int, dict[str, Any], str, Path, Path, Path]] = []
    temp_paths: list[Path] = []
    batch_tag = f"convert_{time.time_ns()}"
    try:
        total = len(validated)
        source_pairs: list[tuple[Path, Path]] = []
        for index, (start, end, payload, formula_type) in enumerate(validated, start=1):
            json_path, svg_path, emf_path = _write_formula_source_files(payload, stem=f"{batch_tag}_{index}")
            temp_paths.extend((json_path, svg_path, emf_path))
            prepared.append((start, end, payload, formula_type, json_path, svg_path, emf_path))
            source_pairs.append((svg_path, emf_path))

        if progress_callback is not None:
            try:
                progress_callback("emf", 0, total)
            except Exception:
                pass

        def report_emf_batch(completed: int, batch_total: int) -> None:
            if progress_callback is not None:
                try:
                    progress_callback("emf", completed, batch_total)
                except Exception:
                    pass

        # 单次 Inkscape shell 批量转换；短 ASCII 临时名 + 显式输出名，缺失项才单独重试。
        svg_batch_to_emf(source_pairs, progress_callback=report_emf_batch)


    except Exception as exc:
        for path in temp_paths:
            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass
        raise WordFormulaError(f"批量生成公式 EMF 失败，Word 内容未修改：{exc}") from exc

    undo_started = False
    try:
        try:
            word.UndoRecord.StartCustomRecord("转换 LaTeX 定界符公式")
            undo_started = True
        except Exception:
            undo_started = False

        if progress_callback is not None:
            try:
                progress_callback("word", len(prepared), len(prepared))
            except Exception:
                pass

        counts = {"inline": 0, "display": 0, "numbered": 0, "chapter-numbered": 0}
        for start, end, payload, formula_type, _json_path, _svg_path, emf_path in prepared:
            if formula_type == "inline":
                _insert_converted_inline(doc, start, end, payload, emf_path)
            elif formula_type == "display":
                _insert_converted_display(doc, start, end, payload, emf_path)
            elif formula_type == "numbered":
                _insert_converted_numbered(word, doc, start, end, payload, emf_path, include_chapter=False)
            elif formula_type == "chapter-numbered":
                _insert_converted_numbered(word, doc, start, end, payload, emf_path, include_chapter=True)
            else:
                raise WordFormulaError(f"不支持的公式定界符类型：{formula_type}")
            counts[formula_type] += 1

        if counts["numbered"] or counts["chapter-numbered"]:
            try:
                doc.Fields.Update()
            except Exception:
                pass

        detail_parts = []
        labels = (("inline", "行内"), ("display", "行间"), ("numbered", "编号"), ("chapter-numbered", "章节编号"))
        for key, label in labels:
            if counts[key]:
                detail_parts.append(f"{label} {counts[key]} 个")
        detail = "、".join(detail_parts)
        return {
            "ok": True,
            "message": f"已转换 {len(validated)} 个公式（{detail}）。",
            "converted": len(validated),
            "counts": counts,
        }
    except Exception as exc:
        # 支持 UndoRecord 的 Word 版本中，尽量把本次批量操作整体回滚。
        if undo_started:
            try:
                word.UndoRecord.EndCustomRecord()
                undo_started = False
                word.Undo()
            except Exception:
                pass
        raise WordFormulaError(f"转换 Word 中的 LaTeX 定界符失败：{exc}") from exc
    finally:
        if undo_started:
            try:
                word.UndoRecord.EndCustomRecord()
            except Exception:
                pass
        for path in temp_paths:
            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass



def _formula_marker_from_payload(data: dict[str, Any]) -> tuple[str, str]:
    latex = str(data.get("latex") or "")
    if not latex.strip():
        raise WordFormulaError("选中的公式缺少 LaTeX 源码，无法回写。")

    if bool(data.get("numbered")):
        if bool(data.get("includeChapterNumber")):
            return f"##{latex}##", "chapter-numbered"
        return f"#{latex}#", "numbered"

    mode = str(data.get("mode") or "").lower()
    if mode not in ("inline", "display"):
        mode = "display" if bool(data.get("display")) else "inline"
    if mode == "display":
        return f"$${latex}$$", "display"
    return f"${latex}$", "inline"


def _object_range_bounds(obj: Any, *, floating: bool = False) -> tuple[int, int]:
    try:
        rng = obj.Anchor if floating else obj.Range
        return int(rng.Start), int(rng.End)
    except Exception as exc:
        raise WordFormulaError("无法读取选中公式在 Word 中的位置。") from exc


def _numbered_formula_table(obj: Any) -> Any:
    try:
        rng = obj.Range
        if int(rng.Tables.Count) < 1:
            raise WordFormulaError("编号公式不在表格中，不能安全删除编号布局。")
        table = rng.Tables.Item(1)
        if int(table.Rows.Count) != 1 or int(table.Columns.Count) != 3:
            raise WordFormulaError("编号公式所在表格不是一行三列结构，已取消回写以避免误删普通表格。")
        return table
    except WordFormulaError:
        raise
    except Exception as exc:
        raise WordFormulaError("无法确认编号公式的一行三列表格结构，已取消回写。") from exc


def _selection_direct_formula_ranges(selection: Any) -> set[tuple[int, int, str]]:
    """记录 Word 明确选中的图片，补充普通范围判断覆盖不到的对象选择状态。"""
    selected: set[tuple[int, int, str]] = set()
    try:
        for index in range(1, int(selection.InlineShapes.Count) + 1):
            obj = selection.InlineShapes.Item(index)
            start, end = _object_range_bounds(obj)
            selected.add((start, end, "inline"))
    except Exception:
        pass
    try:
        for index in range(1, int(selection.ShapeRange.Count) + 1):
            obj = selection.ShapeRange.Item(index)
            start, end = _object_range_bounds(obj, floating=True)
            selected.add((start, end, "shape"))
    except Exception:
        pass
    return selected


def _collect_selected_formula_writebacks(doc: Any, selection: Any) -> list[dict[str, Any]]:
    selection_start = int(selection.Range.Start)
    selection_end = int(selection.Range.End)
    selection_empty = selection_start == selection_end
    direct_ranges = _selection_direct_formula_ranges(selection)
    candidates: list[tuple[str, Any, bool]] = []

    try:
        for index in range(1, int(doc.InlineShapes.Count) + 1):
            candidates.append(("inline", doc.InlineShapes.Item(index), False))
    except Exception:
        pass
    try:
        for index in range(1, int(doc.Shapes.Count) + 1):
            candidates.append(("shape", doc.Shapes.Item(index), True))
    except Exception:
        pass

    targets: list[dict[str, Any]] = []
    seen: set[tuple[str, int, int]] = set()
    for kind, obj, floating in candidates:
        alt_text = _get_alt_text(obj)
        if not is_formula_payload(alt_text):
            continue
        try:
            data = decode_payload(alt_text)
        except Exception:
            continue
        if not isinstance(data, dict) or not str(data.get("latex") or "").strip():
            continue

        obj_start, obj_end = _object_range_bounds(obj, floating=floating)
        explicitly_selected = (obj_start, obj_end, kind) in direct_ranges
        if selection_empty:
            selected = explicitly_selected or obj_start <= selection_start <= obj_end
            if not selected and bool(data.get("numbered")):
                try:
                    table_range = _numbered_formula_table(obj).Range
                    selected = int(table_range.Start) <= selection_start <= int(table_range.End)
                except Exception:
                    selected = False
        else:
            selected = explicitly_selected or (obj_start >= selection_start and obj_end <= selection_end)
        if not selected:
            continue

        marker_text, formula_type = _formula_marker_from_payload(data)
        if bool(data.get("numbered")):
            table = _numbered_formula_table(obj)
            table_start = int(table.Range.Start)
            table_end = int(table.Range.End)
            key = ("table", table_start, table_end)
            if key in seen:
                continue
            seen.add(key)
            targets.append({
                "kind": "table",
                "object": table,
                "start": table_start,
                "end": table_end,
                "markerText": marker_text,
                "formulaType": formula_type,
            })
        else:
            key = (kind, obj_start, obj_end)
            if key in seen:
                continue
            seen.add(key)
            targets.append({
                "kind": kind,
                "object": obj,
                "start": obj_start,
                "end": obj_end,
                "markerText": marker_text,
                "formulaType": formula_type,
            })

    targets.sort(key=lambda item: int(item["start"]), reverse=True)
    for index in range(len(targets) - 1):
        current = targets[index]
        following = targets[index + 1]
        if int(following["end"]) > int(current["start"]):
            raise WordFormulaError("选中的公式范围发生重叠，已取消回写。")
    return targets


def _replace_formula_with_marker(doc: Any, target: dict[str, Any]) -> None:
    start = int(target["start"])
    end = int(target["end"])
    marker_text = str(target["markerText"])
    formula_type = str(target["formulaType"])
    kind = str(target["kind"])

    if kind == "table":
        table = target["object"]
        try:
            table.Delete()
        except Exception as exc:
            raise WordFormulaError("删除编号公式的一行三列表格失败。") from exc
        insertion = doc.Range(start, start)
        insertion.Text = marker_text
        try:
            doc.Range(start, start + _utf16_length(marker_text)).ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_LEFT
        except Exception:
            pass
        return

    if kind == "inline":
        try:
            doc.Range(start, end).Text = marker_text
        except Exception as exc:
            raise WordFormulaError("把公式图片替换为 LaTeX 编码失败。") from exc
    else:
        try:
            target["object"].Delete()
            doc.Range(start, start).Text = marker_text
        except Exception as exc:
            raise WordFormulaError("把浮动公式图片替换为 LaTeX 编码失败。") from exc

    if formula_type == "display":
        try:
            doc.Range(start, start + _utf16_length(marker_text)).ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_LEFT
        except Exception:
            pass


def writeback_selected_formulas() -> dict[str, Any]:
    """把 Word 当前选中的一个或多个公式图片批量回写为对应 LaTeX 定界符文本。"""
    word = _word_app()
    doc = word.ActiveDocument
    selection = word.Selection
    targets = _collect_selected_formula_writebacks(doc, selection)
    if not targets:
        return {
            "ok": True,
            "message": "当前选区中没有由本工具插入的可回写公式，Word 内容未修改。",
            "writtenBack": 0,
        }

    undo_started = False
    try:
        try:
            word.UndoRecord.StartCustomRecord("公式回写为 LaTeX 编码")
            undo_started = True
        except Exception:
            undo_started = False

        counts = {"inline": 0, "display": 0, "numbered": 0, "chapter-numbered": 0}
        for target in targets:
            _replace_formula_with_marker(doc, target)
            counts[str(target["formulaType"])] += 1

        first_start = min(int(item["start"]) for item in targets)
        try:
            doc.Range(first_start, first_start).Select()
        except Exception:
            pass

        details = []
        labels = (("inline", "行内"), ("display", "行间"), ("numbered", "编号"), ("chapter-numbered", "章节编号"))
        for key, label in labels:
            if counts[key]:
                details.append(f"{label} {counts[key]} 个")
        return {
            "ok": True,
            "message": f"已回写 {len(targets)} 个公式（{'、'.join(details)}）。",
            "writtenBack": len(targets),
            "counts": counts,
        }
    except Exception as exc:
        if undo_started:
            try:
                word.UndoRecord.EndCustomRecord()
                undo_started = False
                word.Undo()
            except Exception:
                pass
        if isinstance(exc, WordFormulaError):
            raise
        raise WordFormulaError(f"批量回写公式失败：{exc}") from exc
    finally:
        if undo_started:
            try:
                word.UndoRecord.EndCustomRecord()
            except Exception:
                pass

def check_environment() -> dict[str, Any]:
    """返回 Word/pywin32/Inkscape 环境状态，用于页面诊断。"""
    result: dict[str, Any] = {
        "ok": True,
        "platform": sys.platform,
        "windows": sys.platform.startswith("win"),
        "pywin32": False,
        "wordRunning": False,
        "inkscape": None,
        "formulaWordBuild": FORMULA_WORD_BUILD,
        "formulaWordFile": str(Path(__file__).resolve()),
    }

    try:
        from emf_convert import find_inkscape
        result["inkscape"] = find_inkscape()
    except Exception:
        result["inkscape"] = None

    if sys.platform.startswith("win"):
        try:
            import win32com.client  # type: ignore
            result["pywin32"] = True
            try:
                win32com.client.GetActiveObject("Word.Application")
                result["wordRunning"] = True
            except Exception:
                result["wordRunning"] = False
        except Exception:
            result["pywin32"] = False
    return result
