from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path
from typing import Any

from formula_payload import decode_payload, encode_payload, is_formula_payload
from emf_convert import svg_to_emf

ROOT = Path(__file__).resolve().parent
TEMP_DIR = ROOT / "temp"
LAST_JSON = TEMP_DIR / "last_formula.json"
LAST_SVG = TEMP_DIR / "last_formula.svg"
LAST_EMF = TEMP_DIR / "last_formula.emf"

# Word COM 常量，直接使用数字避免依赖 win32com.constants 初始化。
WD_FIELD_EMPTY = -1
WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
EQUATION_SEQUENCE_NAME = "LatexSvgEq"

# 四类 Word 文本定界符。必须按“长定界符优先”识别，避免 $ 与 $$、# 与 ## 互相串扰。
FORMULA_MARKER_SPECS = (
    {"marker": "##", "formulaType": "chapter-numbered", "mode": "display", "includeChapterNumber": True},
    {"marker": "$$", "formulaType": "display", "mode": "display", "includeChapterNumber": False},
    {"marker": "#", "formulaType": "numbered", "mode": "display", "includeChapterNumber": False},
    {"marker": "$", "formulaType": "inline", "mode": "inline", "includeChapterNumber": False},
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


def _save_formula_files(payload: dict[str, Any], *, stem: str | None = None) -> tuple[Path, Path, Path]:
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

    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    svg_path.write_text(svg, encoding="utf-8")
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


def _insert_equation_number_fields(word: Any, cell: Any, *, include_chapter: bool) -> str:
    r"""
    在右侧单元格插入可更新的 Word 域编号。

    普通编号：        ({ SEQ LatexSvgEq \* ARABIC })
    带章节编号：      ({ STYLEREF 1 \s }.{ SEQ LatexSvgEq \* ARABIC \s 1 })

    说明：SEQ 的 \s 1 会按一级标题重新开始编号；STYLEREF 1 \s 读取当前一级标题编号。
    """
    right_rng = _cell_content_range(cell)
    try:
        right_rng.Text = ""
        right_rng.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
    except Exception:
        pass

    # 用 Selection 插入域比手动构造字段字符更稳定。插入完成后恢复到表格后方。
    right_rng.Select()
    selection = word.Selection
    selection.TypeText("(")

    chapter_text = ""
    if include_chapter:
        chapter_field = _insert_field_at_selection(selection, "STYLEREF 1 \\s")
        chapter_text = _field_result_text(chapter_field)
        selection.TypeText(".")
        seq_field_code = f"SEQ {EQUATION_SEQUENCE_NAME} \\* ARABIC \\s 1"
    else:
        seq_field_code = f"SEQ {EQUATION_SEQUENCE_NAME} \\* ARABIC"

    seq_field = _insert_field_at_selection(selection, seq_field_code)
    seq_text = _field_result_text(seq_field)
    selection.TypeText(")")

    try:
        cell.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
    except Exception:
        pass

    if include_chapter:
        return f"({chapter_text}.{seq_text})" if chapter_text and seq_text else "Word 域编号"
    return f"({seq_text})" if seq_text else "Word 域编号"


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


def insert_numbered_formula(payload: dict[str, Any]) -> dict[str, Any]:
    """以“无边框三列表格”形式插入居中公式和右侧 Word SEQ 域编号。"""
    word = _word_app()
    selection = word.Selection
    doc = word.ActiveDocument

    include_chapter = bool(payload.get("includeChapterNumber"))

    payload = dict(payload)
    payload["mode"] = "display"
    payload["display"] = True
    payload["numbered"] = True
    payload["numberType"] = "word-seq"
    payload["sequenceName"] = EQUATION_SEQUENCE_NAME
    payload["sequenceResetByHeadingLevel"] = 1 if include_chapter else None
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

        equation_label = _insert_equation_number_fields(word, table.Cell(1, 3), include_chapter=include_chapter)

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

        if include_chapter:
            msg = f"已插入编号公式 {equation_label}。编号为 Word 域：STYLEREF 读取一级标题号，SEQ 按一级标题重新开始。"
        else:
            msg = f"已插入编号公式 {equation_label}。编号为 Word SEQ 域，可在 Word 中更新域刷新编号。"
        return {
            "ok": True,
            "message": msg,
            "emf": str(emf_path),
            "equationLabel": equation_label,
            "numberType": "word-seq",
            "sequenceName": EQUATION_SEQUENCE_NAME,
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
        else:
            # 浮动图片也支持读取；更新时为避免复杂定位，先删除再在当前位置插入为行内图。
            obj.Delete()
            new_obj = word.Selection.InlineShapes.AddPicture(
                FileName=str(emf_path),
                LinkToFile=False,
                SaveWithDocument=True,
            )
            _set_alt_text(new_obj, alt_text)

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
    payload["sequenceName"] = EQUATION_SEQUENCE_NAME
    payload["sequenceResetByHeadingLevel"] = 1 if include_chapter else None
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
    _insert_equation_number_fields(word, table.Cell(1, 3), include_chapter=include_chapter)


def convert_word_markers(request: dict[str, Any]) -> dict[str, Any]:
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
    prepared: list[tuple[int, int, dict[str, Any], str, Path]] = []
    temp_paths: list[Path] = []
    batch_tag = f"convert_{time.time_ns()}"
    try:
        for index, (start, end, payload, formula_type) in enumerate(validated, start=1):
            json_path, svg_path, emf_path = _save_formula_files(payload, stem=f"{batch_tag}_{index}")
            temp_paths.extend((json_path, svg_path, emf_path))
            prepared.append((start, end, payload, formula_type, emf_path))
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

        counts = {"inline": 0, "display": 0, "numbered": 0, "chapter-numbered": 0}
        for start, end, payload, formula_type, emf_path in prepared:
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
