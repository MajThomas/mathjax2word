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


def _save_formula_files(payload: dict[str, Any]) -> tuple[Path, Path, Path]:
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    svg = payload.get("svgText") or payload.get("svg")
    if not isinstance(svg, str) or "<svg" not in svg:
        raise WordFormulaError("前端没有传入有效 SVG，无法生成 EMF。")
    LAST_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    LAST_SVG.write_text(svg, encoding="utf-8")
    svg_to_emf(LAST_SVG, LAST_EMF)
    return LAST_JSON, LAST_SVG, LAST_EMF


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
    alt_text = encode_payload(payload)
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
    alt_text = encode_payload(payload)

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

    _, _, emf_path = _save_formula_files(payload)
    alt_text = encode_payload(payload)

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
