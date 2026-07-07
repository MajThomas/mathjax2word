# LaTeX SVG/EMF Word 公式小工具，本地版

本版本在原有“LaTeX → SVG → 剪贴板”的基础上，增加了路线 B 的 Word 增强功能：

- 用本地网页和离线 MathJax 把 LaTeX 渲染为 SVG；
- SVG 内继续写入 `<metadata><latex-source>...</latex-source></metadata>`；
- 选择“行内公式”时，渲染送入 MathJax 的源码会自动临时包一层 `{...}`，避免 `+`、`-`、`=` 等符号造成行内解析截断；输入框、SVG metadata 和 Word 隐藏数据仍保存原始源码；
- 后端调用 Inkscape 把 SVG 转为 EMF；
- 后端调用 Microsoft Word，把 EMF 作为图片插入当前光标位置；
- 同时把 LaTeX 源码、颜色、字号、公式字体、行内/行间模式等配置压缩后写入 Word 图片的 `AlternativeText` 隐藏数据；
- 后续在 Word 中选中该公式图片，回到工具点击“读取 Word 选中公式”，即可恢复公式继续编辑；
- 修改后点击“更新 Word 选中公式”，重新生成 EMF 并替换 Word 中选中的旧公式；
- 新增“插入 Word 编号公式”，自动用无边框三列表格插入：中间单元格居中放公式，右侧单元格右对齐放 Word 可更新编号域；
- 可选“编号含一级标题号”，右侧编号使用 `STYLEREF 1 \s` 读取一级标题编号，并使用 `SEQ LatexSvgEq \* ARABIC \s 1` 按一级标题重新开始公式序列；
- 新增置顶迷你窗口：无标题栏、始终置顶，仅保留单行公式输入框、插入行内、插入行间、插入编号、章节编号、读取、字号下拉框、关闭按钮和公式字体下拉框；
- 默认字号改为 `12 pt`，字号选择新增 `10.5 pt`。

> 注意：这是“小工具量级”的实现，不是自定义 OLE Server。它不会注册 COM/OLE 对象，也不能双击公式自动打开编辑器。编辑方式是：**在 Word 中选中公式图片 → 回到网页工具点击读取/更新按钮**。

---

## 一、运行环境

普通 SVG 功能：

- Python 3.9+
- 浏览器

Word 增强功能：

- Windows
- Microsoft Word 桌面版
- `pywin32`
- Inkscape，用于 SVG → EMF

安装 Python 依赖：

```bat
python -m pip install -r requirements.txt
```

或直接：

```bat
python -m pip install pywin32 pywebview
```

其中 `pywebview` 只用于“置顶迷你窗口”；如果只用普通网页模式，可以不安装它。

Inkscape 安装后，工具会自动从 PATH 和常见目录查找 `inkscape.exe`。如果找不到，可以设置环境变量：

```bat
set INKSCAPE_PATH=C:\Program Files\Inkscape\bin\inkscape.exe
python run_local_server.py
```

---

## 二、启动方法

Windows 直接双击：

```bat
start_server.bat
```

或命令行运行：

```bat
python run_local_server.py
```

浏览器打开：

```text
http://localhost:8000/latex-svg-clipboard.html
```

如需使用始终置顶的迷你窗口，推荐双击无控制台启动脚本：

```text
start_mini_window_silent.vbs
```

也可以双击：

```bat
start_mini_window.bat
```

该 bat 会转发到 `start_mini_window_silent.vbs`，不会保留控制台窗口。调试时可手动运行：

```bat
python start_mini_window.py
```

迷你窗口会自动启动本地服务，并打开：

```text
http://localhost:8000/latex-svg-clipboard.html?mini=1
```

---

## 三、插入 Word 可编辑公式

1. 打开 Word 文档，把光标放到要插入公式的位置。
2. 打开本工具网页。
3. 输入 LaTeX 公式。
4. 设置公式模式、颜色、字号、公式字体。选择“行内公式”时，工具会在渲染解析阶段自动临时使用 `{公式源码}`，但不会改动输入框内容。
5. 点击“插入 Word 可编辑公式”。如果点击“插入 Word 行间公式”，工具会把公式按独立居中段落插入，且不生成编号。

工具会执行：

```text
LaTeX → MathJax SVG → Inkscape EMF → Word 插入 EMF 图片 → AlternativeText 写入公式数据
```

Word 中默认显示的是 EMF 矢量图片，保存 Word 后隐藏数据会随图片一起保存。


---

## 四、插入 Word 编号公式

1. 打开 Word 文档，把光标放到要插入编号公式的位置。
2. 输入 LaTeX 公式。
3. 如需章节号，勾选“编号含一级标题号”。
4. 点击“插入 Word 编号公式”。

工具会在 Word 中插入一个无边框三列表格：

```text
┌──────────┬────────────────────────────┬──────────┐
│          │          公式居中           │     (1)  │
└──────────┴────────────────────────────┴──────────┘
```

如果勾选章节号，编号形式为：

```text
(章节号.公式序号)
```

例如当前光标位于一级标题“2 试验方法”之后，则依次生成：

```text
(2.1)、(2.2)、(2.3)
```

说明：编号现在是 Word 动态域，不再是普通文本。普通编号使用：

```text
{ SEQ LatexSvgEq \* ARABIC }
```

勾选章节号后使用：

```text
{ STYLEREF 1 \s }.{ SEQ LatexSvgEq \* ARABIC \s 1 }
```

其中 `\s 1` 表示公式序列按一级标题重新开始。移动、删除或新增编号公式后，可在 Word 中按 `Ctrl+A` 全选，再按 `F9` 更新域；部分笔记本可能需要 `Fn+F9`。

---

## 五、置顶迷你窗口

迷你窗口用于边写 Word 边快速插公式。窗口特性：

```text
始终置顶
无系统标题栏
单行公式输入框
插入行内公式
插入行间公式
插入编号公式
插入带章节编号公式
读取公式
字号下拉框，默认 12 pt，含 10.5 pt 选项
关闭按钮，仅隐藏窗口，不关闭后台
公式字体下拉框，可直接在迷你窗口内切换字体
```

启动：

```text
start_mini_window_silent.vbs
```

或：

```bat
start_mini_window.bat
```

迷你窗口外层尺寸默认约为 `1240 × 40`。如果你的系统或 pywebview 后端仍保留额外空白，可通过环境变量微调：

```bat
set LATEX_SVG_MINI_HEIGHT=40
set LATEX_SVG_MINI_WIDTH=1240
python start_mini_window.py
```

迷你窗口默认保留普通网页中的设置，例如颜色、公式字体等。字号默认 `12 pt`，下拉框含 `10.5 pt` 选项，选择后会立即更新当前公式渲染设置；右侧公式字体下拉框选择后会保存设置并重新加载迷你页面，使后续插入公式使用所选字体。按回车默认执行“插入行内公式”。点击“插入行间公式”会临时按行间模式生成并插入，不影响编号公式功能。点击“插入编号”插入不带章节号的 Word `SEQ` 编号公式；点击“章节编号”插入带一级标题章节号、且每章重新开始的编号公式。点击“读取公式”会读取 Word 当前选中的可编辑公式，并回填到迷你输入框和字号下拉框。

迷你窗口右下角会常驻系统托盘图标：点击托盘图标可重新唤出迷你窗口；右键托盘图标可选择“显示迷你窗口”或“退出后台”。迷你窗口里的“关闭”按钮只隐藏窗口，不关闭后台。

---

## 六、从 Word 回读并继续编辑

1. 在 Word 中单击选中之前由本工具插入的公式图片。
2. 回到工具网页。
3. 点击“读取 Word 选中公式”。
4. 工具会恢复 LaTeX 源码、颜色、字号、公式字体、行内/行间模式。
5. 修改公式。
6. 点击“更新 Word 选中公式”。

---

## 七、旧版 SVG 回读方式仍然保留

如果 Word 中是旧版本 SVG 公式，而不是新版 EMF 可编辑公式，可以继续使用原方式：

1. 在 Word 中选中原 SVG 公式并复制。
2. 回到网页按 `Ctrl+V`。
3. 网页会尽量从剪贴板中的 SVG / HTML / data URL 中读取 `<latex-source>`。

也可以把 `.svg` 文件直接拖到网页的回读区域。

---

## 八、主要文件说明

```text
latex-svg-clipboard.html   # 前端页面：MathJax 渲染、SVG metadata、Word 操作按钮
run_local_server.py        # 本地 HTTP 服务：静态文件 + Word API
formula_payload.py         # 公式 JSON 压缩/解压，写入或读取 AlternativeText
emf_convert.py             # 调用 Inkscape 把 SVG 转 EMF
formula_word.py            # 调用 Word COM 插入、读取、更新公式、插入编号公式
start_mini_window.py       # 启动置顶无标题栏迷你窗口，并创建系统托盘图标
start_mini_window_silent.pyw # 无控制台入口
start_mini_window_silent.vbs # 推荐双击的无控制台启动脚本
start_mini_window.bat      # Windows 迷你窗口启动脚本，转发到 vbs
temp/                      # 临时保存 last_formula.json / .svg / .emf
vendor/mathjax/            # 离线 MathJax
fonts-manifest.json        # MathJax 公式字体配置
requirements.txt           # Python 依赖
```

---

## 九、隐藏数据格式

Word 图片的 `AlternativeText` 中写入类似下面的字符串：

```text
LATEX_SVG_FORMULA_V1:<base64-zlib-json>
```

解压后的 JSON 大致为：

```json
{
  "type": "latex-svg-formula",
  "version": 1,
  "latex": "\\frac{a}{b}",
  "mode": "display",
  "display": true,
  "fontSize": 18,
  "color": "#000000",
  "mathFont": "mathjax-newcm",
  "mathFontLabel": "MathJax New Computer Modern（内置）"
}
```

SVG 较小时也会作为备份写入隐藏数据；如果 SVG 过大，后端会自动省略 SVG，只保留 LaTeX 与配置。重新编辑时会重新渲染 SVG。

---

## 十、常见问题

### 1. 提示“缺少 pywin32”

运行：

```bat
python -m pip install pywin32
```

### 2. 提示“未找到 Inkscape”

安装 Inkscape，或设置：

```bat
set INKSCAPE_PATH=C:\Program Files\Inkscape\bin\inkscape.exe
```

### 3. 提示“没有检测到正在运行的 Word”

先打开 Word 文档，然后再点击插入、读取或更新。

### 4. 提示“当前选中对象不是由本工具插入的可编辑公式”

说明当前选中的图片没有本工具写入的 `AlternativeText` 隐藏数据。需要选中由“插入 Word 可编辑公式”按钮插入的公式。

### 5. 更新后公式变成行内图

当前路线 B 的更新逻辑优先支持行内图片。浮动图片可以读取，但更新时会重新插入为行内图片，避免复杂定位逻辑。

### 6. 迷你窗口打不开

先确认安装了 `pywebview`：

```bat
python -m pip install pywebview
```

Windows 上通常会使用系统 Edge WebView2 运行网页窗口。如果系统缺少 WebView2 Runtime，请安装 Microsoft Edge WebView2 Runtime。

### 7. 章节编号显示不正确

带章节编号的公式使用 Word 域 `{ STYLEREF 1 \s }` 读取最近的一级标题编号。请确认章节标题使用了 Word 的“标题 1 / Heading 1”样式，并且该标题已经应用多级列表编号。修改标题或移动公式后，可按 `Ctrl+A` 后按 `F9` 更新全文域。

---

## 十一、公式字体说明

MathJax 的 SVG 输出是基于 MathJax 公式字体生成的矢量路径，不是普通 CSS 字体。因此，仅设置 `font-family` 不能真正改变公式字形。

当前包默认内置：

- `mathjax-newcm`：MathJax New Computer Modern

如需扩展字体，把对应 MathJax 字体包放到：

```text
vendor/fonts/
```

目录结构示例：

```text
vendor/fonts/mathjax-stix2-font/
```

并在 `fonts-manifest.json` 中登记：

```json
[
  {
    "id": "mathjax-newcm",
    "label": "MathJax New Computer Modern（内置）",
    "builtin": true
  },
  {
    "id": "mathjax-stix2",
    "label": "MathJax STIX Two",
    "path": "./vendor/fonts/mathjax-stix2-font",
    "builtin": false
  }
]
```

切换公式字体后，页面会自动刷新一次，因为 MathJax 字体需要在 MathJax 初始化前指定。
