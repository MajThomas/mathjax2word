# Mathjax2word




- 用本地网页和离线 MathJax 把 LaTeX 渲染为 SVG
- 后端调用 Inkscape 把 SVG 转为 EMF；
- 后端调用 Microsoft Word，把 EMF 作为图片插入当前光标位置；
- 同时把 LaTeX 源码、颜色、字号、公式字体、行内/行间模式等配置压缩后写入 Word 图片的 `AlternativeText` 隐藏数据；
- 后续在 Word 中选中该公式图片，回到工具点击“读取 Word 选中公式”，即可恢复公式继续编辑；

---

## 一、安装配置

- Python 3.9+
- Windows
- Microsoft Word 桌面版
- `pywin32`
- Inkscape

**安装Python:**
https://www.python.org/downloads/


安装 Python 依赖：

```bat
python -m pip install -r requirements.txt
```
或直接：

```bat
python -m pip install pywin32 pywebview
```

**安装Inkscape:**
[https://inkscape.org/](https://inkscape.org/zh-hans/release/)

Inkscape 安装后，工具会自动从 PATH 和常见目录查找 `inkscape.exe`。如果找不到，可以设置环境变量：
```bat
set INKSCAPE_PATH=C:\Program Files\Inkscape\bin\inkscape.exe
python run_local_server.py
```
---

## 二、启动方法

Windows 直接打开：

```bat
start_mini_window.bat
```

**迷你窗口会自动启动本地服务，创建后台并在任务栏放置小图标。**

点击关闭按钮可以临时关闭迷你窗口，点击任务栏图标可重新唤起。
关闭时直接右键任务栏图标右键退出


---

## 三、插入 Word 可编辑公式

1. 打开 Word 文档，把光标放到要插入公式的位置。
2. 打开本工具网页。
3. 输入 LaTeX 公式，选择公式格式。
4. 点击相应按钮直接插入公式。

## 四、编号功能说明

带编号的公式插入在word表格中，形式如下：

```text
┌──────────┬────────────────────────────┬──────────┐
│          │          公式居中           │     (1)  │
└──────────┴────────────────────────────┴──────────┘
```

插入的带章节编号公式，编号形式为：

```text
(章号.公式序号)
```

```text
(章号.节号.公式序号)
```
其中章号是光标当前所属的**一级标题**，节号是光标当前所属的**二级标题**，标题为中文时，可以自动转换为阿拉伯数字。

**插入的公式编号可以选择后按F9进行更新。**
编号插入后，**普通编号公式**可以在**交叉引用**的**公式**引用类型中引用。
编号插入后，会自动创建引用类型，**章节编号公式**可以在**交叉引用**的**章节公式**引用类型中引用。




## 五、从 Word 回读当前选中公式并继续编辑

1. 在 Word 中单击选中之前由本工具插入的公式图片。
3. 点击mini窗口“读取”。
3. 工具会恢复 LaTeX 源码、颜色、字号、公式字体、行内/行间模式。
此功能可以用作格式刷，读取当前公式的格式设置，并通过**转换/回写**功能批量把选择的公式转化成读取的格式。

---

## 八、批量生成公式/回写公式为LaTeX文本
在word中可以直接以LaTeX语法撰写公式，并且用以下方式声明公式类型：

行内公式：

```
$Expression$
```

行间公式：

```
$$Expression$$
```

编号公式：

```
#Expression#
```

章节编号公式：

```
##Expression##
```
编辑完成后，将光标插入表达式文本中，或框选所有需要转换的公式，单击**转换/回写**按钮，便可以将LaTeX格式公式转换为相应的公式。
同理，可以选中所有想转换成LaTeX文本的公式，单击**转换/回写**按钮，程序将自动把公式回写为LaTeX文本。
## 九、自定义字体


当前包默认可以在线调用所有MathJax官方支持的字体。其中New Computer Modern和STIX Two Math字体可以离线使用。

如需扩展字体，或离线使用非默认字体，把对应 MathJax 字体包放到：

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
...
]
```

由于不同字体的基线不同，每个字体需要单独设置公式偏移，以保证渲染的图片与正文文本对齐。发布版本已经对New Computer Modern和STIX Two Math字体设置好了字体偏移，其它字体默认采用New Computer Modern，可能存在公式不对齐的情况。字体对齐也通过`fonts-manifest.json`文件进行配置：

```
"fonts": [
    {
      "id": "mathjax-newcm",            //字体id
      "label": "New Computer Modern",   //字体名
      "builtin": false,
      "calibrated": true,
      "calibrationStatus": "calibrated",
      "inlineBaselineOffsetCurvePt": {  //字体偏移量
        "10.5": 3.3,                    //10.5pt 字号对应的偏移量，单位为pt，向下为正，向上为负
        "12": 4.1,                      //12pt 字号对应的偏移量
        "18": 6.5,                      //18pt 字号对应的偏移量
        "24": 8.5                       //24pt 字号对应的偏移量
      }                                 //其他字号的偏移量根据已设置值进行线性外推得到。偏移量可以是任意小数
    },
]
```
