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
(章节号.公式序号)
```
其中章节号是光标当前所属的**一级标题**

**插入的公式编号可以选择后按F9进行更新。**

普通编号的域代码：```text
{ SEQ LatexSvgEq \* ARABIC }
```
章节编号的域代码：```text
{ STYLEREF 1 \s }.{ SEQ LatexSvgEq \* ARABIC \s 1 }
```


## 五、从 Word 回读并继续编辑

1. 在 Word 中单击选中之前由本工具插入的公式图片。
3. 点击mini窗口“读取”。
3. 工具会恢复 LaTeX 源码、颜色、字号、公式字体、行内/行间模式。
4. 在窗口继续修改公式。
5. 重新插入公式。

---


## 六、自定义字体


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
