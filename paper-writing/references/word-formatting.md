# Word 字体与三线表实现

这份参考说明中文字体、标题字重和学术三线表的实现，防止样式声明与实际效果不一致。它不替代用户模板，也不自动修复未知文件。

## 字体必须写到正确的位置

Word 同一个文本运行允许分别指定中文与西文字体。`python-docx` 的 `font.name` 主要读写 ASCII 与 High ANSI 槽位，不能用它单独证明中文字体已经设好。参见 [python-docx 字体说明](https://github.com/python-openxml/python-docx/blob/master/docs/dev/analysis/features/text/font.rst) 和 [Microsoft RunFonts 文档](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts)。

1. 优先在段落样式中设置字体。一级标题对应 Heading 1，二级对应 Heading 2，三级对应 Heading 3；模板采用自定义样式时保留映射及大纲级别。
2. 明确写入 `w:rFonts/@w:eastAsia="黑体"`；Windows 字体族 SimHei 对应“黑体”。宋体对应 SimSun。根据机器实际字体注册名选择等价名称，不能用“加粗”替代黑体字体。
3. 西文的 `w:ascii`、`w:hAnsi` 按本次配置设置。清除当前配置不再使用的主题字体属性，避免遗留的主题配置和兼容性差异；不要把中文和西文的字体选择混为一个字段。
4. 将样式赋给段落，清理段落文本中冲突的直接字体和字号，让样式保持可统一修改。保留有意义的斜体、上下标及数学对象；字符样式若定义了字体，也应检查是否需要调整。不得逐个 run 固定同一字体来掩盖样式失效。
5. 保存后重新读取实际 DOCX，核对样式、段落绑定及运行级覆盖。字体声明正确与实际渲染正确分别检查。无法取得用户原文件时，只能报告机制验证，不能声称已找到或修复该文件的具体问题。

默认标题使用黑色，不添加下划线或装饰性段落边框；检查模板自带的 Title/Heading 样式是否残留蓝色下划线或 `pBdr`。模板明确要求这些效果时保留。字体与字重分开设置：中文题目、一级和二级标题为黑体加粗，中文三级标题为黑体常规字重；英文标题按英文配置加粗，不额外继承无关的斜体等效果。

可复用 [scripts/word_styles.py](../scripts/word_styles.py) 中的函数（依赖 `python-docx`）。以下片段放入任务自己的生成脚本；使用环境提供的 Python 解释器执行，`SKILL_DIR` 替换为本技能实际目录：

```python
import sys
sys.path.insert(0, str(SKILL_DIR / 'scripts'))
from word_styles import configure_body_style, set_style_fonts, apply_paragraph_style, three_line_table

# 此处 12 pt 是通用小四号；若有效配置另有字号，传入该数值。
body_style = configure_body_style(document, size_pt=12)
apply_paragraph_style(body_paragraph, body_style)

# 中文默认配置；模板覆盖时传入相应字号和字重。
for name, size, bold in [('Title', 22, True), ('Heading 1', 16, True),
                         ('Heading 2', 14, True), ('Heading 3', 12, False)]:
    set_style_fonts(document.styles[name], east_asia='黑体',
                    latin='Times New Roman', size_pt=size, bold=bold)
# 在此继续设置对齐、段落间距、大纲级别等有效配置。
heading = document.styles['Heading 1']
apply_paragraph_style(paragraph, heading, clear_bold_override=True)
three_line_table(table, header_rows=1)
```

`set_style_fonts` 的 `bold` 参数显式设置常规与复杂文字的加粗属性；省略时保留原字重，兼容仅设置字体的调用。`apply_paragraph_style` 默认清除当前段落的直接字体与字号并保留强调；仅在统一标题字重时传入 `clear_bold_override=True`，清除标题文本中的直接加粗或取消加粗，让其继承标题样式。不要把该选项无差别用于含强调的正文。

辅助函数不会自动审计整个文档的字符样式、编号字体、页眉页脚或字体安装情况。标题的字符样式及编号若另设字重，也要按有效配置核对。数学、特殊符号和需要特殊字体的段落应分别处理。保存后确认 Title、Heading 1、Heading 2 的有效加粗状态及全部同级标题的实际效果。

## 正文小四号须按数值验证

默认中文正文使用宋体 12 pt，四号对应 14 pt。`configure_body_style` 会明确设置 Normal、存在的 Body Text 及选定正文样式；调用者仍须将正文段落实际绑定到返回的样式，并清除冲突的直接字号。模板使用自定义正文样式时传入其名称，字号参数来自本次有效配置。

`set_style_fonts` 同步设置常规字号和复杂文字字号，避免遗留 `szCs` 引入不同字号。12 pt 在 OOXML 中保存为半磅值 24。检查保存后的正文样式及文本运行覆盖，并在 Word 中选中普通正文验证“宋体、12”；标题和题注按各自样式检查。只读取 `document.styles['Normal'].font.size` 不能证明所有正文已实际变成 12 pt。

## 三线表须设置实际边框

使用技能的有效三线表配置覆盖通用文档技能的网格、底色和表头配色预设，避免两套规则互相改写。用户或学校模板明确要求其他形式时，按模板执行。

辅助函数 `three_line_table` 对普通未合并表格清理旧样式、表级与单元格级边框和底色，在首行顶端、完整表头下方、末行底端分别画线，并将表头设为重复。默认线宽来自格式参数表；可用参数覆盖。

已合并的复杂表头、纵向合并单元格、需要局部辅助线的表格应按实际网格单独排版。辅助函数检测到合并单元格会明确报错，避免错误画线；不能为了套用函数取消合并或损坏表格数据。

Word 编辑视图中可能出现不打印的辅助网格线。判定三线表是否合格时，区分辅助网格与真实边框，使用打印预览或渲染效果核对。默认无竖线的要求针对实际输出边框。

## 标题方形编辑标记与跨机器显示

用户选择的默认行为是：**正常阅读时隐藏编辑标记，保留标题分页控制。** 不要求开启“¶ 显示编辑标记”后仍隐藏所有分页属性提示。

Word 标题左侧的小黑方块可能提示该段落具有“与下段同页”“段中不分页”“段前分页”等属性；这些属性有助于控制标题与正文分页。首先检查方块是否存在于标题文字中、是否为真正的项目符号，以及关闭编辑标记和进入打印预览后是否消失，不能仅凭外观判断。

- 仅在编辑标记模式显示、正常阅读及打印预览中不显示：视为非打印标记，保留必要的 `keepNext`、`keepLines`、`pageBreakBefore` 等分页属性和标题层级。用户没有要求取消分页时，不通过全部关闭分页属性消除该提示。
- 正常阅读或打印预览中仍显示：检查标题文本中实际的方块字符、运行中的符号、项目符号及编号定义。修正误加项，同时保留有效的标题编号；不无差别删除方块字符或全部编号。
- 确实无法判断：说明缺少原 DOCX 或视图证据，不能把显示偏好问题当作已确诊的文档错误。

交付前用“开始 → ¶ 显示/隐藏”关闭编辑标记，查看目标文档。若仍显示，检查“文件 → 选项 → 显示”中“始终在屏幕上显示这些格式标记”及“显示所有格式标记”；需要更改这些个人选项时向用户说明操作，而不擅自批量修改全局选项。参见 [Microsoft 显示或隐藏格式标记说明](https://support.microsoft.com/en-us/word/show-or-hide-tab-marks-in-word)。

如果已通过 Word 对象模型打开本次目标文档，可在该窗口使用 `document.ActiveWindow.View.ShowAll = False` 和 `document.ActiveWindow.View.ShowParagraphs = False`，再检查正常阅读效果。这属于当前阅读视图设置，不能冒充可在所有收件人电脑上强制生效的 DOCX 内容设置。避免切换不相关的活动窗口，避免改写 Normal.dotm 或注册表；不向文档注入宏以控制收件人的显示偏好。参见 [ShowAll](https://learn.microsoft.com/en-us/office/vba/api/word.view.showall) 与 [ShowParagraphs](https://learn.microsoft.com/en-us/office/vba/api/word.view.showparagraphs)。

跨电脑交付时，以正文内容、样式、分页和打印输出正确为文件验收条件，并提醒：收件人主动开启编辑标记后仍可能看到非打印分页标记，关闭 ¶ 可恢复正常阅读。不能承诺修改 skill 后无论收件人采用何种显示设置都看不到编辑标记。
