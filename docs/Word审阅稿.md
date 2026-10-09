# 可编辑 Word 稿

项目默认从 LaTeX 源码生成 Word 段落、标题、目录、表格与公式，再应用学校样式。正文与声明正文可以直接输入、删除、替换和换行，表格行高随内容扩展。

## 导出

需要 Windows、Microsoft Word 桌面版、Pandoc 3.x（加入 PATH）、TeX Live 和模板指定字体。当前在 Pandoc 3.10.2、Word 16.0 上核验。

```powershell
python -m pip install -r requirements.txt
python scripts/build.py --check --word
```

输出为 `build/main.pdf` 和 `build/main-review.docx`。也可运行以下命令，先构建当前论文并导出到指定位置：

```powershell
python scripts/export_word.py --output "论文-第1轮审阅.docx"
```

已有本次构建的 PDF 时，可显式指定源码与 PDF：

```powershell
python scripts/export_word.py --source main.tex --pdf build/main.pdf --output "论文-第2轮审阅.docx"
```

源码有更新时先重新构建。封面图、TikZ 图和引用编号需要与当前源码一致的 PDF、AUX 和 BBL。目录通过本机 Word 更新并渲染，临时结果写入 `tmp/source-word-*`。

重新导出时，只有与导出记录哈希一致的原始文件可以被覆盖。已有正文修改、批注、修订或缺少导出记录的目标文件会被保留，此时使用新文件名。

转换随附的自习室预约系统示例，在完整仓库根目录执行：

```powershell
python scripts/export_word.py --source examples/simple-thesis/main.tex --pdf examples/simple-thesis/build/main.pdf --output output/docx/ECJTU_示例论文_自习室预约系统.docx
```

## 编辑与版式

- A4，上下左页边距 25 mm、右 20 mm；正文小四宋体，西文与数字 Times New Roman，首行缩进两字，固定行距 19.5 磅。
- 三级标题使用 Word 标题样式；章标题另起一页。目录使用原生 TOC 域，右键“更新域”可更新条目和页码。
- 正文以完整段落排版，表格为实际单元格。题注可编辑，图题在下、表题在上。
- 封面底图和流程图为图片；封面填写项、声明、题注、正文和表格文字可编辑。源码数学内容由 Pandoc 输出为原生 Word 公式。
- 奇偶页页眉、页脚为独立区域，正文页码从 1 连续编号。前置页面隐藏页码。
- 参考文献采用黑色文字、五号字体和悬挂缩进，网址保留可点击链接。

在 Word 中通过“审阅 → 新建批注”或 `Ctrl+Alt+M` 添加意见。建议每轮审阅用独立文件名。Word 中的修改需要由作者同步到 LaTeX 源文件。

为接近源 PDF 的行尾，部分段落含手动换行。修改这些行的文字后可能出现字距拉大或短行；需要删除或调整相应换行并重新检查排版。可编辑测试检查输入、替换、保存和显示是否成功，不保证编辑后的版式自动保持一致。

Word 与 XeLaTeX 的标点压缩、字距和换行机制不同，部分行尾和分页会有差别。章节、图表和引用编号沿用本次 LaTeX 构建的结果；改变章节结构、图表或文献后重新构建、导出。已核验随附的 27 页模板，以及自习室预约系统样稿中的 3 张表、14 个原生数学对象、3 条参考文献和 3 个代码块。作者宏、`tabularx` 自定义列、`longtable` 数据行和 `lstinputlisting` 完整代码已经加入转换及回归检查；其他特殊环境仍需检查内容、编号和逐页显示。

2026 年 10 月 9 日的样稿实测通过 795 项内容与结构检查、310 项实际编辑检查。严格 PDF 对照未通过：LaTeX 为 19 页，Word 为 18 页；封面和声明后的两张空白页未保留，目录由一页变为两页，表格列宽、边框及部分换行也不同。这些结果证明该样稿可以编辑，尚不能证明转换后的版式与 LaTeX 完全一致。

## 只有 PDF 时

独立 PDF 可使用近似转换入口：

```powershell
python scripts/export_word.py --pdf "论文.pdf" --output "论文-PDF转换.docx"
```

这条命令需要 Python 依赖，按 PDF 推断段落和表格；不恢复学校标题样式、自动目录或完整公式语义。复杂图形默认以 300 dpi 保留，可用 `--dpi 600` 调整。扫描件、无法解码的文字或内容核对失败时停止导出。

有本项目源码时应使用上面的源码导出命令。

## 核验

源码导出在替换目标前检查源文件与 PDF 哈希、完整源段落及段落边界、LaTeX 目录与 Word 标题顺序、学校样式、TOC 域、页眉页脚、表格、图形尺寸与图片数据。输入 Word 前的源 AST 用作内容对照。最终加好书签的文件重新打开渲染，检查文字颜色并输出逐页 PNG。

同名 `.map.json` 和 `.check.json` 用于作者核验，发给老师只需 `.docx`。

实际编辑验证：

```powershell
python scripts/verify_word.py build/main-review.docx --output tmp/word-final-qa
```

程序使用独立 Word 实例和临时副本，在正文页面的原文中输入中英文，执行替换、新段落输入、长表格内容输入、原生公式修改、目录更新和批注，保存、关闭后重新打开。核验新增文字的可见 PDF 绘制结果，并输出全部页面图片。原始交付文件保持原样。公式初始内容按 Word 自身导出的 OMML 文本检查，避免 `Range.Text` 改写数学字母和减号造成误报。

回归测试：

```powershell
python -m unittest discover -s scripts -p "test_*word*.py"
```

覆盖错误页边距、主题覆盖标题字体、题注斜体、段落合并、额外标题编号、图形缩小、源文件变更，以及透明/白色文字、图片损坏、封面缩放、已有修改和批注保护。源码集成测试需要 Word、Pandoc 和已构建的 `build/main.pdf`。
