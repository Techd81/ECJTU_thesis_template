# 校园自习室预约系统设计与实现

这是一篇使用华东交通大学 LaTeX 模板编写的短篇示例论文，封面填写示例作者与教师信息，日期为 2026 年 10 月 9 日。论文包含中英文摘要、三级目录、六章正文、原生结构图、表格、公式、致谢、自动参考文献和两个附录。

预约业务原型采用 Python 标准库和内存 SQLite 数据库。第五章的 12 项功能测试在本地实际执行，结果保存在 `assets/test-results.json`；测试数量、运行版本和统计值由 `assets/test-summary.tex` 自动写入论文。

在仓库根目录执行：

```powershell
python examples/simple-thesis/build.py
```

在独立的示例目录中执行：

```powershell
python build.py
```

构建入口依次运行原型、XeLaTeX / BibTeX 构建与 `--document` 格式检查，生成 `build/main.pdf` 和 `build/format-report.json`。需要 TeX Live、模板所用 Windows 字体，以及 `PyMuPDF`、`Pillow` 和 `pikepdf`。

`main.tex` 为主文件，`metadata.tex` 填写封面信息，`chapters/` 保存各章节，`references.bib` 保存三项实际引用的教材。示例使用 `final` 模式，修改后的正文由 LaTeX 流动排版。

类文件、素材与构建工具沿用本次核验的模板版本。封面和字体素材的来源及使用范围见原仓库说明。本文用于短篇内容与排版测试，正式论文的篇幅和文献要求请按学院规定补充。

此次样稿共 19 页。新正文采用正常 LaTeX 排版，格式检查全部通过；全部页面经 Poppler 144 dpi 渲染检查，图表、公式、跨页段落、引用和代码附录显示正常。

## Word 转换测试

在完整模板仓库根目录执行：

```powershell
python scripts/export_word.py --source examples/simple-thesis/main.tex --pdf examples/simple-thesis/build/main.pdf --output output/docx/ECJTU_示例论文_自习室预约系统.docx
python scripts/verify_word.py output/docx/ECJTU_示例论文_自习室预约系统.docx --output tmp/simple-thesis-word-final-editing
```

该示例通过 795 项内容与结构检查及 310 项本机 Microsoft Word 编辑检查。正文、标题、目录、3 张表、14 个数学对象、3 条参考文献及完整代码可在 Word 中修改；目录可更新，批注可保存。封面底图和结构图为图片，图内元素需要替换图片或在源图中修改。

严格版式对照未通过：LaTeX 为 19 页，Word 为 18 页。两张前置空白页未保留，目录由一页变为两页，表格列宽、边框及部分换行存在差异。部分段落含用于接近源 PDF 行尾的手动换行，增删文字后需复核字距与分页。此 Word 样稿可用于编辑与审阅测试，转换结果尚未达到与 LaTeX 完全一致的版式。详细记录见 `verification/word-summary.json`。
