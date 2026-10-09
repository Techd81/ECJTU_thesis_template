# 华东交通大学软件工程本科毕业论文 LaTeX 模板

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![LaTeX](https://img.shields.io/badge/TeX%20Live-2026-blue.svg)](https://tug.org/texlive/)

> 托管于 [Techd81/ECJTU_thesis_template](https://github.com/Techd81/ECJTU_thesis_template)，MIT 许可开源，欢迎使用与贡献。

本模板由学院 Word 模板转换，当前以根目录《指导老师-学生姓名-毕设论文 2026届软件工程(1).docx》及其同名 PDF 为版式参照，并保留原稿的 19 条批注要求。模板包含封面、诚信声明、中英文摘要、三级目录、七章正文、致谢、参考文献、两个附录，以及原稿的写作提示和图表示例。

## 编译

使用 **XeLaTeX + BibTeX + latexmk**，本地已核验 TeX Live 2026。需要安装以下字体：宋体 `SimSun`、黑体 `SimHei`、楷体 `KaiTi`、`Times New Roman`。字体缺失时需先安装相应授权字体，模板会报错提示。

在本目录执行：

```powershell
python -m pip install -r requirements.txt
python scripts/build.py --check
```

生成的论文为 `build/main.pdf`，格式检查报告为 `build/format-report.json`。使用 TeX 编辑器时，将 `main.tex` 设为主文件，构建命令设为 `latexmk -xelatex main.tex`；首次直接运行 latexmk 前建立 `build/chapters` 目录。

构建会保存 XeLaTeX 的原始输出 `build/main.native.pdf`，并通过 `.latexmkrc` 调用 Python，应用 `assets/guide-layout.json` 中的原稿字距、基线、文字框和边线测量参数，生成最终 PDF。当前随附范例的 27 页在 MuPDF 144 dpi、288 dpi 下均与参考 PDF 逐像素一致。校准记录 `build/main.calibration.json` 关联源文件、参数和 PDF 的哈希，并核对文字、字体、书签和链接的保留情况。

测量参数仅对内容和布局均匹配的随附范例生效。修改正文、元数据、类文件或素材后，构建自动采用正常 LaTeX 流动排版，状态记为 `normal-layout`；新增文字及表格内容保持可编辑。编译读取本地源码、素材和测量参数。依赖安装需要网络，TeX Live、字体和 Python 依赖齐备后可以离线构建。

## 导出可编辑 Word

源码导出需要 Windows、Microsoft Word 桌面版、Pandoc 3.x（加入 PATH），以及本项目 Python 依赖。写完论文后执行：

```powershell
python scripts/build.py --word
```

同时生成 `build/main.pdf` 和 `build/main-review.docx`。Word 正文支持打字、删除、替换、换行和批注，表格支持单元格编辑。老师可直接修改内容，或通过“审阅 → 新建批注”及 `Ctrl+Alt+M` 添加意见。作者将修改同步到 LaTeX，再生成下一版。

导出直接读取 LaTeX 源码，生成完整段落、三级标题、Word 自动目录、原生表格和公式，并应用学校的页边距、字体、字号、缩进、行距及页眉页脚。封面底图和工资流程图以图片保留，题注可编辑。目录可右键“更新域”，正文从第 1 页连续编号。Word 与 XeLaTeX 的换行、标点压缩和分页可能略有差别；引用编号随 LaTeX 构建生成，改动章节或引用后需重新导出。Word 导出最近的完整核验为 2026 年 10 月 6 日，后续 PDF 校准尚未重新验证 Word 导出。

只有独立 PDF 时，可使用近似转换入口，本地只需 Python 依赖；该入口依靠 PDF 推断段落，不应用本项目的源码排版规则：

```powershell
python scripts/export_word.py --pdf "论文.pdf" --output "论文-送审.docx"
```

首次使用或更新导出脚本后，先执行 `python -m pip install -r requirements.txt`。源码导出自动核对源段落、目录层级、学校样式、图片尺寸及文字可见性，并用 Word 更新目录及渲染。生成的 `.map.json`、`.check.json` 用于核验；老师只需要 `.docx`。已修改、带批注或修订的文件会被保留，重新导出时请使用新文件名。详细说明见 [Word 审阅稿](docs/Word审阅稿.md)。

## 填写论文

1. 修改 `metadata.tex` 中的题目、姓名、学号、专业班级、导师和日期。封面、摘要标题、偶数页页眉、PDF 元数据同步更新。
2. 在 `chapters/abstracts.tex` 填写摘要与关键词，在 `chapters/chapter01.tex` 至 `chapter07.tex` 撰写正文。章、节、三级标题分别使用 `\chapter`、`\section`、`\subsection`，编号及目录由编译生成。
3. 替换 `chapters/figure-table-example.tex` 中的工资流程图和表格。图用 `\caption` 置于图后，表用 `\caption` 置于表前；用 `\label`、`\ref` 引用。工资流程图使用 `assets/salary-flow.png`，表格为可编辑的原生 LaTeX 表格。
4. 在 `references.bib` 维护真实文献，在正文使用 `\cite{文献键}`。编号自动以上标显示，著录采用 GB/T 7714—2015 顺序编码制。完成文献替换后删去 `main.tex` 的 `\nocite{*}`，让文后只列实际引用的条目。
5. 修改致谢、软件使用说明书和主要源代码附录。用 `\templatehint{...}` 标记的段落是写作提示，`\templateheading{...}` 是与下一段保持同页的提示标题，正式正文直接写在命令外。
6. 完成内容替换后，可将主文件改为 `\documentclass[final]{ecjtu-thesis}` 隐藏写作提示。该选项仅控制提示显示，原稿的摘要、绪论范文、图表和文献范例仍需要自行替换。

章节调整后，运行 `python scripts/build.py`，再运行 `python scripts/check_format.py --document`。默认检查还会核对原 Word 的 59 个章节标题及 4 个后置条目；`--document` 允许你修改章节名称和结构。

## 排版检查

`python scripts/build.py --check` 检查当前源码对应的 PDF、编译错误和警告、缺字、松紧盒、实际 A4 尺寸和页边距、字体嵌入、黑色正文、字号、行距、目录层次与页码、目录跳转、页眉页脚、连续页码和图表示例。检测实际渲染墨迹是否越出版心，避免中文标点字面框的空白造成误判。

修改内容后还应运行 `python scripts/render_pdf.py`，逐页打开 `tmp/qa/page-XX.png`，检查长标题、复杂公式、图表和分页。自动检查不能判断论文内容质量，也不等同于学校未提供的外部检测系统。

与指定 PDF 的严格对照使用 `scripts/compare_word_pdf.py`，在相同分辨率下比较完整页面，不对齐、裁切或缩放结果。常规格式检查通过与逐像素一致分别记录；当前核验版本及剩余差异见 [核验记录](docs/核验记录.md)。

维护模板时可运行 `python scripts/verify_template.py`。它在临时副本中验证首次编译的全部页面可复现、改写正文后的内容保留、`final` 模式、陈旧 PDF 和错误页边距的检出，以及公式、跨页表格、源代码和脚注示例。

`python scripts/build.py --check --word` 同时运行 PDF 格式检查和 Word 导出结构检查。在已安装 Microsoft Word 的 Windows 上，可安装 `pywin32` 后执行 `python scripts/verify_word.py`，在临时副本中验证实际输入、替换、换行、表格扩展及批注保存，重新打开后检查新增文字的可见渲染，并生成逐页图片供检查。

具体格式依据及原稿内的差异处理见 [格式对照](docs/格式对照.md)；章节写作要求见 [原模板写作说明](docs/原模板写作说明.md) 和 [原模板批注](docs/原模板批注.md)；公式、续表和代码用法见 [补充排版示例](docs/补充排版示例.md)。

## 提交前的内容要求

- 题目按学院批注建议控制在 20 字以内，中文摘要约 300～500 字，中英文内容对应，关键词 3～5 个。
- 正文按学院批注要求 15000 字以上。
- 原模板要求参考文献 30 篇以上、近 5 年文献不少于 70%、英文文献至少 3 篇；应以本人实际阅读和引用的文献为准。
- 主要源代码附录按原批注控制在 10 页以内，选取核心实现并添加注释。
- 本项目保留原稿范例用于演示排版，字数、文献数量和时效不构成论文提交结果。
- 彩色封面沿用原稿内嵌底图。原 Word 说明纸质封面与资料袋需按学院要求领取、装订；任务书、开题报告、评阅书、答辩记录等使用学院配发的表格一并装订。

## 文件布局

| 文件 | 用途 |
| --- | --- |
| `main.tex` | 论文主文件与装订顺序 |
| `ecjtu-thesis.cls` | 统一格式、标题、页眉页脚及前置页面 |
| `metadata.tex` | 个人信息 |
| `chapters/` | 摘要、正文及附录 |
| `references.bib` | 文献数据库 |
| `assets/cover-background.png` | 原 Word 内嵌封面底图 |
| `assets/salary-flow.png` | 原稿工资流程图 |
| `assets/` 中的字体 | 与参考 PDF 字形一致所需的嵌入字体资产 |
| `assets/guide-layout.json` | 随附范例的源文件指纹与数字排版参数 |
| `scripts/` | 编译、检查与渲染工具 |
| `docs/` | 使用依据、写作说明和核验记录 |

封面图及校名校徽来自用户提供的学校模板，按学校规定使用；仓库代码许可不改变这些素材及字体的权利归属。当前参考文件及派生资产的来源、哈希见 [来源记录](docs/reference-files.json)。
