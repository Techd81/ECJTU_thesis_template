# Word 版式校准进度

2026-10-06，用户确认正文和表格须保持可编辑，要求逐页修正到 PDF 版式。工作尚未完成。

## 当前证据

原交付 `build/main-review.docx` 的哈希与原 map 一致，重新用 Microsoft Word 渲染到 `tmp/fidelity-baseline/word.pdf`。参考为 `build/main.pdf`。二者都是 23 页，但存在明显换行、标题位置、段落位置和目录排版差异。原来的结构检查不能证明视觉一致。

最新候选为 `tmp/fidelity-cjk.docx`，map 中的 work 为 `tmp/source-word-_z_aoktn`。最终渲染为该目录的 `word.pdf`。逐页对照及严格像素报告在 `tmp/fidelity-cjk-comparison`。尚未覆盖原交付文件。

已完成：

- 在 Word 实例内显式设置 CompatibilityMode 15，修复保存后退回 12 的问题。
- 启用中文标点压缩和中西文、中文数字间距。
- 从参考 PDF 和 Word 实际渲染提取章标题与紧随段落基线，校准标题字形位置及段后间距，保持流式正文与原生表格。
- 校准后更新 Word 目录页码，保存并重新渲染最终文件。
- 新增独立 `scripts/compare_word_pdf.py`，输出所有页面的原图、并排图、红色差异图和字符基线诊断。严格相等判定不做配准、裁切或缩放；多出或缺少的页面也会检出。

最新候选第 7 页匹配字符的水平偏差中位数为 0.04 pt，垂直为 0.08 pt。多数章首页的垂直中位差已降到 0.03～0.21 pt。所有 23 页仍有像素差异，因此严格一致检查仍失败。

26 项 Word 回归通过，日志 `tmp/fidelity-regression.log`；随后新增章基线回归，15 项源码 Word 测试再次全部通过，日志 `tmp/fidelity-source-tests.log`。实际编辑、替换、换行、表格扩展、批注、保存重开检查 467/467 通过，报告 `tmp/fidelity-editing-qa/word-verification.json`，对应最新候选哈希。

## 待解决与复核

1. 目录：标题约低 7.5 pt，条目低约 5 pt；二三级标题文字起点缺少 LaTeX 编号盒的固定宽度，点引导线的间距也不同。Word 原生 TOC 需要保留。
2. 续页：第 9、11、12、15、16 页常见约 3.7 pt 的顶端基线差；第 17 页多带入上一页的一行，垂直中位差约 23.19 pt。
3. 长段落：标点压缩与英文空格仍有换行差异；部分条目中的项目符号字形大小不同。
4. 封面、声明、中英文摘要、页眉横线与页码位置仍需进一步对齐。声明横向偏差明显；摘要垂直偏差约 6～9 pt。
5. 逐页核验图表、公式、参考文献及全部页面；当前仅人工查看了对照中的第 5、7、11、17 页，不能声称最终全页视觉验收通过。
6. 更新正式交付、使用说明与核验记录前，运行全部回归、实际编辑验证及最终逐页视觉检查。保持 goal active。

## 环境

Windows 本机 Word 16.0 (16.0.20326)，使用项目现有 Python/Pandoc/Word COM 流程。当前工具列表没有技能要求的容器 workspace dependency loader；已有渲染使用实际 Microsoft Word。已执行 documents 的 artifact-operation marker。未用 LibreOffice 代替用户的 Word 渲染。

探索子代理启动失败，服务不支持其默认 gpt-5.6-luna；本轮改为主代理直接探索。fast-context 已定位源码导出和结构检查路径。
