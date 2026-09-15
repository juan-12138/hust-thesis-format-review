# 阶段3解析接口

## 能力边界

阶段3只读取 DOCX，不改写源文件。它实现以下底层能力：

- OPC/ZIP 部件读取和安全 XML 解析；
- 正文、表格单元格、内容控件、修订、页眉、页脚、脚注、尾注及原批注中的段落和文字片段遍历；
- Section 页面设置、页码设置以及页眉页脚关系继承；
- 样式、编号、设置、关系、域、书签、超链接、DrawingML、VML、OMML、OLE、文本框及空白字符清点；
- 文档默认值、表格整体样式、段落样式继承、字符样式继承和直接格式合并；
- 中文、英文/数字和复杂文字字体分别解析；
- 主题字体、主题颜色、tint/shade 到 RGB 的解析；
- 段落结构角色分类，并为每个结果提供置信度和分类依据。

阶段3不执行规则判定，不生成错误/警告，不写入 Word 批注，也不创建审查报告。这些功能属于阶段4至阶段7。

## 命令行

解析器依赖 `lxml`。在 Codex 中先调用工作区依赖加载器，并使用其返回的 Python 可执行文件；不要默认使用系统 PATH 中的 `python`。首次运行或运行环境变化后，可先执行导入检查：

```powershell
$codexPython = "<工作区依赖加载器返回的 Python executable>"
& $codexPython -c "import lxml; print(lxml.__version__)"
& $codexPython scripts/inspect_docx.py "待审查论文.docx" --output "论文结构.json"
& $codexPython scripts/inspect_docx.py "待审查论文.docx" --include-runs --output "论文详细结构.json"
```

`--max-paragraphs N`只限制 JSON 中输出的段落明细，不影响整篇文档的统计和结构分类。

## Python 接口

### docx_model.py

`DocxPackage(path)`在内存中建立只读部件索引。主要接口：

- `xml(part_name)`：解析并缓存 XML 部件；禁止外部实体和网络访问。
- `relationships(source_part)`：返回关系 ID、类型、目标、解析后的内部部件路径及是否外部关系。
- `iter_paragraphs(part_name)`：遍历指定部件内的所有段落，包括嵌套在表格、内容控件和修订中的段落。
- `iter_all_paragraphs()`：遍历正文、页眉、页脚、脚注、尾注和原批注。
- `iter_runs(paragraph)`：遍历段落内文字片段并保留父段落定位。
- `sections()`：返回纸张、方向、页边距、装订线、节起始方式、首页不同、页码格式/起始值，以及显式和继承后的页眉页脚引用。
- `settings()`：返回默认制表位、奇偶页页眉、修订跟踪、打开时更新域、镜像页边距和兼容模式。
- `numbering_summary()`：返回抽象编号定义、编号实例和各级编号文本。
- `styles_summary()`：返回段落、字符、表格和编号样式清单。
- `relationship_inventory()`：返回主文档及辅助文字部件的关系清单。
- `inventory()`：返回对象、域、空白字符、分页符、修订、批注和媒体结构统计。

### effective_style.py

`EffectiveStyleResolver(package)`公开：

- `get_effective_run_format(run, paragraph)`：返回有效字体、字号、粗斜体、上下标、下划线、删除线、字符间距、缩放、位置、颜色、高亮、隐藏等属性，并在 `provenance` 中记录每个属性的来源层。
- `get_effective_paragraph_format(paragraph)`：返回对齐、缩进、段前段后、行距、制表位、与下段同页、段中不分页、孤行控制、段前分页、上下文间距、大纲级别和编号等属性及来源。
- `style_chain(style_id, style_type)`：返回由基础样式到当前样式的继承链；循环继承会安全终止。
- `ThemeResolver`和`apply_tint_shade()`：解析主题字体及 RGB、themeTint、themeShade。

当前只合并表格样式本体和 `wholeTable` 条件。首行、末行、首列、末列和条带条件属于后续高级表格阶段。

### structure_classifier.py

`StructureClassifier(package, resolver).classify_document()`返回段落及其：

- `role`：封面/摘要/关键词/目录/各级标题/正文/图题/表题/公式/参考文献/致谢/附录等角色；
- `confidence`：0 到 1 的候选置信度；
- `evidence`：命中的文字、样式、大纲级别或上下文依据。

低置信度分类不能直接作为格式错误，必须由后续规则引擎结合上下文处理。

## JSON 输出

命令行用法：

```powershell
$codexPython = "<工作区依赖加载器返回的 Python executable>"
& $codexPython scripts/inspect_docx.py "待审查论文.docx" --output "论文结构检查.json"
& $codexPython scripts/inspect_docx.py "待审查论文.docx" --include-runs --output "论文逐文字格式.json"
```

`--max-paragraphs N`只限制写入 JSON 的段落详情数量，不影响整篇统计。`--include-runs`会显著增大输出文件，仅在排查具体字体、字号或颜色继承时使用。

顶层包括：

- `source`：文件名、绝对路径、大小和 SHA-256；
- `capabilities`：明确列出已实现及未实现能力；
- `inventory`、`sections`、`settings`、`numbering`、`styles`、`relationships`；
- `role_counts`：结构角色计数；
- `paragraphs`：定位文字、样式、节索引、结构角色和有效段落格式；
- `runs`：仅在使用 `--include-runs` 时输出；
- `limitations`：不能从 OOXML 可靠获得的项目。

## 已知限制

- OOXML 不能可靠提供最终页码、自动分页、视觉重叠、孤行和大面积留白。
- 结构分类是后续规则引擎的输入，不是最终审查结论。
- 字体替换和系统缺字造成的最终显示效果需要渲染环境确认。
- 解析器不更新域，也不把缓存的目录或交叉引用结果视为最新结果。
