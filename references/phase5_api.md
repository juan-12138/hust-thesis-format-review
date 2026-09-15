# 阶段5 Word 原生批注

## 能力与边界

阶段5在阶段4的 19 条基础规则检查之上生成 Word 原生批注副本。源 DOCX 保持不变，默认输出名为：

`<原文件名>_格式审查批注版.docx`

自动批注仅覆盖同时满足以下条件的 Finding：

- 严重程度为 `ERROR` 或 `WARNING`；
- `anchor.part_name` 为 `word/document.xml`；
- 段落索引有效且锚点摘要仍与目标段落匹配。

`MANUAL_REVIEW`、页/节级问题、页眉页脚部件问题、失效锚点和其他故事部件不会被强行附着到正文。它们保留在 manifest 中并带有 `skip_reason`，供阶段6报告完整呈现。

## 一键运行

先调用 Codex 工作区依赖加载器，并使用其返回的 Python executable：

```powershell
$codexPython = "<工作区依赖加载器返回的 Python executable>"
& $codexPython scripts/review_thesis.py "待审查论文.docx" --manifest "待审查论文_格式审查结果.json"
```

阶段6版本中的一键入口会同时生成批注版和 Word 审查报告。只需要批注版时，直接使用下节的 `comment_writer.py`。可用 `--output` 指定批注版路径，用 `--author` 和 `--initials` 改变批注作者信息，用 `--rules` 指定另一份已经确认且结构兼容的运行时规则镜像。

## 分步运行

如已有阶段4 JSON，可直接调用：

```powershell
& $codexPython scripts/comment_writer.py "待审查论文.docx" `
  --findings "待审查论文_阶段4格式审查.json" `
  --output "待审查论文_格式审查批注版.docx" `
  --manifest "待审查论文_批注映射.json"
```

Python 接口：

- `comment_writer.default_output_path(source)`：生成规定的默认输出名；
- `comment_writer.write_comments(source, findings, output=None, ...)`：写批注副本并返回映射清单；
- `review_thesis.run_review(docx_path, ...)`：组合阶段4检查和阶段5批注写入。

## 批注格式与合并

每个 Finding 在批注中使用以下字段：

```text
【格式审查：规则编号】
问题：……
当前：……
要求：……
依据：……
建议：……
严重程度：错误/警告
```

同一主文档段落上的多个 Finding 合并到一条批注，避免相同或交叉范围的多条批注破坏 Word 锚点。每个 Finding 仍保留自己的完整字段和规则编号。

## OOXML 保持策略

批注过程最多修改或新增以下四个部件：

- `word/document.xml`；
- `word/comments.xml`；
- `word/_rels/document.xml.rels`；
- `[Content_Types].xml`。

其他 ZIP 部件使用源文件原始字节和条目元数据写回。已有批注会被保留，新批注 ID 只依据现有批注及批注锚点避碰，不会被书签等无关 `w:id` 干扰。

批注写入不会更改任何 `w:t` 正文文本，也不会主动重建样式、图片、公式、域、书签、超链接、内容控件、修订、脚尾注或分节。批注范围无法稳定映射到直接 run 时，会退化为段落级安全锚点，而不是拆分 run。

## Manifest

返回或导出的 manifest 对每个 Finding 记录：

- `commented`：是否已写入批注；
- `comment_id`：对应 Word 批注 ID；
- `anchor_mode`：`run_range` 或 `paragraph`；
- `skip_reason`：未批注原因，如 `manual_review`、`missing_anchor`、`unsupported_story_part`、`paragraph_not_found` 或 `snippet_mismatch`。

`comment_writer.py` 本身不生成独立 Word 审查报告，也不声称完成依赖最终分页或视觉渲染的判定。完整两文件输出见 `phase6_api.md`。
