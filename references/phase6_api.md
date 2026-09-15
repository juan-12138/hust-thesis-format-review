# 阶段6 Word 格式审查报告（历史接口）

阶段7已在此报告与批注工作流之上加入57条高级规则；当前执行入口和能力边界以 `phase7_api.md` 为准。本文件保留用于说明报告器的设计来源。

## 输出

阶段6一次审查默认生成两个独立 Word 文件：

- `<原文件名>_格式审查批注版.docx`
- `<原文件名>_格式审查报告.docx`

源论文保持不变。报告汇总阶段4产生的全部 Finding，并通过 `finding_id` 合并阶段5批注状态。已经写入批注的问题仍保留在报告中；不能可靠批注的问题和人工复核项分别进入专项章节。

## 一键运行

先调用 Codex 工作区依赖加载器，使用其返回的 Python executable：

```powershell
$codexPython = "<工作区依赖加载器返回的 Python executable>"
& $codexPython scripts/review_thesis.py "待审查论文.docx" `
  --manifest "待审查论文_格式审查结果.json"
```

可选参数：

- `--output`：指定批注版 DOCX；
- `--report-output`：指定审查报告 DOCX；
- `--manifest`：保存完整审查、批注和报告映射 JSON；
- `--author`、`--initials`：批注作者信息；
- `--template-name`：报告概况中显示的模板名称；
- `--rules`：经过确认的兼容规则 JSON 镜像。

Python 接口为 `review_thesis.run_review()`。返回值顶层包含 `review`、`comments` 和 `report`。

## 只生成报告

已有阶段4或阶段5 JSON 时，可直接运行：

```powershell
& $codexPython scripts/report_writer.py "原论文.docx" `
  --review "阶段5完整结果.json" `
  --output "原论文_格式审查报告.docx" `
  --manifest "报告生成结果.json"
```

若审查 JSON 只包含阶段4结果，可另用 `--comments` 指定阶段5批注映射。Python 接口为：

- `report_writer.default_report_path(source)`；
- `report_writer.write_report(source, review, comments, output=...)`。

## 固定报告结构

报告固定包含：

1. 审查概况
2. 严重问题
3. 页面与版式问题
4. 字体与字符格式问题
5. 段落问题
6. 标题与编号问题
7. 图问题
8. 表问题
9. 公式问题
10. 目录问题
11. 页眉页脚与页码问题
12. 脚注、参考文献与附录问题
13. 无法通过 Word 批注定位的问题
14. 需要人工复核的问题
15. 审查统计

第2节是 ERROR 摘要；第3至12节是按对象类别组织的完整明细；第13、14节是便于处理的专项子集。重复展示不影响统计，统计始终按唯一 `finding_id` 计算。

## 问题明细与位置

问题表包含：编号、规则、位置、问题与对象、当前格式、标准格式、严重度、是否已批注。问题与对象单元格同时保留依据和建议。

位置由可用的节号、正文/页眉页脚段落、可读标签和前40字定位摘要组成。报告不会把 `word/document.xml ¶123` 作为唯一位置。因为 OOXML 不能稳定提供最终分页，位置中明确写入：

`页码：Word XML 无法可靠获得最终页码`

## 覆盖声明

阶段6只汇总当前已经实现的19条基础规则。页面、页眉页脚、字体、段落和标题章节在没有 Finding 时可以写“已执行且未发现问题”。图、表、公式、目录、脚注、参考文献和附录专项检查仍属于阶段7；这些章节必须写“当前规则覆盖中尚未包含此类专项检查”，不能写成“未发现问题”。

## 报告版式

报告使用 A4 横向版式，以容纳八列问题表。标题和章节标题均为黑色，中文使用黑体或宋体，拉丁文字使用 Times New Roman。表头采用深蓝底白字，正文交替白色和浅蓝色，全部边框为浅灰色；表头设置跨页重复，行高不固定。页脚使用 Word `PAGE` 域。

## 验证边界

生成后至少检查 DOCX ZIP、所有 XML、15个章节、Finding 完整性、严重度统计、批注映射、重复表头和页码域。若工作区依赖未提供打包 LibreOffice，则不能声称完成视觉渲染验证，应在交付说明中记录该限制。
