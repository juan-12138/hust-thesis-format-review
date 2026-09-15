# 阶段7高级对象审查

## 覆盖范围

阶段7将阶段4的19条基础规则与57条高级规则合并，`review_rules.py` 对规则模型中的76条规则逐条返回 `RuleResult`。新增检查包括：

- 封面、声明页、摘要与关键词；
- 自动目录及目录条目格式；
- 图、图题、正文引用和视觉复核队列；
- 三线表、表题、表中文字、跨页表和表注；
- OMML公式对象、公式对齐结构和编号；
- 脚注字体、字号、段落、编号设置和网页来源候选；
- 参考文献连续编号、作者数量、末尾标点、类型结构、数量构成和姓名写法；
- 形状、文本框、域、书签、交叉引用、修订、原批注和内容控件；
- 缩写、数字与单位空格、章节构成、本章小结和附录候选。

## 一键运行

使用工作区依赖加载器返回的 Python executable：

```powershell
$codexPython = "<工作区依赖加载器返回的 Python executable>"
& $codexPython scripts/review_thesis.py "待审查论文.docx" `
  --output "待审查论文_格式审查批注版.docx" `
  --report-output "待审查论文_格式审查报告.docx" `
  --manifest "待审查论文_格式审查结果.json"
```

只生成JSON：

```powershell
& $codexPython scripts/review_rules.py "待审查论文.docx" --output "格式审查.json"
```

Python接口：

- `advanced_rule_engine.AdvancedRuleEngine(package, rules_path=...).review()`；
- `review_rules.build_review(docx_path, rules_path=...)`；
- `review_thesis.run_review(...)`。

## 结果语义

- `ERROR`：OOXML证据能够确认的明确违规；
- `WARNING`：候选识别可靠，但分类、例外或上下文仍可能影响结论；
- `MANUAL_REVIEW`：必须依赖语义、事实核验、最终分页或视觉判断；
- `PASS`：在已检查的适用对象中未发现问题。若 `evaluated_count` 为0，只表示未发现适用对象，不能证明该类内容合规。

图表题注相邻方向、编号、目录域、脚注有效样式、表格边框、参考文献序号和书签目标等可结构化检查。图表科学性、摘要双语语义、参考文献真实性、页面裁切/碰撞、最终页码及可读性始终保留人工复核。

## 批注与报告

只有 `ERROR` / `WARNING` 且锚点位于 `word/document.xml` 的发现写入Word原生批注。脚注部件、页眉页脚部件、节级问题和所有 `MANUAL_REVIEW` 项只进入报告，避免将问题错误地挂到无关正文。

报告继续使用15节固定结构。阶段7的图、表、公式、目录、脚注和参考文献章节在其规则确实进入 `implemented_rule_ids` 后，才可显示“已执行当前阶段相应规则”；旧的阶段4结果仍显示未覆盖。

## 验证

执行完整测试：

```powershell
& $codexPython -m unittest discover -s . -p 'test_*.py' -v
```

测试覆盖图题在图上方、非三线表、公式编号缺少圆括号、参考文献跳号及末尾标点、脚注字号/缩进错误、76条规则结果完整性，以及原有解析、有效格式、批注、报告和模板集成回归。

若Microsoft Word或LibreOffice不可用，验证仅包括ZIP/XML、结构清单、关系、Finding覆盖和输出文件完整性；不得声称已完成最终视觉渲染检查。
