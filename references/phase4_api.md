# 阶段4基础规则引擎

## 能力与边界

阶段4在阶段3只读解析结果之上执行 19 条规则：

- 页面与 Section：`PAGE-001`～`PAGE-004`；
- 页眉、页脚和页码：`HF-001`～`HF-004`；
- 正文字体、字号、颜色和分语种字体：`FONT-001`～`FONT-003`；
- 正文段落、空段和人工分页：`PARA-001`～`PARA-003`；
- 一至三级标题格式、编号层级和标题表述候选：`HEAD-001`～`HEAD-005`。

阶段4不修改源 DOCX，只输出 JSON。Word 原生批注属于阶段5，Word 审查报告属于阶段6，图、表、公式、目录、参考文献和脚注专项检查属于阶段7。

## 运行环境与命令

先调用 Codex 工作区依赖加载器，使用其返回的 Python executable。运行时需要 `lxml`，但不需要 PyYAML；规则引擎读取完整 YAML 的 JSON 镜像，并核对源 YAML 的 SHA-256。

```powershell
$codexPython = "<工作区依赖加载器返回的 Python executable>"
& $codexPython scripts/review_rules.py "待审查论文.docx" --output "论文_阶段4格式审查.json"
```

可用 `--rules` 指定另一份同结构的 JSON 规则配置。不要在未经用户确认的情况下替换标准规则。

## Python 接口

`rule_engine.py` 公开：

- `load_rule_config(path)`：加载规则、检查重复 ID、确认 19 条阶段4规则齐全，并在源 YAML 存在时验证镜像哈希；
- `RuleEngine(package, rules_path=...).review()`：返回 `ReviewResult`；
- `ReviewResult.rule_results`：每条规则的 `status`、`finding_count` 和 `evaluated_count`；
- `ReviewResult.findings`：可供后续批注与报告阶段直接使用的定位化问题记录。

`review_rules.build_review(docx_path, rules_path)` 组合解析器与规则引擎并生成完整 JSON 对象。

## Finding 结构

每条 Finding 至少包含：

- `finding_id`、`rule_id`、`category`、`object_type`；
- `severity`：`ERROR`、`WARNING` 或 `MANUAL_REVIEW`；
- `problem`、`actual`、`expected`、`source`、`suggestion`；
- `location`：部件、节、段落和可读定位文字；
- `anchor`：可可靠定位时给出段落和 run 范围，供阶段5写批注；
- `evidence`：分类依据、属性来源、容差或需要渲染的说明。

同一段中连续且相同的 run 格式问题合并为一条 Finding。页或节级问题不会为了制造批注锚点而插入正文。

## 判定原则

- 正文检查从首个带正文章编号的一级标题开始，在参考文献、致谢或附录处结束；表格、题注、公式和注释不套用正文规则。
- 中文字体与英文/数字字体分别比较；解析 `eastAsia`、`ascii`、`hAnsi`、样式继承和主题回退。
- 正文颜色允许自动黑色和 `#000000`；超链接、域、修订和隐藏文字不按普通正文颜色误报。
- 页脚居中同时识别段落居中和居中制表位，不能只读 `w:jc`。
- PAGE-004 和 PARA-003 只定位候选。实际空白页、分页碰撞和视觉位置必须渲染后人工复核。
- 有效格式缺失且无法确认 Word 回退结果时，输出 `MANUAL_REVIEW`，不得升级为确定错误。
- `PASS` 需要结合 `evaluated_count` 阅读；计数为 0 表示文档中没有可适用对象，而不是证明该类对象存在且合规。

## JSON 顶层字段

- `source`：绝对路径、文件名、大小和 SHA-256；
- `rule_set`：规则集 ID、版本、运行时配置路径和镜像来源；
- `capabilities`、`coverage`、`limitations`；
- `statistics`：Finding 严重度统计和规则状态统计；
- `rule_results`：19 条规则逐条状态与检查对象数；
- `findings`：完整问题明细。
