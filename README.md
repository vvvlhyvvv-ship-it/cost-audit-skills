# 造价审核技能族（Cost Audit Skills）

面向**造价工程师**的 WorkBuddy 技能集合，覆盖合同价格台账构建、审核公文撰写、公文版式、结算书汇总、投资风险评估等完整工作流。

## 技能清单

| 技能 | 用途 |
|---|---|
| `contract-ledger-builder` | 从散落合同源文件（docx/doc/pdf/jpg/xls）构建/整改合同价格台账：源文件转换与 OCR、供应商/电话/价目表抽取（带甲乙方方位校验）、版本定版去重、Decimal 勾稽自验 |
| `cost-audit-doc-writer` | 造价审核类公文的内容撰写与成文（进度复核报告、情况说明、结算审核报告）：主体具名、现场复核措辞、事实与结论分章 |
| `cost-audit-farmland-workflow` | 高标准农田项目审核端到端流程：三步法、合同变更汇总表枢纽、概算控制层、核减建议 |
| `investment-risk-assessment` | 政府投资项目投资风险评估报告：五席核验、规范依据链、Decimal 复算口径 |
| `gongwen-doc-format` | 公文/长文档版式标准与套版（Word docx）：页边距/字体/字号/行距/目录/页码的唯一权威来源 |
| `settlement-merge-summary` | 工程结算书明细表合并汇总：精度保持、公式链审计、差异报告 |
| `fee-control-price-doc` | 工程建设其他费用（二类费用）咨询价格文档生成 |

## 使用方式

在 WorkBuddy 中通过技能菜单选择，或直接以技能名调用（如"用 contract-ledger-builder 技能建台账"）。

## 设计原则

- **单一职责**：版式规格只存于 `gongwen-doc-format`，其他技能只指路不复制；
- **证据链闭环**：所有数字带口径与来源，Decimal 精算，实测/估算/待核三级标注；
- **先备份后动手**：台账类操作每批大改前双版本备份留证；
- **自验齐备**：金额勾稽、双版本一致性、保护行核对通过才算完成。

## 目录结构

```
contract-ledger-builder/     # 合同价格台账构建与整改
cost-audit-doc-writer/       # 造价审核公文编写
cost-audit-farmland-workflow/# 农田项目审核流程
investment-risk-assessment/  # 投资风险评估
gongwen-doc-format/          # 公文版式标准
settlement-merge-summary/    # 结算书汇总
fee-control-price-doc/       # 二类费用文档
```

> 注：技能内容已脱敏（单位/客户/项目名均为泛称），方法论与工作流完整保留。
