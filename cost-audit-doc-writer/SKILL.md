---
name: cost-audit-doc-writer
description: "造价审核类公文的**内容**撰写与成文（工程进度复核报告、初步审核情况说明、结算审核报告等）：主体具名、现场复核措辞、事实与结论分章、完成比例归结论、后附现场影像、真实单位落款、篇幅 1–2 页。**版式（页边距/字体/字号/行距/缩进/目录/页码）一律以技能 `gongwen-doc-format` 为准，本技能不重复定义**。当用户要求编写/生成造价审核报告、情况说明、进度复核、结算审核等审核类公文时使用。"
version: "1.3.0"
agent_created: true
---

# 造价审核公文编写（工程进度复核 / 情况说明 / 结算审核）

> 本 Skill 固化用户在「某肉牛精深加工仓储冷链物流项目进度复核报告」「某县高标准农田初步审核情况说明」「某农业产业园蛋鸡项目情况说明」中确立的公文格式与内容范式。
> 适配场景：造价工程师出具各类审核类公文，要求格式严谨、内容合规、落款规范。

---

## ⛔ 强制执行规则（MUST）

0. **版式不在这里定**。凡涉及页边距 / 字体 / 字号 / 行距 / 缩进 / 目录 / 页码 / 标点，先读技能 **`gongwen-doc-format`**，并以**现读母本实测**为准。**本技能不复述任何版式数值。**
1. **格式必须逐 run 直接格式化**（直设字体 + eastAsia），严禁依赖 Word 样式名。原因：用户的 WPS/Word 环境中样式克隆不稳定（editor_sdk 无头读取对该类模板不可靠），python-docx 直设最稳。
2. **主体必须具名**：引言点明报送/委托方全称（如"某农牧业开发有限公司"），禁用泛称"施工方"。
3. **落款用真实单位全称**（如"某造价咨询单位"），日期填实际出具日，禁用"复核小组"等占位。

---

## 一、格式规范 → 唯一来源：技能 `gongwen-doc-format`

**本 Skill 只管"写什么"（内容范式），不管"怎么排"（版式）。**

字体 / 字号 / 页边距 / 行距 / 缩进 / **目录** / 页码 / 表格列宽 / 标题防落单 / 标点规范——
**全部版式与文本规范规则，一律以 `gongwen-doc-format` 为准，本文件不另存副本。**

> **历史教训**：此前两处各存一份规格，其中一份边距记错（1.0/1.25cm）导致返工。
> **规格只允许存一处**；需要时指章节，不复制数字。
> ⚠️ 即便 `gongwen-doc-format` 中的数字也要**现读母本实测**，不凭记忆。

> 下文"二、内容结构范式"的代码块中保留了 `（黑体18居中）` `（楷体16）` 一类标注，
> **仅作"层级 ↔ 版式"的对应索引**，具体规格以 `gongwen-doc-format` 为准。

---

## 二、内容结构范式（以进度复核报告为例）

```
标题：<项目名> / <报告名>（黑体18居中）
引言（仿宋16）：根据<报送主体全称>报送的<资料名>及相关资料，我们对<项目名>工程进度情况进行了现场复核，现将复核情况说明如下：
一、项目基本情况（楷体16）
  工程名称：…
  工程地点：…
  项目内容：…
二、工程进度复核情况说明（楷体16）
  经现场复核，本项目已完成的主要工程内容如下：
  1、<已完成项1>；
  2、<已完成项2>；
  3、<已完成项3>；
  4、<未完成项并列作为末条>尚未完成。
三、复核结论（楷体16）
  经复核，本项目实际完成情况与申报进度一致，已完成工程量属实，整体完成比例约<XX>%，未完成部分主要为<…>，建议按合同约定继续推进后续工作，完成后及时组织验收。（后附现场施工影像）
落款（仿宋16右对齐）：复核单位：<真实单位全称>
                      <出具日期，另起一行、靠右>
```

**关键范式要点：**
- 已完成内容与进度说明**合并为一章**，未完成项作为末条并列，不拆章。
- 整体完成比例属结论性数据，**放在"复核结论"章**，不在事实陈述章。
- 结论须标注"**（后附现场施工影像）**"体现现场核实证据。
- 控制在 1–2 页，结构宁紧凑勿割裂。

---

## 三、生成方法（python-docx，直接格式化）

用 managed venv：`C:\Users\lihaoyuan\.workbuddy\binaries\python\envs\default\Scripts\python.exe`

核心片段（必须逐 run 设 eastAsia；**字体名与字号以 `gongwen-doc-format` 为准，下方默认值仅为示例**）：
```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

def add_para(text, font_name='仿宋', size=16, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if indent:
        p.paragraph_format.first_line_indent = Pt(32)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)  # 关键：中文字体
    return p
```
完整可运行脚本：
- 短公文（新建 1–2 页）→ 本 Skill 的 `references/format-spec.md`
- 长文档套版 / 目录页码 / 核验 → 技能 `gongwen-doc-format` 的 `scripts/`（`format_docx.py` / `update_toc_fields.py` / `verify_format.py`）

---

## 四、母本 / 参考来源

| 文档 | 路径 | 用途 |
|---|---|---|
| 某县情况说明母本（标准模板） | `<项目工作区>\某国家现代农业产业园蛋鸡产村融合发展示范项目\审核过程文件\某国家现代农业产业园蛋鸡产村融合发展示范项目初步审核情况说明.docx` | 情况说明类**格式标准母本**（版式规格见 `gongwen-doc-format`） |
| 某县情况说明（格式参考） | 桌面 `某县高标准农田建设项目初步审核情况说明.docx` | 进度复核等通用报告格式参考 |

> 用户已实测：editor_sdk 无头 `doc_get_document_content` 对本机该类模板不稳定，一律回退 python-docx 直设。

---

## 五、标准工作流

1. **版式规则先读技能 `gongwen-doc-format`**（边距/字体/行距一律以它为准，并现读母本实测）。
2. 读取用户提供的数据源（进度确认单 / 结算书 / 复核报告）。
3. 按"二、内容结构范式"组织文案，落实主体具名、现场复核、比例归结论、后附影像、真实落款。
4. 按"三、生成方法"用 python-docx 直设字体生成 .docx。
5. 校验：页数 1–2 页；字体/缩进/对齐符合 `gongwen-doc-format` 的规格；落款为真实单位全称。
6. present_files 交付，提示图片需用户自行附入（docx 内嵌图 python-docx 可写，但现场照片常由用户另附）。
