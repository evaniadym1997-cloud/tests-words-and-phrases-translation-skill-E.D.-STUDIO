# 语言点翻译材料制作（E.D. STUDIO）

> Tests, Words and Phrases Translation Skill — 把英语试卷 / 外刊精读 `.docx` 里的**粗体英文语言点**，一键产出 **背诵版 + 默写版** 两份 Word 文档。

[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Skill](https://img.shields.io/badge/type-WorkBuddy%20Skill-orange.svg)](SKILL.md)

---

## 这是什么

给一份含粗体语言点的 Word 试卷（套卷 / 限时训练 / 外刊精读语篇），自动生成：

| 文档 | 版式 | 用途 |
|---|---|---|
| **背诵版** | `中文(含词性)` ｜ `英文` 两列 | 学生背记、老师核对 |
| **默写版** | `中文宽列` ｜ `空白` 三列 | 学生默写英文 |

两者**同目录生成**，篇目结构一致：

```
阅读 B 篇                    ← 篇头（合并行）
【文本话题】：……             ← 描述性主题句（概括本文讲什么）
【文本体裁】：记叙文
  挤满；充满                    ← 中文释义（单词类带 [词性]）
  be packed with               ← 英文（已还原为基本形式）
```

---

## 核心特性

### 1️⃣ 固定搭配标准化（Collocation Normalization）

把粗体内容的**表面时态 / 人称形式**还原为**词组基本形式**：

| 原文表面形式 | 输出基本形式 |
|---|---|
| `is packed with` | `be packed with` |
| `caters to all tastes` | `cater to all tastes` |
| `came to believe` | `come to believe` |
| `lost all faith in` | `lose all faith in` |
| `weathered the storm` | `weather the storm` |
| `is characterized by` | `be characterized by` |
| `aim to do` / `contribute to doing` | ✅ 非谓语形式完整保留 |

- **be 动词**任意位置 → `be`
- **have/has/had**：助动词 → 还原后继动词原形或 `be`；实义 → `have`
- 三单 / 过去式 / 过去分词 → 原形
- 开头被动分词 → 补 `be`
- **绝不误伤**名词（`stories`）、形容词（`naive`）、短语（`purely for`）
- **整句守卫**：含句末标点的完整句跳过还原，句子结构不被破坏

### 2️⃣ 词性标注

- 单词类条目：中文释义前**方括号前置**词性，`[n.] / [v.] / [adj.] / [adv.] / [prep.]`
- **一词多性全部列出**：`[v.] 促使；导致 [adj.] 迅速的；及时的 [n.] 提示`
- **短语 / 搭配不标词性**（它们不是单词）

### 3️⃣ 细化文本话题 + 体裁

每篇目必须有**描述性主题句**（概括本文讲什么），而非笼统类别词：

| ❌ 笼统 | ✅ 细化 |
|---|---|
| 社会现象 | 剖析「独处美学」的社交媒体叙事如何遮蔽真实孤独 |
| 文化与传统 | 爆米花与电影绑定关系的百年历史渊源及其在大萧条中的演变 |

体裁：应用文 / 记叙文 / 说明文 / 议论文。

### 4️⃣ 0 待译

纯英文源（无双语词对）通过**离线精选词典** `TRANSLATIONS` + 权威词性词典 `_POS` 确定性生成，绝不留 `【待译】`。

---

## 安装与使用

### 作为 WorkBuddy 技能使用

```bash
# 复制到技能目录
cp -r "tests-words-and-phrases-translation-skill-E.D.-STUDIO" ~/.workbuddy/skills/
```

然后在 WorkBuddy 中说：「用语言点翻译材料制作技能处理这个 .docx」。

### 作为 Python 库使用

```bash
pip install python-docx requests
```

```python
from cloze_core import process_docx_bytes_two

data = open('试题.docx', 'rb').read()

section_meta = {
    'A': {'topic': '伦敦 2026 年度精选戏剧演出指南', 'genre': '应用文'},
    'B': {'topic': '作者回忆童年经济拮据的圣诞节', 'genre': '记叙文'},
}

recite, dictate = process_docx_bytes_two(data, base='套卷语言点', section_meta=section_meta)

open('套卷语言点背诵版.docx', 'wb').write(recite)
open('套卷语言点默写版.docx', 'wb').write(dictate)
```

### CLI

```bash
python references/cloze_core.py 输入.docx 输出.docx
```

---

## 仓库结构

```
.
├── SKILL.md                    # 技能说明（完整规则与验收标准）
├── README.md                   # 本文件
├── LICENSE                     # MIT
├── requirements.txt            # 依赖
└── references/
    └── cloze_core.py           # 核心引擎（提取 → 标准化 → 建文档）
```

---

## 核心 API

| 函数 | 作用 |
|---|---|
| `process_docx_bytes_two(input_bytes, base='', section_meta=None)` | 一键产出 `(背诵版, 默写版)` 字节流 |
| `_collect_translations(input_bytes)` | 从粗体提取 `(en, cn, section, pos)` |
| `_build_recite_docx(sections, title, section_meta)` | 构建背诵版 |
| `_build_dictate_docx(sections, title, section_meta)` | 构建默写版 |
| `normalize_collocation(s)` | 固定搭配标准化（基本形式还原） |
| `apply_pos(en, cn)` | 词性标注（方括号前置） |
| `strip_paren_glosses(s)` | 去掉内嵌中文括号 |

---

## 排版规范

- 中文 **华文楷体** / 英文 **Times New Roman**
- 字号 **10.5 pt**，行距 **1.5 倍**
- 默写版留白区域供学生书写

---

## 验收标准

- [ ] 双文档已产出，命名正确（无「语言点语言点」重复）
- [ ] 0 个 `【待译】`
- [ ] 每个篇目【文本话题】为描述性主题句、【文本体裁】非空
- [ ] 单词类条目全部带 `[词性]`，一词多性全列出；短语不标词性
- [ ] 默写版英文列全空白
- [ ] 非整句粗体英文已还原为基本形式
- [ ] 字体 / 字号 / 行距符合规范

---

## License

MIT — 见 [LICENSE](LICENSE)。
