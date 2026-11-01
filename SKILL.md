---
name: 语言点翻译材料制作-E.D.STUDIO
description: "把英语考试类 .docx（阅读理解/七选五/完形填空/语法填空）批量转换成「语言点翻译」教学材料。从文档中提取粗体英文词块，自动翻译并归类到各篇目（A篇/B篇/七选五/完形填空/语法填空），为每篇写出【细化的文本话题】（通读正文后的概括性主题句，禁用笼统类别词）与【文本体裁】，生成两个独立文档：背诵版（中英对照，含词性）和默写版（英文留空供学生填写）。内置有道翻译兜底、双语词块配对、结构化分篇检测、选项行污染过滤。当用户提交英语试题 docx 并要求「挖空翻译/语言点/背诵版+默写版/中英对照表」时触发。"
agent_created: true
---

# 语言点翻译材料制作

把一份英语考试 `.docx` 一键转成两套教学材料：**背诵版**（中文释义 + 词性 + 英文，即对答案版）和**默写版**（中文 + 词性在左，英文留空在右，供学生自测）。

## 何时使用

- 用户提交英语试题 `.docx`，要求做「挖空翻译 / 语言点 / 背诵版 + 默写版 / 中英对照表」。
- 源文件典型结构：阅读理解 A–D 篇、七选五、完形填空、语法填空，正文中的**粗体英文**即为要提取的「语言点」词块。
- 有些源文件是**双语词块表**（粗体英文后紧跟粗体中文释义，如 `rented property 出租屋`），本流程会直接把中文当作释义，不调用网络翻译。

## 硬性交付约定（务必遵守，用户反复强调）

1. **必须产出两个独立文档**：`XX语言点背诵版.docx` 与 `XX语言点默写版.docx`。
2. **答案表格式**：左列 = 词性(方括号前置) + 中文释义（如 `[n.] 出租屋`；一词多性全列如 `[v.] 促使；导致 [adj.] 迅速的；及时的`），右列 = 英文。`背诵版` 两列都填；`默写版` 右列留空（下划线/空格）供填写。**单词类条目必带词性、短语/搭配不标词性**（详见下方「词性标注」一节）。
3. **篇头用纯文本标题**：如 `A篇`、`B篇`、`七选五`、`完形填空`、`语法填空`，不要花哨样式。
4. **【文本话题】与【文本体裁】绝不允许留空（`＿＿＿＿＿`）**：必须直接显示在背诵版和默写版中（每篇目一个 merged 行）。
   - **细化文本话题（必须）**：话题**不能是**笼统的类别词（如「社会与现象」「人物与故事」「健康与生活」）。必须是 Agent **通读该篇正文**后写出的**概括性主题句**（通常 1 句，点明文章核心内容 / 主旨）。示例：
     - `免费社区物品循环平台（Freegle）的运作机制与本地资源共享`
     - `潮汕电影《见你》借“侨批”符号彰显传统美学与跨时空乡情`
     - `尼日利亚留学生因名字“Goodluck”在美国的文化认同与归属感`
     - `社交媒体上公众对名人“虚假亲密感”的成因、影响与边界`
     - `食物线索诱发过度进食的大脑奖赏机制与自控研究`
   - **标准做法**：用 `_detect_section_change` + `_iter_paragraphs` 定位各篇正文 → 通读原文 → 为每篇写一句主题概括 → 通过 `process_docx_bytes_two(..., section_meta={篇目:{'topic':<细化句>,'genre':..}})` 传入，文档即采用该准确话题（主入口已支持此参数）。**禁止直接把 8 类类别词当话题输出**（那是退化兜底，不是默认）。
   - 体裁（5 类，可用类别词）：应用文 / 新闻报道 / 议论文 / 记叙文 / 说明文；**完形填空强制判为记叙文**。
   - 兜底例外：若源文件仅有词汇表、无正文可读（如 `8.23 限时训练` 类纯词块表），无法概括主旨时，才退化为 8 类话题之一（校园生活 / 科技与网络 / 环境与自然 / 健康与生活 / 文化与传统 / 社会与现象 / 体育与活动 / 人物与故事），但需向用户说明。

## 词性标注（通用原则，所有来源）

**单词类**（单个英文词，不含空格）语言点，其中文释义**必须带对应词性**；若一词多性，所有词性**全部列出**。

- **格式**：词性用方括号置于中文释义**之前**——`[n.] 释义` / `[v.] 释义` / `[adj.] 释义` / `[adv.] 释义`；一词多性连续标注：`[v.] 促使；导致 [adj.] 迅速的；及时的 [n.] 提示`。
- **仅限单词类**：含空格的短语 / 固定搭配（如 `be packed with`、`run out of`、`in advance`、`prompt` 作为整词除外）**不标词性**——它们是短语不是单词。
- **数据来源（两套，优先级递减）**：
  1. **离线权威主词典 `_POS`**（在 `cloze_core.py` 中）：键为 `normalize_collocation` 后的基本形式（小写），值为已嵌入词性 + 释义的完整中文串（含多性）。**首选**；Agent 遇到考试高频词应持续往里补充准确词性与完整多性。
  2. `get_pos(word)` 本地后缀 / 常见词规则兜底：仅给**主词性**前缀 `[x.]`，不保证多性完整，且对 `-ing/-ed/-or` 等易混后缀可能误判——因此优先补全 `_POS`。
- **引擎已自动套用**：`_build_recite_docx` / `_build_dictate_docx` 在中文列统一调用 `apply_pos(en, cn)`——单词命中 `_POS` 或 `get_pos` 即前缀标注；短语原样返回。背诵版、默写版中文列都带词性。
- **验收**：单词类条目 100% 带词性；一词多性无遗漏；短语类不出现词性标签。

## 输入文件约定

- 语言点 = 正文中的**粗体英文 run**。非粗体 run 原样保留（不参与提取）。
- 双语源：`连续粗体英文 run` + `紧随其后的粗体中文 run` = 一个「英文术语 + 中文释义」词对，直接取中文为释义，**不联网翻译**。
- 纯中文粗体（无配对英文）→ 调用 `translate_zh_en` 翻成英文。
- 全大写多词标题、纯标点噪声 → 过滤，不翻译。

## 分篇检测（关键逻辑）

结构化试卷标题通过 `_detect_section_change(text, runs_len)` 识别：
- `第一部分 阅读` / `第二部分 语言运用` → 标记 `READ` / `LANG` 大块。
- `第一节` → 完形填空；`第二节` → 七选五（READ 块内）或 语法填空（LANG 块内）。
- `A/B/C/D` 篇、或含「从短文后的选项中」「选出可以填入空白处」「A、B、C、D 四个选项」等提示语 → 对应分篇。
- 别名：`完型填空`（"型"）同样识别。
- `_iter_paragraphs` 同时遍历正文段落**与表格单元格**内的段落（七选五/语法填空常在表格里）。
- **选项行自动过滤（关键）**：完整试卷里，选择题的 A./B./C./D. 选项常把正确答案**粗体高亮**，这些不是语言点。提取时自动跳过「以单个字母 A–G + '.'/'、' 开头」或「内联含 A. B. C./D.」的段落（见 `_is_option_line`）。这样能排除谚语选项（如 `Virtue is its own reward`）等污染，只保留篇章/完形/语法填空/读后续写正文里的真词汇。

## 翻译策略（沙箱内网络受限，务必按此顺序）

1. **本地词典优先**：`cloze_core.py` 中的 `TRANSLATIONS` 大字典（小写键），命中即返回，**不依赖网络**。考试专属短语要持续往里加。
2. **有道翻译（主）**：`dict.youdao.com/jsonapi`，英文→中文取 `ec.word[].trs[].tr[].l.i`，中文→英文取 `ce` 字段；超时 6s。
3. **MyMemory（兜底）**：多数情况下被墙，仅作降级。
4. **Google gtx**：沙箱内不可用，不要依赖。
5. 大小写不敏感：提取到的英文可能是 Title Case，查 `TRANSLATIONS` 前先 `lower()`。
6. 去掉 MyMemory 偶发返回的 `v.`/`n.` 等词性前缀（用 `_POS_PREFIX` 正则）。

## 外刊精读类源文件处理流程（语言点=表格，整句=编号列表）

当源文件是**外刊精读语篇**（如「新高考英语外刊精读 U5 文化习俗 语篇文本.docx」）时，结构与普通试卷不同，采用**混合版式**：
- 正文里的**粗体英文**是语言点；中文释义常**内嵌在括号内**随文出现（如 `an interwoven tapestry (织锦) of…`、`aural (听觉的) cues`），不再是独立的「英文+粗体中文」词对。
- 文档**已预先写好**【文本话题】【文本体裁】且为细化主题句——**直接复用**，不要再自动概括也不要留空。
- **语言点（非整句粗体）用「表格形式」呈现**（与 203/303 标准表格版一致：背诵版 中文｜英文 两列；默写版 中文宽列｜空白 一列）；**完整句单独归入末尾【句子翻译】板块，用编号列表**。

### 输出格式（语言点=表格，整句=编号列表）
- **语言点（非整句）表格**：每个篇目一张表（或同表分段），结构为 `标题(合并行)` → `篇头(合并行)` → `【文本话题】：…(合并行)` → `【文本体裁】：…(合并行)` → 逐条 `中文 ｜ 英文`（背诵版）/ `中文宽列(合并) ｜ 空白`（默写版）。背诵版两列都填；默写版英文列留空。表格列宽与 203/303 一致（背诵版 8cm×2；默写版 5.3+5.3+5.0cm）。
- **完整句 →【句子翻译】编号列表**（见下节），位于文档最末尾，**不用表格**：
  - **背诵版**：`N. 中文翻译` 换行 `英文原句`。
  - **默写版**：`N. 中文翻译` 换行后提供 **2 行行间距 1.5 倍的空行**（供学生写英文），不出现英文。
- 字体：中文 **华文楷体** / 英文 **Times New Roman**，均 **10.5pt**，段行距 **1.5 倍**（全局排版规范）。
- 注：语言点表格**不显示序号**（行序即序号）；【句子翻译】保留全局连续编号 1,2,3…。

### 固定搭配标准化（Collocation Normalization）—— 通用原则，所有来源必须做
处理**所有来源**的粗体内容（非完整句）时，将表面时态/人称形式还原为**词组基本形式**——这是通用原则，标准套卷与外刊精读同样适用（引擎已在 背诵版 英文列自动套用 `normalize_collocation`，无需手工处理）：
1. **谓语时态还原**：一般过去时 / 三单 / 完成时 → 动词原形（`played → play`、`contributes → contribute`、`has contributed → contribute`）。
2. **be 动词还原**：`is/was/are/were + 补语` → `be + 补语`（`is characterized by → be characterized by`、`was purely for → be purely for`、`is responsible for → be responsible for`）。
3. **还原的是谓语的时态/人称形式，不是删 be**：`characterized by → ❌`，`be characterized by → ✓`。
4. **非谓语保持搭配要求**：介词后 `doing`、不定式 `to do` 保留（`contribute to doing`、`aim to do`）；`being hindered by` 类被动分词 → 还原为 `be hindered by`。
5. 先 `strip_paren_glosses` 去掉内嵌中文括号，再输出纯英文基本形式。
- 引擎已提供 `normalize_collocation(s)`（保守实现：只还原确知为动词的表面形式，绝不误伤名词如 `stories` / 形容词如 `naive`；含 `_VERB_BASE` 与 `_PART_NEEDS_BE` 两张映射表）。**Agent 仍需逐条复核**，对映射表未覆盖的动词表面形式手动改成原形。

### 完整句 →【句子翻译】板块
- 判定「完整句」：以 `. ? !` 结尾、含主语+限定动词、能独立成句的粗体内容（如 `Language doesn't fossilize`、`As I considered these issues, I realized that…`）。**祈使句小标题**（如 `Be open-minded and adaptable`）与名词/动词短语**不算**完整句，留在语言点列表。
- 在文档**最末尾**另开一个 **【句子翻译】** 板块，把所有完整句**按出现顺序连续编号**（1,2,3…），可按原篇目加小标题（如 `（阅读 C 篇）`）。
- 格式与语言点一致：**背诵版** `N. 中文翻译` 换行 `英文原句`；**默写版** `N. 中文翻译` 换行后 2 行 1.5 倍空行。

### 实现提示
- **语言点走表格**：复用 `cloze_core` 的 `_build_recite_docx` / `_build_dictate_docx` 思路（或同款合并行 + 单元格字体 helper），每篇目一张表；英文基本形式先 `normalize_collocation`（非整句）。
- **整句走编号列表**：遍历完整句 → 末尾另起【句子翻译】板块，`add_paragraph`（`line_spacing=1.5`，run 设 华文楷体/ Times New Roman 10.5pt）；默写版每条后加 2 个空 `Paragraph`（行距 1.5）。
- 英文原句同样先 `strip_paren_glosses` 去中文括号。

## 完整句分流与【句子翻译】板块（通用，适用于所有来源）

> **自 2026-11-01 起本规则对所有来源通用**，不再限于外刊精读。标准试卷的粗体里常混有完整句（实测占 10%–30%），若与词块同表，学生要背的东西没有语法骨架，复习价值低。

### 判定标准（Agent 必须逐条人工判定，不得只靠标点）

**是完整句**（应分流）：含主语 + 限定动词、能独立成句。
- `what he created was anything but conventional wooden furniture that…`（宾语从句）
- `when the translation reaches readers, it gains new life`（状语从句 + 主句，**无句末标点也成立**）
- `Start with something simple and build from there.`（祈使句）
- `Life is too short to let ideas stagnate; now's the time to turn them to reality.`

**不是完整句**（留在语言点表）：
- 动词短语 / 名词短语：`invested such extraordinary effort into…`、`distinguished poets`
- **半句**（谓语在原文下一段）：`What struck me most was` ← 极易误判
- 从句内的修饰性成分：`a master craftsman whose hands told stories of decades working with wood`

### 调用方式

```python
recite, dictate = process_docx_bytes_two(
    data, base='卷名',
    section_meta={...},                      # 篇目话题/体裁
    sentences={'B': [英文原句, ...], ...},    # 完整句（按篇目分组）
    keep_raw=['What struck me most was', ...] # 豁免还原的条目
)
```

引擎行为：
- `sentences` 中的条目**从语言点表格中移除**，末尾另起 **【句子翻译】** 板块，按 `_SENTENCE_ORDER`（A→B→C→D→七选五→完形填空→语法填空）顺序、**全局连续编号**。
- 背诵版：`N. 中文翻译 （篇目）` 换行 `英文原句`；默写版：`N. 中文翻译 （篇目）` 换行后 **2 行 1.5 倍行距空行**（供学生写英文）。
- 完整句的英文**不做固定搭配标准化**，保持原貌。
- 传入 `sentences=None`（默认）时**不产生该板块**，既有产出完全不变。

### keep_raw：豁免「固定搭配标准化」

固定搭配还原是通用原则，但**只适用于搭配/短语**；若粗体内容是**句子的谓语成分**，还原会改变原句语法。实测两类典型误伤：

| 原句（还原后错误） | 原因 | 处置 |
|---|---|---|
| `hands told stories` → `hands tell stories` | 从句内过去时谓语被当动词还原 | 加进 `keep_raw` |
| `What struck me most was` → `What struck me most be` | 半句被当短语还原 be 动词 | 加进 `keep_raw` |

判定口诀：**还原后读起来仍是一句语法正确的话，才是对的**。逐条复核 `normalize_collocation` 的输出，发现异常即加入 `keep_raw`。

## 核心 API（见 references/cloze_core.py）

- `process_docx_bytes_two(input_bytes, base='', section_meta=None, sentences=None, keep_raw=None)` → 返回 `(背诵版_bytes, 默写版_bytes)`。**这是主入口**。
  - `section_meta` 为可选覆盖 dict：`{篇目:{'topic':..,'genre':..}}`；传入则直接采用（保证话题/体裁准确），为 `None` 时按关键词自动概括。
  - `sentences` 为可选 dict：`{篇目:[英文原句, ...]}` —— 完整句从语言点表格分离，末尾以【句子翻译】编号列表呈现；为 `None`（默认）时不产生该板块，**既有行为完全不变**。
  - `keep_raw` 为可选 list：逐条豁免「固定搭配标准化」的条目（还原会改变原句语法者，见下节）。
- `process_docx_bytes_table(input_bytes, base='')` → 仅默写版单文档（向后兼容）。
- `classify_section_meta(text, section=None)` → `{'topic':..., 'genre':...}`。
- `translate(en_text)` / `translate_zh_en(zh_text)` → 单条翻译。
- `apply_pos(en, cn)` → 单词类条目词性标注主入口：优先用离线权威 `_POS`（含多性完整串），否则 `get_pos` 兜底前缀 `[x.]`；短语原样返回。引擎在建表时自动调用，无需手工处理。
- `get_pos(word)` → 纯本地词性判断（不联网，仅主词性，易混后缀可能误判）。
- `normalize_collocation(s)` / `strip_paren_glosses(s)` → 将粗体词块（非整句）还原为词组基本形式、去掉内嵌中文括号；**通用原则，作用于所有来源**（引擎在 背诵版 英文列自动套用）。详见上方「固定搭配标准化」章节。
- 运行依赖（managed venv）：`python-docx`、`requests`。
  - venv: `~/.workbuddy/binaries/python/envs/default/bin/python`（WorkBuddy 自带 managed venv）
  - 示例：
    ```python
    from references.cloze_core import process_docx_bytes_two
    data = open('源文件.docx','rb').read()
    recite, dictate = process_docx_bytes_two(data, base='8.23 限时训练试题')
    open('8.23 限时训练试题语言点背诵版.docx','wb').write(recite)
    open('8.23 限时训练试题语言点默写版.docx','wb').write(dictate)
    ```

## 网页版 / 在线链接（见 references/cloze_app.html + cloze_web_v3.py）

- `cloze_app.html`：纯前端（JSZip + DOMParser）在浏览器里做同样的处理，逻辑与 Python 版对齐（`detectSectionChange` / `translateZhEn` / `localTrans` / `hasChinese` / `classifySectionMeta`）。
- `cloze_web_v3.py`：把 `cloze_app.html` 内嵌进 `HTML = r'''...'''`，启动静态服务（默认 **端口 7860**），表格模式返回两个 docx 的 ZIP。
- **同步原则**：改了 `cloze_app.html` 后，必须重新把它的内容写回 `cloze_web_v3.py` 的 `HTML` 常量（注意 html 内不能有未转义的三引号），再重启服务。
- 提供在线链接：本地跑服务后用 `present_files` 给出 `http://localhost:7860/`；如需他机可访问，用 CloudStudio 部署（`workbuddy_cloudstudio_deploy`）。

## 常见坑（按经验必看）

1. **网络翻译失败 → 大量「待译」**：先确认用的是有道而非 MyMemory/Google；考试专属短语补进 `TRANSLATIONS`。
2. **双语源被合并导致翻译失败**：旧逻辑会把 `rented property出租屋` 当一个英文串，需配对「连续粗体英文 + 紧随粗体中文」为词对。
3. **大小写漏翻**：`TRANSLATIONS` 键是小写，提取到的英文是 Title Case → 查前先 lower。
4. **分篇漏检（七选五/完形/语填缺失）**：只识别了 A–D 字母；要用结构化标题 + `_iter_paragraphs` 遍历表格单元格。
5. **后台生成被 SIGKILL（137）**：用工具的 `run_in_background=true` 启动，**不要**用 Bash 的 `&`（父 shell 退出即被杀）。长任务用分阶段脚本 + pickle 缓存续跑。
6. **话题/体裁留空**：用户明确禁止；确保 `_build_section_meta` 在两份文档都输出概括结果，且兜底类别非空。
7. **选项行污染**：完整试卷里 A./B./C./D. 答案选项若被粗体高亮，会被误当语言点。现已自动跳过选项行（`_is_option_line`），生成后仍需抽查——若背诵版出现谚语/整句选项（如 `Virtue is its own reward`），说明该源格式例外，需补过滤规则。
8. **话题/体裁关键词误判**：自动分类对完整试卷常错（说明文↔议论文、文化↔人物等）。推荐做法：Agent 通读正文后通过 `section_meta` 覆盖准确标签，而非依赖自动分类。
9. **英文句含内嵌中文括号 → 被误判为「纯中文词块」**（2026-11-01 修复）：粗体 `Life is too short to let ideas stagnate (停止); now's…` 因含 CJK 字符而走进「纯中文」分支，结果**中文列直接显示英文原文**、翻译彻底失败。修复：改用 `_cjk_dominant()` 判断「中文是否占优」（CJK 字符数 > 拉丁字母数），含内嵌中文注释的英文句仍走英文路径。
10. **完整句与词块混在一张表里**（2026-11-01 起通用）：见「完整句分流与【句子翻译】板块」章节。关键词块之外必须人工逐条判定完整句，并传入 `sentences` 参数；否则完整句会同时出现在表格和翻译板块（重复）。
11. **固定搭配还原误伤句中谓语**：`hands told stories` → `hands tell stories`、`What struck me most was` → `What struck me most be`。见 `keep_raw` 用法。
12. **翻译键必须与源文逐字节一致**：源文常含**全角句号「。」**与**右单引号「’」**（如 `True craftsmanship isn't about… did。It's about…`）。手写翻译键极易不匹配——应**从源文提取原串再构造键**，否则 `translate()` 查表落空，整句翻译失败。

## 文件清单

- `references/cloze_core.py` —— Python 处理引擎（提取/翻译/分篇/概括/完整句分流/成 docx），主入口 `process_docx_bytes_two`。
- `references/cloze_app.html` —— 浏览器端等价实现。
- `references/cloze_web_v3.py` —— 内嵌 html 的本地网页服务（端口 7860）。

## 验收标准

- 两份文档均生成，命名含「背诵版」「默写版」。
- **表格版**（普通试卷源）：背诵版左(词性前置 + 中文) 右(英文) 对照；默写版右列留空；**单词类条目 100% 带词性、一词多性全列、短语不标词性**。
- **混合版式**（含完整句的来源）：语言点（非整句）为**表格**（背诵版 中文｜英文；默写版 中文宽列｜空白），每篇目 `【文本话题】/【文本体裁】` 齐全；末尾有 **【句子翻译】** 板块（编号列表、全局连续编号，背诵版附英文原句、默写版留空 2 行 1.5 倍行距）；**完整句不得同时出现在语言点表格中**（重复即不合格）；英文基本形式已做**固定搭配标准化**（be→be、时态/三单还原、非谓语保留），且被还原破坏语法的条目已列入 `keep_raw`。
- 每篇目均含【文本话题】【文本体裁】且不为空；**话题须为细化主题句（非笼统类别词）**。
- 抽取条目 0 待译（或仅余极少数需补 `TRANSLATIONS`）；话题/体裁 0 留空。
- **回归要求**：引擎改动后，须对既有材料重跑抽取对比（条目集合应完全一致），确认零破坏。
