#!/usr/bin/env python3
"""
DOCX 挖空处理器 - 核心处理模块
策略：
1. 遍历段落的所有 run
2. 粗体 run：替换为 下划线空格 + （中文）的 run
3. 非粗体 run：完全保留，不修改
4. 用 python-docx 的 XML 层操作实现精确控制
"""

from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
import re
import requests
import io

# ── 翻译 ────────────────────────────────────────────────────────────────────
# 本地词典：API 失败/限流时的兜底，优先命中避免依赖网络
TRANSLATIONS = {
    # 通用阅读词汇
    'affiliate': '附属机构', 'serve': '担任；服务', 'serve as': '充当；作为',
    'stepping stone': '踏脚石；垫脚石', 'rising stars': '冉冉升起的新星',
    'giveaway': '赠品；赠予', 'postpone': '推迟；延期', 'postponed': '推迟的；延期的',
    'reschedule': '重新安排', 'rescheduled': '已改期的', 'crucial': '关键的；至关重要的',
    'promising': '有前途的；有希望的', 'atmosphere': '氛围；气氛', 'promotion': '促销活动',
    'discount': '折扣', 'refund': '退款', 'postpone': '推迟', 'reschedule': '改期',
    'venue': '场所；场地', 'theme': '主题', 'fireworks': '烟花', 'senior': '老年人',
    'reserved': '预订的；保留的', 'stadium': '体育场', 'ticket': '票',
    'occur': '发生', 'occur as scheduled': '如期举行', 'holder': '持有者',
    'attend': '参加；出席', 'weather': '天气', 'reason': '原因；理由',
    'postponed or rescheduled': '推迟或改期',
    # A.docx 专用短语/词块（确保本地即可翻译，不依赖 API）
    'as a crucial stepping stone': '作为关键的踏脚石',
    'rising stars in action': '冉冉升起的新星在',
    'a game is postponed': '一场比赛被推迟',
    'purchase': '购买；采购',
    # 英语试题1 专用短语/词块
    'waterproof': '防水的', 'tear-resistant': '抗撕裂的', 'their eyes': '他们的眼睛',
    'light up': '点亮；亮起', 'is intended for': '适用于；旨在', 'step-by-step instructions': '分步说明',
    'Scientific spirit fuels': '科学精神推动', 'enthusiasm': '热情；积极性', 'pave': '铺平（道路）',
    'the way for a fresh perspective on': '以全新视角看待', 'Pair': '配对；搭配', 'with': '和；与；用',
    'outperform': '胜过；表现优于', 'liken': '把…比作；比喻', 'to': '向；到',
    'a wise and environmentally responsible use': '明智且对环境负责的使用', 'safeguard': '保护；维护',
    'and advance': '并推进；促进', 'responsible ocean development': '负责任的海洋开发',
    'Dismissive': '轻蔑的；不屑一顾的', 'Conservative': '保守的', 'exemplifies': '是…的典范；例证',
    'relatable and dynamic to readers': '让读者有共鸣且生动', 'heroic ideals': '英雄理想',
    'see humanity in a more complex way': '以更复杂的方式看待人性',
    'for the exploration of personal identity and internal conflict': '用于探索个人身份与内心冲突',
    'reveal human': '揭示人性', 'complexity, showing life beyond black and white': '复杂性，展现超越非黑即白的生活',
    'stand out for complexity and depth': '以复杂性与深度脱颖而出', 'originate': '起源；首创', 'in': '在…里；在',
    # 常用学术/考试词
    'however': '然而', 'therefore': '因此', 'moreover': '此外', 'furthermore': '而且',
    'nevertheless': '尽管如此', 'whereas': '然而；鉴于', 'although': '尽管',
    'despite': '尽管', 'according to': '根据', 'in order to': '为了',
    'as a result': '结果', 'for example': '例如', 'in addition': '此外',
    'on the other hand': '另一方面', 'in contrast': '相反', 'in fact': '事实上',
    'significant': '显著的', 'benefit': '好处；受益', 'beneficial': '有益的',
    'approach': '方法；接近', 'challenge': '挑战', 'opportunity': '机会',
    'potential': '潜力；潜在的', 'individual': '个人；个体的', 'society': '社会',
    'environment': '环境', 'technology': '科技', 'development': '发展',
    'increase': '增加', 'decrease': '减少', 'improve': '改善', 'affect': '影响',
    'reflect': '反映', 'suggest': '表明；建议', 'indicate': '表明', 'reveal': '揭示',
    'conclude': '得出结论', 'analyze': '分析', 'compare': '比较', 'contrast': '对比',
    'evidence': '证据', 'theory': '理论', 'research': '研究', 'study': '研究；学习',
    'factor': '因素', 'process': '过程；处理', 'system': '系统', 'resource': '资源',
    'global': '全球的', 'modern': '现代的', 'traditional': '传统的', 'culture': '文化',
    'economic': '经济的', 'political': '政治的', 'social': '社会的', 'personal': '个人的',
    'physical': '身体的；物理的', 'mental': '心理的', 'emotional': '情感的',
    'natural': '自然的', 'human': '人类的；人', 'world': '世界', 'life': '生活；生命',
    'experience': '经历；体验', 'knowledge': '知识', 'skill': '技能', 'ability': '能力',
    'behavior': '行为', 'attitude': '态度', 'value': '价值；重视', 'goal': '目标',
    'problem': '问题', 'solution': '解决方案', 'method': '方法', 'result': '结果',
    'reason': '原因', 'effect': '影响；效果', 'cause': '原因；导致', 'change': '改变',
    'create': '创造', 'provide': '提供', 'require': '需要', 'allow': '允许',
    'consider': '考虑；认为', 'believe': '相信；认为', 'expect': '期望；预期',
    'understand': '理解', 'recognize': '认识到；认出', 'realize': '意识到',
    'achieve': '实现；达成', 'maintain': '维持；保持', 'reduce': '减少',
    'prevent': '防止', 'protect': '保护', 'support': '支持', 'encourage': '鼓励',
    'influence': '影响', 'determine': '决定', 'establish': '建立', 'explain': '解释',
    'describe': '描述', 'demonstrate': '证明；展示', 'represent': '代表；表示',
    'focus': '聚焦；关注', 'emphasize': '强调', 'highlight': '突出；强调',
    'include': '包括', 'involve': '涉及；包含', 'contain': '包含',
    'contribute': '贡献；导致', 'lead to': '导致', 'result in': '导致',
    'depend on': '依赖', 'relate to': '与……相关', 'refer to': '指的是',
    'deal with': '处理；应对', 'cope with': '应对', 'respond to': '回应',
    'aware of': '意识到', 'based on': '基于', 'due to': '由于', 'instead of': '而不是',
    'rather than': '而不是', 'as well as': '以及', 'in terms of': '就……而言',
    # 8.23 限时训练语言点（源文件为双语，以下为无附中文释义的英文表达/短语）
    'trace and access': '追踪与查找（保险理赔术语）',
    'raising his eyebrows': '扬起眉毛（表惊讶/疑惑）',
    'bent down and kissed me gently on the cheek': '弯下腰轻轻吻了我的脸颊',
    'jumped right to it': '立刻着手做',
    'the catch': '陷阱；隐藏的问题',
    'a diverse next generation': '多元化的下一代',
    'been outspoken about': '对……直言不讳',
    'delivering new technologies to all people': '把新技术带给所有人',
    'pay attention to the demand side': '关注需求侧',
    'sharpen thinking skills': '锻炼思维技能',
    'body language collection': '肢体语言的收集（汇总）',
    'hug themselves, rub their forehead, and shift their body weight from one foot':
        '抱紧自己、揉额头、把重心在双脚间转移',
    'listening sympathetically to': '同情地倾听',
    'adopting a similar stance': '采取相似的姿态',
    'the person you are engaged with': '你与之互动的人',
    'faced with the': '面对……时',
    'advanced research and so should we all.': '前沿研究，我们也应如此',
    # 203高二练习26上 语境校准
    'engage': '（使）参与；从事；啮合，扣住',
    'dive into a body of research': '深入钻研；投身于大量研究',
    'went to his conservation efforts': '用于他的保护事业（投入精力/资金）',
    # 303高三练习26上 语言点（离线校准，避免依赖有道网络）
    'run on a very tight budget': '运行在极紧张的预算下；资金极为有限',
    'visibly frustrated': '明显受挫的；明显沮丧的',
    'from scratch': '从零开始；白手起家',
    'are scripted': '（对话/互动等）是事先安排好的；非即兴的',
    'caption': '（图片、视频的）说明文字',
    'feel more connected and less intimidated': '感到更亲近、更少拘束/畏惧',
    'consistent and accessible': '（沟通）稳定且易于获取的',
    'unconsciously shapes everything': '在不知不觉中塑造一切',
    'offer a sense of belonging': '带来归属感',
    'a constant stream of': '源源不断的；连续不断的（事物）',
    'appetite': '食欲；胃口',
    'when faced with': '当面对……时',
    'shrink': '（使）缩小；减少；收缩',
    'in the split second when': '在……的那一瞬间；转瞬之间',
    'step in': '介入；干预；出面（帮助或调解）',
    'prejudice can break down': '偏见可以消除（或瓦解）',
    'in my case': '就我而言；以我自己来说',
    'stopped blending in': '不再融入（群体）；不再显得格格不入',
    'clarification': '澄清；阐明',
    'a quiet sense of gratitude': '一种默默的/宁静的感恩之情',
    'deeply rooted in distinctive chaoshan culture': '深深植根于独特的潮汕文化之中',
    'highlights the aesthetic of subtlety': '凸显了含蓄之美（的美学）',
    'illiterate': '不识字的；文盲的',
    'go beyond time': '超越时间；历久弥新',
    'touch audiences of all ages across different lands': '打动不同地域、各个年龄层的观众',
}

# 去掉 MyMemory 偶尔返回的前缀词性标记（如 "v.送达" → "送达"）
import re as _re
_POS_PREFIX = _re.compile(
    r'^(v\.|n\.|a\.|adj\.|adv\.|prep\.|conj\.|pron\.|art\.|num\.|int\.|动词|名词|形容词|副词|介词|连词|代词)\s*',
    _re.I,
)

_trans_cache = {}

_youdao_session = None

def _get_youdao_session():
    global _youdao_session
    if _youdao_session is None:
        _youdao_session = requests.Session()
    return _youdao_session


def _youdao_translate(en_text):
    """有道词典 jsonapi（国内可达，作为主翻译源）。返回精简中文义项或 None。"""
    try:
        s = _get_youdao_session()
        r = s.get('https://dict.youdao.com/jsonapi', params={'q': en_text}, timeout=4)
        if r.status_code != 200:
            return None
        j = r.json()
        out = []
        ec = j.get('ec')
        if ec:
            for w in ec.get('word', []):
                for trs in w.get('trs', []):
                    for tr in trs.get('tr', []):
                        li = tr.get('l', {}).get('i')
                        if isinstance(li, list):
                            txt = ''.join(str(x) for x in li).strip()
                        elif isinstance(li, str):
                            txt = li.strip()
                        else:
                            txt = ''
                        if txt:
                            out.append(txt)
        if not out:
            wt = j.get('web_trans')
            if wt:
                for grp in wt.get('web-trans-group', []):
                    for it in grp.get('trans', []):
                        v = it.get('value', '')
                        if v:
                            out.append(v)
        if not out:
            return None
        # 取第一义项（分号前），并去除前置词性标记（如 "n. 宝库" → "宝库"）
        first = out[0].split('；')[0].split(';')[0].strip()
        first = re.sub(r'^[a-zA-Z]+\.\s*', '', first).strip()
        return first if first else None
    except Exception:
        return None


def translate(en_text):
    t = en_text.strip()
    if not t:
        return None
    if t in TRANSLATIONS:
        return TRANSLATIONS[t]
    tl = t.lower()
    if tl in TRANSLATIONS:
        return TRANSLATIONS[tl]
    cache_key = t.lower()
    if cache_key in _trans_cache:
        return _trans_cache[cache_key]
    result = None
    # 1) 有道（国内可达，优先）
    result = _youdao_translate(t)
    # 2) MyMemory 兜底（偶尔可用）
    if not result:
        try:
            r = requests.get(
                'https://api.mymemory.translated.net/get',
                params={'q': t, 'langpair': 'en|zh-CN'},
                timeout=4
            )
            data = r.json()
            cand = data.get('responseData', {}).get('translatedText', '')
            if cand and cand != t and not cand.startswith('MYMEMORY'):
                result = _POS_PREFIX.sub('', cand).strip()
        except Exception:
            pass
    _trans_cache[cache_key] = result
    return result


def translate_zh_en(zh_text):
    """中文 → 英文（有道 ce 字段，兜底 web_trans）。返回精简英文义项或 None。"""
    t = zh_text.strip()
    if not t or not _contains_cjk(t):
        return None
    cache_key = 'zh:' + t
    if cache_key in _trans_cache:
        return _trans_cache[cache_key]
    result = None
    try:
        s = _get_youdao_session()
        r = s.get('https://dict.youdao.com/jsonapi', params={'q': t}, timeout=4)
        if r.status_code == 200:
            j = r.json()
            out = []
            ce = j.get('ce')
            if ce:
                for w in ce.get('word', []):
                    for trs in w.get('trs', []):
                        for tr in trs.get('tr', []):
                            li = tr.get('l', {}).get('i')
                            if isinstance(li, list):
                                for x in li:
                                    if isinstance(x, dict) and '#text' in x:
                                        out.append(x['#text'])
                                    elif isinstance(x, str):
                                        out.append(x)
                            elif isinstance(li, str):
                                out.append(li)
            if not out:
                wt = j.get('web_trans')
                if wt:
                    for grp in wt.get('web-trans-group', []):
                        for it in grp.get('trans', []):
                            v = it.get('value', '')
                            if v:
                                out.append(v)
            if out:
                first = out[0].split('；')[0].split(';')[0].strip()
                result = first if first else None
    except Exception:
        result = None
    _trans_cache[cache_key] = result
    return result

def blank_display_len(en):
    """下划线显示长度（空格数）"""
    n = int(len(en.strip()) * 1.5)
    return max(8, min(50, n))

# ── 字体设置 ─────────────────────────────────────────────────────────────────
FONT_EN = 'Times New Roman'
FONT_ZH = '华文楷体'
FONT_SIZE = Pt(10.5)

def set_run_font(run, is_chinese=False):
    """设置 run 的字体和字号"""
    run.font.size = FONT_SIZE
    if is_chinese:
        run.font.name = FONT_ZH
        # 设置东亚字体
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = rPr.makeelement(qn('w:rFonts'))
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), FONT_ZH)
    else:
        run.font.name = FONT_EN

def has_chinese(text):
    return any('\u4e00' <= c <= '\u9fff' for c in (text or ''))

def _contains_cjk(text):
    """判断字符串是否含有中日韩字符（用于区分英文词块与中文释义）。"""
    return has_chinese(text)

# ── 固定搭配标准化（Collocation Normalization）─────────────────────────────────
# 规则（用户明确）：将粗体词块的表面时态/人称形式还原为「词组基本形式」。
#   1) 谓语时态还原：一般过去时 / 三单 / 完成时 → 动词原形（played→play, contributes→contribute, has contributed→contribute）
#   2) be 动词还原：is/was/are/were + 补语 → be + 补语（is characterized by→be characterized by）
#   3) 还原的是谓语的时态/人称形式，不是删 be：characterized by→❌ be characterized by→✓
#   4) 非谓语保持：介词后 doing、to do 保留（contribute to doing, aim to do）
_PAREN_GLOSS = re.compile(r'[（(][^）)]*[\u4e00-\u9fff][^）)]*[）)]')

def strip_paren_glosses(s):
    """去掉英文词块里内嵌的中文释义括号，如 an interwoven tapestry (织锦) of → an interwoven tapestry of"""
    return re.sub(r'\s+', ' ', _PAREN_GLOSS.sub(' ', s)).strip()

# 已知动词「表面形式 → 原形」映射（只还原确知为动词的词，绝不误伤名词如 stories）
_VERB_BASE = {
    'played': 'play', 'contributes': 'contribute', 'contributed': 'contribute',
    'witnessed': 'witness', 'established': 'establish', 'stimulated': 'stimulate',
    'quizzed': 'quiz', 'drove': 'drive', 'driven': 'drive', 'given': 'give',
    'taken': 'take', 'seen': 'see', 'been': 'be', 'began': 'begin', 'begun': 'begin',
    'made': 'make', 'did': 'do', 'done': 'do', 'said': 'say', 'got': 'get',
    'took': 'take', 'came': 'come', 'went': 'go', 'saw': 'see', 'thought': 'think',
    'brought': 'bring', 'known': 'know', 'shown': 'show', 'written': 'write',
    'broken': 'break', 'chose': 'choose', 'chosen': 'choose', 'frozen': 'freeze',
    'rose': 'rise', 'risen': 'rise', 'spoke': 'speak', 'stole': 'steal', 'threw': 'throw',
    'drew': 'draw', 'flew': 'fly', 'grew': 'grow', 'forgot': 'forget', 'realized': 'realize',
    'organized': 'organize', 'modernized': 'modernize', 'recognized': 'recognize',
    'specialized': 'specialize', 'gives': 'give', 'makes': 'make', 'takes': 'take',
    'uses': 'use', 'used': 'use', 'needed': 'need', 'provided': 'provide', 'felt': 'feel',
    'held': 'hold', 'told': 'tell', 'kept': 'keep', 'left': 'leave', 'meant': 'mean',
    'paid': 'pay', 'put': 'put', 'set': 'set', 'met': 'meet', 'led': 'lead',
    'read': 'read', 'cut': 'cut', 'hit': 'hit', 'spent': 'spend', 'built': 'build',
    'sent': 'send', 'lost': 'lose', 'found': 'find', 'brought': 'bring', 'taught': 'teach',
    'caught': 'catch', 'bought': 'buy', 'thought': 'think', 'fought': 'fight',
    'sought': 'seek', 'wrought': 'work',
}
_VERB_BASE.update({
    'weathered': 'weather', 'caters': 'cater', 'unfolds': 'unfold',
    'directs': 'direct', 'provides': 'provide', 'casts': 'cast',
    'looking': 'look', 'shapes': 'shape', 'tends': 'tend',
    'is': 'be', 'was': 'be', 'were': 'be', 'are': 'be', 'am': 'be',
    'muscled': 'muscle', 'turned': 'turn', 'committed': 'commit',
    'preyed': 'prey', 'invited': 'invite', 'secured': 'secure',
    'referred': 'refer', 'becomes': 'become', 'raises': 'raise',
    'brings': 'bring', 'forced': 'force', 'woke': 'wake',
    'posting': 'post', 'posted': 'post', 'growing': 'grow',
    'clapped': 'clap', 'looked': 'look', 'drew': 'draw', 'took': 'take',
})
# 常见「被动分词 + by/as/with…」搭配，原文常省 be，需补回 be
_PART_NEEDS_BE = {
    'characterized', 'hindered', 'driven', 'followed', 'accompanied', 'based', 'known',
    'located', 'interested', 'satisfied', 'concerned', 'involved', 'married', 'lost',
    'founded', 'built', 'made', 'done', 'given', 'taken', 'seen', 'written', 'broken',
    'considered', 'regarded', 'viewed', 'treated', 'used', 'applied', 'described',
    'defined', 'classified', 'divided', 'shared', 'connected', 'combined', 'mixed',
    'fueled', 'shaped', 'marked', 'noted', 'rooted', 'steeped', 'tended',
}

def _is_full_sentence(s):
    """粗体内容是否为一句完整句（用于跳过基本形式还原，避免破坏句子结构）。
    仅当含句末标点时判定为整句——短语（即使较长）一律还原，不误伤。"""
    t = s.strip()
    if not t:
        return False
    return t.endswith(('.', '?', '!'))


def normalize_collocation(s):
    """将粗体词块（非整句）还原为词组基本形式 —— 通用原则，作用于所有来源。

    规则：
    1. be 动词（is/are/was/were/am）任何位置 → be。
    2. have/has/had 作助动词（后接 been/being/过去分词）→ 还原为后继动词原形（或 be）；
       作实义动词（如 have an ability）保留，has/had 还原为 have。
    3. _VERB_BASE 命中的动词表面形式（三单/过去/分词）→ 原形，任何位置。
    4. 开头被动分词（_PART_NEEDS_BE）→ 补 be。
    5. 绝不误伤名词/形容词/短语；不改变首词大小写（除非该词本身被还原为动词原形）。
    Agent 仍应逐条复核，对映射表未覆盖的动词表面形式手动改成原形。
    """
    if _is_full_sentence(s):
        return s
    s = strip_paren_glosses(s)
    toks = s.split()
    if not toks:
        return s
    low = [t.lower().strip('.,;:') for t in toks]
    out = list(toks)
    changed = False
    # 1) be 动词任何位置 → be
    for k, t in enumerate(low):
        if t in ('is', 'are', 'was', 'were', 'am'):
            out[k] = 'be'
            changed = True
    # 2) have/has/had 助动词还原
    i = 0
    while i < len(low):
        t = low[i]
        if t in ('has', 'have', 'had') and i + 1 < len(low):
            nxt = low[i + 1]
            if nxt in ('been', 'being'):
                out[i] = ''
                out[i + 1] = 'be'
                changed = True
                i += 2
                continue
            if nxt in _VERB_BASE:
                out[i] = ''
                out[i + 1] = _VERB_BASE[nxt]
                changed = True
                i += 2
                continue
            if t in ('has', 'had'):  # 实义动词三单/过去 → have
                out[i] = 'have'
                changed = True
        i += 1
    # 3) 已知动词表面形式（三单/过去/分词）→ 原形，任何位置
    for k, t in enumerate(low):
        if t in _VERB_BASE and out[k].lower().strip('.,;:') == t:
            out[k] = _VERB_BASE[t]
            changed = True
    # 4) 开头被动分词 → 补 be（若该词已在 _VERB_BASE 中还原为原形，则不重复加 be）
    if low and low[0] in _PART_NEEDS_BE and low[0] not in _VERB_BASE:
        out = ['be'] + out
        changed = True
    # 清理去掉助动词产生的空串
    out = [w for w in out if w != '']
    # 仅当首词本身是被还原的动词原形时保持小写（不强行小写专有名词/冠词）
    if out and changed and out[0][:1].isupper() and low[0] in _VERB_BASE:
        out[0] = out[0][0].lower() + out[0][1:]
    return ' '.join(out)

# ── 词性检测 ─────────────────────────────────────────────────────────────────
_pos_cache = {}
_pos_session = None

def _get_pos_session():
    global _pos_session
    if _pos_session is None:
        _pos_session = requests.Session()
    return _pos_session

def is_single_word(text):
    """判断是否为单个英文单词（不包含空格、不含中文）"""
    t = text.strip()
    return bool(re.match(r'^[a-zA-Z\-]+$', t)) and len(t) >= 2

# 词性缩写映射
POS_MAP = {
    'noun': 'n.', 'verb': 'v.', 'adjective': 'a.', 'adverb': 'adv.',
    'preposition': 'prep.', 'conjunction': 'conj.', 'pronoun': 'pron.',
    'interjection': 'interj.', 'determiner': 'det.', 'numeral': 'num.',
    'particle': 'part.', 'auxiliary': 'aux.', 'modal': 'modal',
}

# 后缀规则（优先于API，速度快）
# 注意：许多后缀跨词性（如 -ent/-ant 可名可形，-ing/-ed 可形可动），
# 优先选择规则安全的后缀，模棱两可的不放入
_SUFFIX_POS = [
    # 副词（高置信度）
    (('ly',), 'adv.'),
    # 名词后缀（高置信度）
    (('tion', 'sion', 'ment', 'ness', 'ity', 'ance', 'ence', 'ism', 'ship', 'dom'), 'n.'),
    # 形容词后缀（高置信度）
    (('able', 'ible', 'ful', 'less', 'ous', 'ious', 'ish'), 'a.'),
    # 动词后缀（高置信度）
    (('ize', 'ise', 'ify'), 'v.'),
    # 名词（施动者后缀 -er/-or/-ist，但排除比较级陷阱）
    # -or 仅匹配明确施动者：词长>4 且不以 i 结尾（排除 superior/senior）
    (('er', 'ist'), 'n.'),
]

# 常见词直接映射（人工校准的主词性，避免歧义/网络误判）
_COMMON_POS = {
    'the': 'det.', 'a': 'det.', 'an': 'det.',
    'and': 'conj.', 'but': 'conj.', 'or': 'conj.', 'so': 'conj.', 'yet': 'conj.',
    'of': 'prep.', 'in': 'prep.', 'to': 'prep.', 'for': 'prep.', 'with': 'prep.',
    'on': 'prep.', 'at': 'prep.', 'by': 'prep.', 'from': 'prep.', 'as': 'prep.',
    'is': 'v.', 'are': 'v.', 'was': 'v.', 'were': 'v.', 'be': 'v.', 'been': 'v.',
    'have': 'v.', 'has': 'v.', 'had': 'v.', 'do': 'v.', 'does': 'v.', 'did': 'v.',
    'not': 'adv.', 'very': 'adv.', 'too': 'adv.', 'also': 'adv.', 'just': 'adv.',
    'he': 'pron.', 'she': 'pron.', 'it': 'pron.', 'they': 'pron.', 'we': 'pron.',
    'i': 'pron.', 'you': 'pron.', 'me': 'pron.', 'him': 'pron.', 'her': 'pron.',
    'this': 'pron.', 'that': 'pron.', 'these': 'pron.', 'those': 'pron.',
    # 常见实词（避免后缀规则误判）
    'reveal': 'v.', 'alter': 'v.', 'monitor': 'v.', 'profound': 'a.',
    'humble': 'a.', 'loose': 'a.', 'invisible': 'a.',
    'surpass': 'v.', 'demonstrate': 'v.', 'navigate': 'v.',
    'superior': 'a.', 'prior': 'a.', 'major': 'a.', 'minor': 'a.',
    'particular': 'a.', 'available': 'a.', 'possible': 'a.',
    'approach': 'v.', 'concern': 'v.', 'support': 'v.',
    # 阅读/文档高频词（主词性校准）
    'serve': 'v.', 'affiliate': 'n.', 'giveaway': 'n.', 'purchase': 'n.',
    'postpone': 'v.', 'reschedule': 'v.', 'postponed': 'v.', 'rescheduled': 'v.',
    'provide': 'v.', 'require': 'v.', 'achieve': 'v.', 'create': 'v.',
    'increase': 'v.', 'decrease': 'v.', 'improve': 'v.', 'reduce': 'v.',
    'prevent': 'v.', 'protect': 'v.', 'support': 'v.', 'encourage': 'v.',
    'influence': 'v.', 'determine': 'v.', 'establish': 'v.', 'explain': 'v.',
    'describe': 'v.', 'demonstrate': 'v.', 'represent': 'v.', 'focus': 'v.',
    'emphasize': 'v.', 'highlight': 'v.', 'include': 'v.', 'involve': 'v.',
    'contain': 'v.', 'contribute': 'v.', 'develop': 'v.', 'affect': 'v.',
    'reflect': 'v.', 'suggest': 'v.', 'indicate': 'v.', 'reveal': 'v.',
    'conclude': 'v.', 'analyze': 'v.', 'compare': 'v.', 'consider': 'v.',
    'believe': 'v.', 'expect': 'v.', 'understand': 'v.', 'recognize': 'v.',
    'realize': 'v.', 'maintain': 'v.', 'allow': 'v.', 'require': 'v.',
    'benefit': 'n.', 'challenge': 'n.', 'opportunity': 'n.', 'potential': 'n.',
    'factor': 'n.', 'process': 'n.', 'system': 'n.', 'resource': 'n.',
    'evidence': 'n.', 'theory': 'n.', 'research': 'n.', 'study': 'n.',
    'method': 'n.', 'result': 'n.', 'effect': 'n.', 'cause': 'n.',
    'problem': 'n.', 'solution': 'n.', 'value': 'n.', 'goal': 'n.',
    'knowledge': 'n.', 'skill': 'n.', 'ability': 'n.', 'behavior': 'n.',
    'attitude': 'n.', 'experience': 'n.', 'culture': 'n.', 'development': 'n.',
    'technology': 'n.', 'environment': 'n.', 'society': 'n.', 'world': 'n.',
    'life': 'n.', 'significant': 'a.', 'beneficial': 'a.', 'crucial': 'a.',
    'promising': 'a.', 'global': 'a.', 'modern': 'a.', 'traditional': 'a.',
    'economic': 'a.', 'political': 'a.', 'social': 'a.', 'personal': 'a.',
    'physical': 'a.', 'mental': 'a.', 'emotional': 'a.', 'natural': 'a.',
    'human': 'a.', 'however': 'adv.', 'therefore': 'adv.', 'moreover': 'adv.',
    'furthermore': 'adv.', 'nevertheless': 'adv.',     'although': 'conj.',
    'despite': 'prep.', 'according': 'prep.',
}

# ── 单词类条目词性主词典（权威，离线）─────────────────────────────────────
# 键：normalize 后的基本形式（小写）；值：已嵌入词性标注的完整中文释义。
# 一词多性：按 [词性] 释义 顺序全部列出（如 prompt 既是 v. 也是 adj.）。
# 短语/搭配不在此词典（不标词性）；未命中且非短语时由 get_pos 兜底前缀标注。
_POS = {
    'groundbreaking': '[adj.] 开创性的；突破性的',
    'bottomless': '[adj.] 无底的；取之不尽的',
    'laugh-out-loud': '[adj.] 令人捧腹大笑的',
    'prompt': '[v.] 促使；导致 [adj.] 迅速的；及时的 [n.] 提示；提词',
    'symmetrical': '[adj.] 对称的；匀称的',
    'solitude': '[n.] 独处；孤独',
    'problematic': '[adj.] 成问题的；有困难的',
    'dilute': '[v.] 稀释；冲淡 [adj.] 稀释的；经稀释的',
    'applicable': '[adj.] 适用的；适当的',
    # —— Z20一联 套卷 生词 ——
    'timepiece': '[n.] 计时装置；钟表',
    'newly-launched': '[adj.] 新推出的，刚发布的',
    'stereotype': '[n.] 刻板印象，老一套',
    'retrieve': '[v.] 取回，拿回，找回',
    'cue': '[n.] 暗示，提示，线索',
    'sycophancy': '[n.] 谄媚，阿谀奉承',
    'sycophantic': '[adj.] 谄媚的，阿谀奉承的',
    'prioritize': '[v.] 优先考虑，按优先顺序处理',
    'tolerance': '[n.] 容忍，宽容，忍受',
    'would-be': '[adj.] 想要成为…的，潜在的，未来的',
    'ill-intentioned': '[adj.] 恶意的，居心不良的',
    'unimpressive': '[adj.] 给人印象不深的，平平无奇的',
    'weary': '[adj.] 疲惫的，疲倦的',
    'sprout': '[v.] （种子）发芽，抽芽',
    'atop': '[prep.] 在……顶上，在……上方',
    'wrestler': '[n.] 摔跤运动员，摔跤手',
    'herder': '[n.] 牧民，放牧者',
    'orchestra': '[n.] （大型）管弦乐队',
    'creator': '[n.] 创作者，创造者',
    'influencer': '[n.] 网络红人，影响者',
}


def apply_pos(en, cn):
    """单词类条目：在中文释义前标注词性（[n.]/[v.]/[adj.]/[adv.]），一词多性全部列出。

    短语/搭配（含空格）不标词性，原样返回。优先采用离线权威 _POS；
    未命中且确为单词时，用 get_pos 兜底前缀标注（不保证多性完整）。
    """
    if not cn or cn == '【待译】':
        return cn
    base = normalize_collocation(en)
    if not is_single_word(base):
        return cn  # 短语/搭配不标词性
    key = base.lower()
    if key in _POS:
        return _POS[key]  # 权威：可能含多词性 + 释义
    p = get_pos(base)
    if p:
        return f'[{p}] {cn}'
    return cn


def get_pos(word):
    """获取英文单词的词性缩写（仅本地规则，不依赖网络，保证稳定且准确）"""
    w = word.strip().lower()
    if w in _pos_cache:
        return _pos_cache[w]

    # 1. 常见词映射（人工校准主词性）
    if w in _COMMON_POS:
        result = _COMMON_POS[w]
        _pos_cache[w] = result
        return result

    # 2. 后缀规则（高置信度后缀）
    base = w.rstrip('s')  # 去复数
    for suffixes, pos in _SUFFIX_POS:
        for suf in suffixes:
            if base.endswith(suf) and len(base) > len(suf) + 1:
                _pos_cache[w] = pos
                return pos

    # 2b. -or 特殊处理：仅匹配明确施动者（排除 superior/senior/junior/major 等）
    if base.endswith('or') and len(base) > 4 and not base.endswith('ior'):
        _pos_cache[w] = 'n.'
        return 'n.'

    # 2c. -ive/-al/-ic/-ary 作为形容词（相对安全但非绝对）
    if base.endswith(('ive', 'al', 'ic', 'ary')) and len(base) > 4:
        _pos_cache[w] = 'a.'
        return 'a.'

    # 3. 特殊规则：-ing/-ed → 动词
    if w.endswith(('ing', 'ed')) and len(w) > 5:
        _pos_cache[w] = 'v.'
        return 'v.'

    _pos_cache[w] = None
    return None

# ── 文本话题 / 体裁 自动概括（关键词启发式，离线可用，无需联网）────────────────
_TOPIC_KEYWORDS = {
    '校园生活': ['student', 'students', 'school', 'teacher', 'teachers', 'class', 'classroom',
               'exam', 'exams', 'learn', 'learning', 'lesson', 'lessons', 'course', 'courses',
               'education', 'educational', 'study', 'studying', 'campus', 'homework', 'subject',
               'subjects', 'professor', 'lecture', 'lectures', 'graduation', 'college', 'university'],
    '人物与故事': ['story', 'stories', 'life', 'experience', 'experiences', 'family', 'friend',
                'friends', 'childhood', 'memory', 'memories', 'character', 'hero', 'grandfather',
                'grandmother', 'mother', 'father', 'parent', 'parents', 'daughter', 'son',
                'neighbor', 'neighbour', 'writer', 'artist', 'scientist', 'biography'],
    '科技与网络': ['internet', 'online', 'computer', 'computers', 'app', 'apps', 'digital',
                'technology', 'technological', 'robot', 'robots', 'smartphone', 'software',
                'website', 'websites', 'social media', 'network', 'networks', 'ai',
                'artificial intelligence', 'data', 'device', 'devices', 'virtual', 'algorithm',
                'research', 'researchers', 'researcher', 'experiment', 'experiments',
                'microscope', 'invent', 'invention', 'invented', 'scientist', 'scientific',
                'cognitive', 'psychology', 'findings', 'innovation', 'innovate'],
    '环境与自然': ['environment', 'environmental', 'protect', 'protection', 'pollution', 'climate',
                'recycle', 'recycling', 'nature', 'planet', 'ocean', 'wildlife', 'animal', 'animals',
                'forest', 'forests', 'earth', 'green', 'sustainable', 'sustainability', 'species',
                'ecology', 'wild', 'global warming', 'carbon'],
    '健康与生活': ['health', 'healthy', 'exercise', 'diet', 'sleep', 'disease', 'mental', 'hospital',
                'doctor', 'fitness', 'stress', 'body', 'food', 'nutrition', 'medical', 'medicine',
                'wellbeing', 'well-being', 'habit', 'habits', 'smoking', 'weight',
                'self-care', 'selfcare', 'burnout', 'balance', 'boundaries', 'boundary',
                'rest', 'relax', 'relaxation', 'anxiety', 'care', 'caring'],
    '文化与传统': ['culture', 'cultural', 'tradition', 'traditional', 'festival', 'festivals',
                'custom', 'customs', 'art', 'music', 'history', 'historical', 'heritage',
                'celebrate', 'celebration', 'language', 'religion', 'literature', 'painting'],
    '社会与现象': ['society', 'social', 'community', 'communities', 'government', 'economy',
                'economic', 'problem', 'problems', 'change', 'global', 'world', 'people',
                'population', 'policy', 'policies', 'employment', 'job', 'jobs', 'poverty'],
    '体育与活动': ['sport', 'sports', 'game', 'games', 'match', 'team', 'player', 'players',
                'stadium', 'race', 'competition', 'competitions', 'olympic', 'olympics',
                'athlete', 'athletes', 'championship', 'tournament', 'score', 'win', 'winning'],
}

_GENRE_APPLIED = ['dear ', 'yours ', 'sincerely', 'best regards', 'kind regards', 'notice',
                  'application', 'email', ' from:', ' to:', 'dear sir', 'dear madam',
                  'yours faithfully', 'yours truly',
                  'ticket', 'tickets', 'promotion', 'promotions', 'discount',
                  'admission', 'reserve', 'subscribe', 'membership', 'festival',
                  'fare', 'fares', 'venue', 'venues', 'poster', 'posters', 'free entry',
                  'booking', 'book your', 'buy tickets', 'on sale']
_GENRE_NEWS = ['reported that', 'it was reported', 'spokesman', 'announced that',
               'the report said', 'survey showed', 'a survey found', 'news',
               'reporters', 'according to official', 'press release']
_GENRE_ARGUE = ['in my opinion', 'i think', 'we should', 'should be', 'i believe',
                'it is important to', 'argue that', 'in conclusion', 'from my point',
                'as far as i', 'we must', 'it is necessary']
_GENRE_NARRATE = ['one day', 'when i was', 'we were', 'my childhood', 'last year',
                  'yesterday', 'grew up', 'i remember', 'a few years ago', 'when i was young',
                  'my mother', 'my father', 'my friend', 'my family', 'childhood',
                  'one evening', 'i began', 'i realized', 'i learned', 'i decided',
                  'i collapsed', 'i felt', 'staring at', 'that day', 'at that moment',
                  'i started', 'i knew', 'i thought']


def _classify_topic(text):
    """根据英文文本关键词概括话题类别（如 校园生活 / 科技与网络）。"""
    if not text or not text.strip():
        return '生活与学习'
    low = text.lower()
    best, best_score = '生活与学习', 0
    for cat, kws in _TOPIC_KEYWORDS.items():
        s = 0
        for kw in kws:
            s += len(_re.findall(r'\b' + _re.escape(kw) + r'\b', low))
        if s > best_score:
            best, best_score = cat, s
    return best if best_score > 0 else '生活与学习'


def _classify_genre(text):
    """根据英文文本特征判断体裁（应用文/新闻报道/议论文/记叙文/说明文）。"""
    if not text or not text.strip():
        return '说明文'
    low = ' ' + text.lower() + ' '
    if any(k in low for k in _GENRE_APPLIED):
        return '应用文'
    if any(k in low for k in _GENRE_NEWS):
        return '新闻报道'
    if any(k in low for k in _GENRE_ARGUE):
        return '议论文'
    if any(k in low for k in _GENRE_NARRATE):
        return '记叙文'
    return '说明文'


def classify_section_meta(text, section=None):
    """返回 {'topic': 话题类别, 'genre': 体裁}。"""
    genre = _classify_genre(text)
    if section == '完形填空':
        genre = '记叙文'
    return {'topic': _classify_topic(text), 'genre': genre}


def _read_section_texts(input_bytes):
    """读取原始文档，按篇目聚合每段的纯文本（用于话题/体裁判断）。"""
    doc = Document(io.BytesIO(input_bytes))
    texts = {}
    current = '正文'
    current_part = None
    for para in _iter_paragraphs(doc):
        runs = list(para.runs)
        t = para.text.strip()
        sec, part = _detect_section_change(t, len(runs))
        if part == 'LANG':
            current_part = 'LANG'
            current = '完形填空'
            continue
        if part == 'READ':
            current_part = 'READ'
            continue
        if sec == 'SUBSECOND':
            current = '语法填空' if current_part == 'LANG' else '七选五'
            continue
        if sec:
            current = sec
            continue
        if not t:
            continue
        texts.setdefault(current, []).append(t)
    return {k: ' '.join(v) for k, v in texts.items()}


def _build_section_meta(sections, input_bytes=None):
    """为每篇目计算话题/体裁。有原文时用全文，否则退化为词块英文。"""
    texts = _read_section_texts(input_bytes) if input_bytes else {}
    meta = {}
    for s, its in sections.items():
        if not its:
            continue
        text = texts.get(s, '')
        if not (text or '').strip():
            text = ' '.join(en for en, _, _ in its)
        meta[s] = classify_section_meta(text, s)
    return meta


# ── 选项行识别（过滤选择题答案高亮，避免污染语言点）─────────────────────────
def _is_option_line(text):
    """跳过选择题选项行：A./B./C./D.… 及七选五 A–G 选项。
    这些粗体是「正确答案高亮」，不是要背的语言点。"""
    t = (text or '').strip()
    if not t:
        return False
    # 以单个 A–G / a–g 字母 + '.' 或 '、' 开头（如 "A. ..." "B. ..." "G. ..."）
    if re.match(r'^[A-Ga-g][\.、]', t):
        return True
    # 一题多选项挤在一行：A. … B. … C./D.
    if ('A.' in t and 'B.' in t and ('C.' in t or 'D.' in t)):
        return True
    return False


# ── 翻译收集（供各输出模式复用）──────────────────────────────────────────────
def _collect_translations(input_bytes):
    """
    从 .docx 中提取所有粗体英文词块并翻译。
    返回: list of (en, cn, section, pos)  其中 pos 仅对单单词有效
    """
    doc = Document(io.BytesIO(input_bytes))
    translation_list = []
    current_section = '正文'
    current_part = None

    for para in _iter_paragraphs(doc):
        runs = list(para.runs)
        para_text = para.text.strip()

        # 试卷结构识别：大题分隔 / 篇目头 / 小题头 / 题型指令句
        sec, part = _detect_section_change(para_text, len(runs))
        if part == 'LANG':
            current_part = 'LANG'
            current_section = '完形填空'   # 语言运用第一小节默认完形填空
            continue
        if part == 'READ':
            current_part = 'READ'
            continue
        if sec == 'SUBSECOND':
            current_section = '语法填空' if current_part == 'LANG' else '七选五'
            continue
        if sec:
            current_section = sec
            continue

        # 跳过选择题选项行（粗体为答案高亮，非语言点），避免污染背诵/默写版
        if _is_option_line(para_text):
            continue

        if not runs:
            continue

        i = 0
        while i < len(runs):
            r = runs[i]
            is_bold = r.bold in (True, 'True', 1, '1')
            txt = r.text or ''

            if is_bold and txt.strip():
                # 合并连续「英文」粗体 run（构成一个英文词块 / 短语）
                en_parts = [txt]
                j = i + 1
                while (j < len(runs)
                       and runs[j].bold in (True, 'True', 1, '1')
                       and not _contains_cjk(runs[j].text or '')):
                    en_parts.append(runs[j].text or '')
                    j += 1
                en = ''.join(en_parts).strip()
                en = en.replace('\xa0', ' ')  # 归一化不间断空格，避免与词典键不匹配
                # 紧随其后的连续「中文」粗体 run 视为该词块的中文释义
                cn_gloss = []
                while (j < len(runs)
                       and runs[j].bold in (True, 'True', 1, '1')
                       and _contains_cjk(runs[j].text or '')):
                    cn_gloss.append((runs[j].text or '').strip())
                    j += 1
                cn_hint = ''.join(cn_gloss) if cn_gloss else None

                # 跳过全大写多词英文标题（篇章强调，非词汇点）
                if en and en.isupper() and ' ' in en and not cn_hint:
                    i = j
                    continue
                # 跳过纯标点 / 无字母无汉字的噪声（如 "?"、"."）
                if not re.search(r'[A-Za-z]', en) and not _contains_cjk(en):
                    i = j
                    continue

                if cn_hint:
                    # 已自带中文释义（bilingual 语言点源）
                    cn = cn_hint
                    pos = get_pos(en) if is_single_word(en) else None
                    translation_list.append((en, cn, current_section, pos))
                elif _contains_cjk(en):
                    # 纯中文粗体词块：需给出英文
                    cn = en
                    en_trans = translate_zh_en(en)
                    translation_list.append((en_trans or en, cn, current_section, None))
                else:
                    # 纯英文粗体词块：译中文（有道 + MyMemory 兜底）
                    cn = translate(en)
                    pos = get_pos(en) if is_single_word(en) else None
                    translation_list.append((en, cn, current_section, pos))
                i = j
            else:
                i += 1

    return translation_list
# 题目分类关键词（显式篇目名）
SECTION_KEYWORDS = ['七选五', '完形填空', '语法填空', '短文填空',
                    '阅读表达', '任务型阅读', '选词填空', '词汇运用',
                    'A篇', 'B篇', 'C篇', 'D篇', 'E篇', 'F篇']

# 形如 "A篇" / "B" / "C." / "第一篇" / "Passage A" 等的篇目标记
import re as _re
_SECTION_RE = _re.compile(r'^[A-Fa-f][篇．.\s]?$|^第[一二三四五六七八九十]+篇$|^Passage\s*[A-Fa-f]$')

# 结构化试卷头指令句（无单独篇目名时据此判定题型）
# 注意：完形填空与七选五都含"选出可以填入空白处"，须用各自专属短语区分：
#   七选五 —— "从短文后的选项中"（选项列于文后，七选五）
#   完形填空 —— "A、B、C 四个选项" / "掌握其大意"
_STRUCTURAL = [
    ('七选五', ['从短文后的选项中', '短文后的选项', '七选五']),
    ('语法填空', ['在空白处填入', '括号内单词的正确形式', '填入 1 个适当的单词', '填入一个适当的单词']),
    ('完形填空', ['掌握其大意', '完形填空', '完型填空', '四个选项中', 'A、B、C', 'A. B. C. D.']),
]


def _section_label(text):
    """返回干净篇目名；非篇目头则返回 None。可识别结构化试卷头。"""
    t = (text or '').strip()
    if not t:
        return None
    # 显式篇目关键词
    for kw in ('七选五', '完形填空', '语法填空', '短文填空', '阅读表达',
               '任务型阅读', '选词填空', '词汇运用'):
        if kw in t:
            return kw
    for kw in ('A篇', 'B篇', 'C篇', 'D篇', 'E篇', 'F篇'):
        if kw in t:
            return kw[0]
    # 篇目字母标记（A / B / C / D … 或 Passage A）
    if _SECTION_RE.match(t):
        return t[0].upper() if t[0].isalpha() else t
    # 结构化指令句
    for name, sigs in _STRUCTURAL:
        if any(s in t for s in sigs):
            return name
    return None


def _is_part_divider(text):
    """返回 '阅读' / '语言运用' / None（试卷大题分隔）。"""
    t = (text or '').strip()
    if not t:
        return None
    if '语言运用' in t:
        return '语言运用'
    if '阅读' in t and ('第一部分' in t or '共两节' in t or '满分' in t or '阅读理解' in t):
        return '阅读'
    return None


def _detect_section_change(text, runs_len):
    """
    返回 (section_label_or_None, part_or_None)。
    识别高考英语试卷结构：阅读(A-D + 七选五) / 语言运用(完形填空 + 语法填空)。
    'SUBSECOND' 为占位：调用方按当前大题推导为 七选五 或 语法填空。
    """
    t = (text or '').strip()
    if not t:
        return (None, None)
    # 大题分隔（不依赖 run 数，结构性标题）
    if ('语言运用' in t) or ('语言知识运用' in t):
        return (None, 'LANG')
    if ('第一部分' in t or '第二部分' in t or '第三部分' in t) and '阅读' in t:
        return (None, 'READ')
    # 显式篇目名 / 字母标记
    lbl = _section_label(t)
    if lbl:
        return (lbl, None)
    # 小题头（结构性标记，正文不会出现，故不限制 run 数）
    if '第一节' in t:
        return ('完形填空', None)
    if '第二节' in t:
        return ('SUBSECOND', None)
    # 题型指令句
    if ('从短文后的选项中选出' in t) or ('选出可以填入空白处' in t):
        return ('七选五', None)
    if ('在空白处填入' in t) or ('括号内单词的正确形式' in t):
        return ('语法填空', None)
    if '掌握其大意' in t:
        return ('完形填空', None)
    return (None, None)


def _iter_paragraphs(doc):
    """遍历正文段落与表格单元格内的段落（完形/语法常置于表格中）。"""
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def is_section_header(text):
    """判断是否是题目分类标记（如'A篇'、'C篇'、'完形填空'等）"""
    t = (text or '').strip()
    if not t:
        return False
    if _section_label(t):
        return True
    if len(t) <= 10 and any(kw in t for kw in SECTION_KEYWORDS):
        return True
    return False

def process_docx_bytes(input_bytes):
    """
    段落挖空模式：粗体 → 下划线 +（中文释义），其他内容原样保留
    
    输入：原始 .docx 文件内容（bytes）
    输出：处理后的 .docx 文件内容（bytes）
    """
    doc = Document(io.BytesIO(input_bytes))
    translation_list = []  # [(en, cn, section)]
    current_section = '正文'

    for para in doc.paragraphs:
        runs = list(para.runs)
        para_text = para.text.strip()
        
        if is_section_header(para_text) and len(runs) <= 3:
            lbl = _section_label(para_text)
            current_section = lbl if lbl else para_text
            for r in runs:
                if r.text:
                    set_run_font(r, is_chinese=True)
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            continue

        if not runs:
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            continue

        new_run_specs = []
        i = 0
        while i < len(runs):
            r = runs[i]
            is_bold = r.bold in (True, 'True', 1, '1')
            txt = r.text or ''

            if is_bold and txt.strip():
                bold_parts = [txt]
                j = i + 1
                while j < len(runs) and (runs[j].bold in (True, 'True', 1, '1')):
                    bold_parts.append(runs[j].text or '')
                    j += 1
                en = ''.join(bold_parts).strip()
                cn = translate(en)
                translation_list.append((en, cn, current_section))
                n = blank_display_len(en)
                new_run_specs.append(('underline', n, en, cn))
                i = j
            else:
                new_run_specs.append(('plain', txt))
                i += 1

        for r in runs:
            r.text = ''

        for spec in new_run_specs:
            if spec[0] == 'plain':
                txt = spec[1]
                if not txt:
                    continue
                r_new = para.add_run(txt)
                set_run_font(r_new, is_chinese=has_chinese(txt))
            elif spec[0] == 'underline':
                n = spec[1]
                en = spec[2]
                cn = spec[3]
                r_u = para.add_run('_' * n)
                r_u.underline = True
                set_run_font(r_u, is_chinese=False)
                if cn:
                    r_cn = para.add_run(f'（{cn}）')
                    set_run_font(r_cn, is_chinese=True)

        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    # ── 参考答案 ────────────────────────────────────────────────
    doc.add_paragraph()
    h2 = doc.add_paragraph('【参考答案】')
    for run in h2.runs:
        run.font.bold = True
        set_run_font(run, is_chinese=True)

    sections = {}
    for idx, (en, cn, section) in enumerate(translation_list):
        if en:
            sections.setdefault(section, []).append((idx+1, en, cn))

    for section, items in sections.items():
        p_sec = doc.add_paragraph(f'{section}')
        for run in p_sec.runs:
            run.font.bold = True
            set_run_font(run, is_chinese=True)
        
        for num, en, cn in items:
            display_cn = cn if cn is not None else '【待译】'
            p = doc.add_paragraph(f'{num}. {en}  →  {display_cn}')
            for run in p.runs:
                set_run_font(run, is_chinese=has_chinese(run.text))

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _set_row_height(row, h_pt):
    """设置表格行高（半磅 = pt*20），hRule=atLeast 保证最小高度"""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = trPr.makeelement(qn('w:trHeight'),
                                {qn('w:val'): str(int(h_pt * 20)), qn('w:hRule'): 'atLeast'})
    trPr.append(trHeight)


def _block_header_text(section):
    """每篇文本的分类标题（纯篇头，如 A篇 / B篇）"""
    return f'{section}'


def _add_merged_row(table, text, bold=False, is_chinese=True, height=22):
    """在表格末尾追加一行，并将首格合并至整行，写入 text。"""
    row = table.add_row().cells
    row[0].merge(row[-1])
    row[0].text = text
    _format_cell(row[0], bold=bold, is_chinese=is_chinese)
    _set_row_height(table.rows[-1], height)
    return table.rows[-1]


def _derive_base(name):
    """从源文件名推导套卷名（去掉 .docx 及常见后缀），用于文档标题/文件名。"""
    base = re.sub(r'\.docx$', '', name or '', flags=re.I)
    base = re.sub(r'[_\s\-－]?(已处理|表格版|答案版|默写版|背诵版)\s*$', '', base)
    return base


def _build_recite_docx(sections, title='语言点背诵版', section_meta=None):
    """
    背诵版（答案/记忆用）：2 列 —— 左中文(+词性) | 右英文。
    结构参照用户参考文档：标题行(合并) + 每篇[篇头(合并) / 文本话题(合并) / 文本体裁(合并)] + 词块。
    无单独的列头行（与参考文档一致）。
    """
    doc = Document()
    if not sections:
        p = doc.add_paragraph('（文档中未找到任何翻译内容）')
        for run in p.runs:
            set_run_font(run, is_chinese=True)
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    section_items = [(s, its) for s, its in sections.items() if its]

    table = doc.add_table(rows=1, cols=2, style='Table Grid')
    _set_table_width(table, [Cm(8), Cm(8)])

    _add_merged_row(table, title, bold=True, is_chinese=True, height=28)

    first = True
    for section, items in section_items:
        if not first:
            sp = table.add_row().cells
            sp[0].text = ''
            sp[1].text = ''
            _format_cell(sp[0], is_chinese=True)
            _format_cell(sp[1], is_chinese=False)
            _set_row_height(table.rows[-1], 10)
        first = False

        _add_merged_row(table, section, bold=True, is_chinese=True, height=22)
        _sm = ((section_meta or {}).get(section, {}) or {})
        _topic = '文本话题：' + (_sm.get('topic') or '＿＿＿＿＿')
        _genre = '文本体裁：' + (_sm.get('genre') or '＿＿＿＿＿')
        _add_merged_row(table, _topic, bold=False, is_chinese=True, height=20)
        _add_merged_row(table, _genre, bold=False, is_chinese=True, height=20)

        for (en, cn, pos) in items:
            cn_display = cn if cn else '【待译】'
            cn_display = apply_pos(en, cn_display)
            row = table.add_row().cells
            row[0].text = cn_display
            _format_cell(row[0], bold=False, is_chinese=True)
            row[1].text = normalize_collocation(en)
            _format_cell(row[1], bold=False, is_chinese=False)
            _set_row_height(table.rows[-1], 20)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _build_dictate_docx(sections, title='语言点默写版', section_meta=None):
    """
    默写版（练习用）：3 网格列 —— 左2列合并为中文(+词性)宽列，第3列为空白英文默写格。
    结构参照用户参考文档：标题行(合并) + 每篇[篇头+文本话题+文本体裁 合并为一行] + 词块(中文宽列|空白)。
    无单独的列头行（与参考文档一致）。
    """
    doc = Document()
    if not sections:
        p = doc.add_paragraph('（文档中未找到任何翻译内容）')
        for run in p.runs:
            set_run_font(run, is_chinese=True)
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    section_items = [(s, its) for s, its in sections.items() if its]

    table = doc.add_table(rows=1, cols=3, style='Table Grid')
    _set_table_width(table, [Cm(5.3), Cm(5.3), Cm(5.0)])

    _add_merged_row(table, title, bold=True, is_chinese=True, height=28)

    first = True
    for section, items in section_items:
        if not first:
            sp = table.add_row().cells
            sp[0].merge(sp[1])
            sp[0].text = ''
            sp[2].text = ''
            _format_cell(sp[0], is_chinese=True)
            _format_cell(sp[2], is_chinese=False)
            _set_row_height(table.rows[-1], 10)
        first = False

        # 篇头 + 话题 + 体裁 合并为整行（与参考文档一致）
        _sm = ((section_meta or {}).get(section, {}) or {})
        _topic = '文本话题：' + (_sm.get('topic') or '＿＿＿＿＿')
        _genre = '文本体裁：' + (_sm.get('genre') or '＿＿＿＿＿')
        hdr_text = f'{section}\n{_topic}\n{_genre}'
        _add_merged_row(table, hdr_text, bold=True, is_chinese=True, height=44)

        for (en, cn, pos) in items:
            cn_display = cn if cn else '【待译】'
            cn_display = apply_pos(en, cn_display)
            row = table.add_row().cells
            row[0].merge(row[1])
            row[0].text = cn_display
            _format_cell(row[0], bold=False, is_chinese=True)
            row[2].text = ''
            _format_cell(row[2], bold=False, is_chinese=False)
            _set_row_height(table.rows[-1], 20)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def process_docx_bytes_table(input_bytes, base=''):
    """从原始文档（粗体英文）生成默写版（练习）单文档，向后兼容。"""
    sections = _group_translations(input_bytes)
    section_meta = _build_section_meta(sections, input_bytes)
    return _build_dictate_docx(sections, (base + '语言点默写版') if base else '语言点默写版', section_meta)


def process_docx_bytes_two(input_bytes, base='', section_meta=None):
    """
    从原始文档（粗体英文）生成两个独立文档：背诵版 + 默写版。
    返回 (背诵版_bytes, 默写版_bytes)。
    section_meta: 可选覆盖，dict[{section}] = {'topic':.., 'genre':..}。
        传入时直接采用（保证话题/体裁准确）；为 None 时自动按关键词概括。
    """
    sections = _group_translations(input_bytes)
    if section_meta is None:
        section_meta = _build_section_meta(sections, input_bytes)
    recite = _build_recite_docx(sections, (base + '语言点背诵版') if base else '语言点背诵版', section_meta)
    dictate = _build_dictate_docx(sections, (base + '语言点默写版') if base else '语言点默写版', section_meta)
    return recite, dictate


def _group_translations(input_bytes):
    """从原始文档提取翻译并分组为 sections dict"""
    translation_list = _collect_translations(input_bytes)
    sections = {}
    for en, cn, section, pos in translation_list:
        if en and section != '正文':
            sections.setdefault(section, []).append((en, cn, pos))
    return sections


def convert_paragraph_to_table(input_bytes, base=''):
    """段落→表格模式：从已处理段落挖空版解析，生成默写版（练习）单文档，向后兼容。"""
    sections = _parse_answer_key(input_bytes)
    section_meta = _build_section_meta(sections)
    return _build_dictate_docx(sections, (base + '语言点默写版') if base else '语言点默写版', section_meta)


def convert_paragraph_to_table_two(input_bytes, base=''):
    """段落→表格模式：生成两个独立文档（背诵版 + 默写版）。返回 (背诵版, 默写版)。"""
    sections = _parse_answer_key(input_bytes)
    section_meta = _build_section_meta(sections)
    recite = _build_recite_docx(sections, (base + '语言点背诵版') if base else '语言点背诵版', section_meta)
    dictate = _build_dictate_docx(sections, (base + '语言点默写版') if base else '语言点默写版', section_meta)
    return recite, dictate


def _parse_answer_key(input_bytes):
    """
    从段落挖空版文档中解析「参考答案」部分，提取 (en, cn, section) 数据。
    """
    doc = Document(io.BytesIO(input_bytes))
    translation_list = []
    in_answer = False
    current_section = '正文'

    for para in doc.paragraphs:
        t = para.text.strip()
        if not t:
            continue

        # 检测「参考答案」标记
        if '参考答案' in t and any(r.bold for r in para.runs if r.text.strip()):
            in_answer = True
            continue

        if not in_answer:
            continue

        # 检测题目分类（A篇, B篇等）——粗体短文本
        is_section_header = (
            any(r.bold for r in para.runs if r.text.strip()) and
            len(t) <= 5 and '篇' in t
        )
        if is_section_header:
            current_section = t
            continue

        # 解析翻译行: "1. route  →  路线；途径；方式"
        m = re.match(r'^\d+\.\s*(.+?)\s*→\s*(.+)$', t)
        if m:
            en = m.group(1).strip()
            cn = m.group(2).strip()
            pos = get_pos(en) if is_single_word(en) else None
            translation_list.append((en, cn, current_section, pos))

    # 分组
    sections = {}
    for en, cn, section, pos in translation_list:
        if en:
            sections.setdefault(section, []).append((en, cn, pos))

    return sections


# ── 表格辅助函数 ────────────────────────────────────────────────────────────
def _set_table_width(table, widths):
    """设置表格列宽"""
    for row in table.rows:
        for idx, width in enumerate(widths):
            if idx < len(row.cells):
                row.cells[idx].width = width


def _format_cell(cell, bold=False, is_chinese=False):
    """设置单元格内所有段落的字体"""
    for para in cell.paragraphs:
        para.paragraph_format.space_before = Pt(2)
        para.paragraph_format.space_after = Pt(2)
        for run in para.runs:
            run.font.bold = bold
            set_run_font(run, is_chinese=is_chinese)


def text_to_docx_bytes(text):
    """纯文本 → .docx（**text** 标记粗体）"""
    doc = Document()
    translation_list = []

    for line in text.split('\n'):
        if '**' in line:
            para = doc.add_paragraph()
            parts = re.split(r'(\*\*.*?\*\*)', line)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    en = part[2:-2].strip()
                    cn = translate(en)
                    translation_list.append((en, cn, '正文'))
                    n = blank_display_len(en)
                    r1 = para.add_run('_' * n)
                    r1.underline = True
                    set_run_font(r1, is_chinese=False)
                    if cn:
                        r2 = para.add_run(f'（{cn}）')
                        set_run_font(r2, is_chinese=True)
                elif part:
                    r = para.add_run(part)
                    set_run_font(r, is_chinese=has_chinese(part))
        else:
            para = doc.add_paragraph(line)
            for run in para.runs:
                set_run_font(run, is_chinese=has_chinese(line))
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    # 参考答案
    doc.add_paragraph()
    h2 = doc.add_paragraph('【参考答案】')
    for run in h2.runs:
        run.font.bold = True
        set_run_font(run, is_chinese=True)

        for idx, (en, cn, section) in enumerate(translation_list):
            if en:
                display_cn = cn if cn is not None else '【待译】'
                p = doc.add_paragraph(f'{idx+1}. {en}  →  {display_cn}')
            for run in p.runs:
                set_run_font(run, is_chinese=has_chinese(run.text))

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── CLI 测试 ────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    import sys
    src = sys.argv[1] if len(sys.argv) > 1 else 'input.docx'
    dst = sys.argv[2] if len(sys.argv) > 2 else src.replace('.docx', '_已处理.docx')
    with open(src, 'rb') as f:
        out = process_docx_bytes(f.read())
    with open(dst, 'wb') as f:
        f.write(out)
    print(f'✅ 已生成：{dst}')
    print(f'   共 {len([x for x in [1] if x])} 个翻译条目')  # 简化输出
