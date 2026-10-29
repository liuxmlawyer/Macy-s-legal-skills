#!/usr/bin/env python3
"""
labor-compliance-assistant · 用工合同关键条款抽取

用途
----
从劳动合同、竞业协议、规章制度、解除通知书等文档中抽取
劳动用工相关的关键条款，并标注风险线索，供后续法律评估使用。

【重要边界】
本脚本是**可选增强**，不是核心依赖：
  · 抽取结果是**线索**，不是法律结论，必须回A1 权威源核验法条
  · 抽取失败时降级为直接读原文，不阻塞主流程
  · 纯标准库实现，无需安装依赖（仅 .docx 读取需 python-docx）

支持的格式
----------
.md / .txt / .docx（需python-docx，未安装时降级提示）

用法
----
python extract_labor_terms.py <文件路径>
python extract_labor_terms.py <文件路径> --json
python extract_labor_terms.py <文件路径> --type probation,noncompete
python extract_labor_terms.py <文件路径> --min-risk high

退出码
------
0  正常（含"未命中已知条款类型"，JSON 模式仍输出约定结构）
1  文件不存在 / 文档损坏 / 读取失败（受控错误，无裸堆栈）
2  参数错误

输出契约（--json）
-----------------
无论是否命中，stdout 均为一个 JSON 对象，字段固定：
  meta      {file, path, format, extracted_at, types_scanned}
  status    "ok" | "no_match" | "error"
  clauses   [{type, label, items:[{section, text, risks:[{trigger,note,level}]}]}]
  error     null | {code, message, degraded, hint}
  disclaimer  抽取线索非法律结论，须回权威源核验法条后定性
"""

import argparse
import io
import json
import os
import re
import sys
from datetime import datetime

# --------------------------------------------------------------------------
# 条款类型定义：劳动用工语义（区别于通用合同条款）
#
# 设计约定
#  1. patterns 一律带"类型关键词"约束，避免宽泛时间表述被跨类型吞并
#     （例："竞业限制期限为五年"不得被归入"试用期"）。
#  2. risk_high / risk_note 按插入顺序匹配；risk_note 命中项即为该风险
#     附带的法条提示与适用条件。
#  3. 提示文本中凡涉及适用条件的，必须区分"在职 / 离职后""固定期限 /
#     无固定期限"等分支，不得输出与现行法相反的绝对表述。
#  4. 全部提示对应的法条已登记于 references/script-clause-registry.md，
#     引用前仍须回 A1 权威源核验。
# --------------------------------------------------------------------------
CLAUSE_TYPES = {
    "probation": {
        "label": "试用期",
        "section_headers": ["试用期", "试用期限", "试用期待遇"],
        "patterns": [
            r"试用期[^。；\n]{0,40}(?:期限|时间|不超过|最长)",
            r"试用期为[^。；\n]{0,30}",
            r"(?:约定|设置|不得设置)[^。；\n]{0,10}试用期",
            r"试用期[^。；\n]{0,15}(?:为|是|:|：)?\s*[0-9一二三四五六七八九十百两]+\s*(?:个)?\s*(?:月|年|日|天)",
            r"[0-9一二三四五六七八九十百两]+\s*(?:个)?\s*(?:月|年)试用期",
        ],
        # 试用期风险由"试用期长度 + 合同期限"联合判断，见 assess_probation()
        "risk_high": [],
        "risk_note": {},
    },
    "noncompete": {
        "label": "竞业限制",
        "section_headers": ["竞业限制", "竞业禁止", "保密义务"],
        "patterns": [
            r"竞业(?:限制|禁止)[^。；\n]{0,80}",
            r"离职后[^。；\n]{0,40}(?:不得|禁止)[^。；\n]{0,40}(?:任职|任职于|从事)",
            r"保密(?:义务|期限)[^。；\n]{0,50}",
        ],
        "risk_high": [
            # 竞业限制期限超过二年（24个月）——用数值比较，避免只匹配"三年/五年"而漏掉
            # "30个月""两年零一个月"等表述。

            # 覆盖：不另行支付 / 不予支付 / 无需支付 / 未约定 / 没有约定 补偿
            r"(?:不[^。；\n]{0,4}(?:支付|发放|另付)|无需[^。；\n]{0,4}支付|"
            r"没有[^。；\n]{0,6}(?:约定|支付)|未[^。；\n]{0,4}约定)[^。；\n]{0,20}"
            r"(?:经济)?(?:补偿|报酬|费用|津贴)",
            r"(?:违约金|赔偿金)[^。；\n]{0,25}(?:竞业|限制)[^。；\n]{0,25}(?:五十万|一百万|百万|十万|50万|100万|20万|三十万)",
            # 竞业限制义务一概及于全体员工（超出第24条限定的人员范围）
            r"(?:所有|全部|全体)(?:员工|人员|劳动者|职工)[^。；\n]{0,30}(?:竞业|限制|不得|适用)",
            r"(?:适用于|约束于|及于)[^。；\n]{0,15}(?:所有|全部|全体)(?:员工|人员|职工)",
        ],
        # 有序匹配：先区分在职/离职后适用条件，再落到期限、违约金、人员范围。
        "risk_note": {
            r"在职(?:期间|期间内)?[^。；\n]{0,25}(?:竞业|限制)|"
            r"(?:竞业|限制)[^。；\n]{0,25}在职":
                "在职期间竞业限制与离职后竞业限制的适用条件不同，须分开判断："
                "《劳动合同法》第23条的按月经济补偿义务指向「**在解除或者终止劳动合同后**，"
                "在竞业限制期限内」按月给予；"
                "劳动者以不得约定在职期间竞业限制、未支付经济补偿为由请求确认在职竞业条款无效的，"
                "人民法院不予支持（法释〔2025〕12号第14条）。"
                "在职约定是否有效，仍须核验是否属第24条限定的竞业限制人员范围、"
                "以及是否与劳动者实际接触的商业秘密相适应（法释〔2025〕12号第13条）",
            r"离职后|解除(?:或|或者)终止劳动合同后|终止劳动合同后|劳动合同解除(?:后|终止后)":
                "离职后的竞业限制期间，用人单位应当按月给予劳动者经济补偿（《劳动合同法》第23条）；"
                "未约定补偿标准而劳动者已履行竞业限制义务的，"
                "可请求按劳动合同解除或者终止前十二个月平均工资的30%按月支付，"
                "低于劳动合同履行地最低工资标准的按最低工资标准支付（法释〔2020〕26号第36条）；"
                "劳动合同解除或者终止后因用人单位原因导致三个月未支付经济补偿，"
                "劳动者请求解除竞业限制约定的，人民法院应予支持（法释〔2020〕26号第38条）",
            r"三年|五年|十年|三年以上":
                "《劳动合同法》第24条：竞业限制期限**不得超过二年**",
            r"违约金|赔偿金":
                "竞业限制违约金可约定（《劳动合同法》第23条）；"
                "约定的违约金**过分高于**造成的损失的，人民法院或者仲裁机构"
                "可以根据当事人的请求予以适当减少（《民法典》第585条第2款）",
            r"商业秘密|所有|全部|全体|适用于":
                "《劳动合同法》第24条：竞业限制人员限于用人单位的高级管理人员、"
                "高级技术人员和其他负有保密义务的人员；"
                "劳动者未知悉、接触用人单位商业秘密和与知识产权相关的保密事项的，"
                "竞业限制条款不生效；范围、地域、期限与所接触事项不相适应的，"
                "超过合理比例部分无效（法释〔2025〕12号第13条）",
        },
    },
    "confidentiality": {
        "label": "保密义务",
        "section_headers": ["保密", "商业秘密", "保密条款"],
        "patterns": [
            r"保密(?:义务|条款|期限)[^。；\n]{0,60}",
            r"(?:商业秘密|技术秘密)[^。；\n]{0,40}(?:定义|包括|指)",
            r"不得(?:泄露|披露|公开)[^。；\n]{0,50}",
        ],
        "risk_high": [
            r"保密(?:期限|义务)[^。；\n]{0,20}(?:永久|无限期|长期)",
            r"违约金[^。；\n]{0,20}(?:保密|泄密)[^。；\n]{0,40}(?:五十万|一百万|十万)",
        ],
        "risk_note": {
            r"永久|无限期":
                "保密义务的存续期限可长于劳动合同期限；"
                "但约定的违约金**过分高于**实际损失的，人民法院或者仲裁机构"
                "可以根据当事人的请求予以适当减少（《民法典》第585条第2款）",
            r"违约金":
                "《劳动合同法》第25条：除第22条（专项培训服务期）、第23条（竞业限制）"
                "规定的情形外，用人单位不得与劳动者约定由劳动者承担违约金",
        },
    },
    "liquidated_damages": {
        "label": "违约金",
        "section_headers": ["违约金", "违约责任", "赔偿"],
        "patterns": [
            r"违约金[^。；\n]{0,60}",
            r"违约(?:责任|赔偿)[^。；\n]{0,60}",
        ],
        "risk_high": [
            r"(?:培训|专项培训)[^。；\n]{0,30}违约金",
            r"竞业限制[^。；\n]{0,30}违约金",
            r"违约金[^。；\n]{0,20}(?:工资|月薪|年薪)[^。；\n]{0,20}(?:十倍|二十倍|数倍)",
        ],
        "risk_note": {
            r"培训|专项培训":
                "《劳动合同法》第22条：仅**专项培训**可约定服务期与违约金，"
                "违约金数额不得超过用人单位提供的培训费用，"
                "且不得超过服务期尚未履行部分所应分摊的培训费用",
            r"竞业限制":
                "《劳动合同法》第23条：竞业限制可约定违约金；"
                "第24条：竞业限制期限不得超过二年",
            r"十倍|二十倍":
                "约定的违约金**过分高于**造成的损失的，人民法院或者仲裁机构"
                "可以根据当事人的请求予以适当减少（《民法典》第585条第2款）",
        },
    },
    "work_hours": {
        "label": "工时与休息休假",
        "section_headers": ["工作时间", "工时制度", "休息休假", "加班"],
        "patterns": [
            r"工时制度[^。；\n]{0,40}",
            r"(?:每日|每周)[^。；\n]{0,20}(?:工作|上班)[^。；\n]{0,20}(?:小时|天)",
            r"加班[^。；\n]{0,50}",
            r"(?:带薪)?年休假[^。；\n]{0,50}",
        ],
        "risk_high": [
            r"综合计算工时[^。；\n]{0,40}(?:未经|未)[^。；\n]{0,10}审批",
            r"不定时工作制[^。；\n]{0,40}(?:未经|未)[^。；\n]{0,10}审批",
            r"加班[^。；\n]{0,20}(?:不(?:予|无须|无需)[^。；\n]{0,15}支付|免费)",
            r"年休假[^。；\n]{0,30}(?:不(?:予|不安排|取消))",
        ],
        "risk_note": {
            r"综合计算工时|不定时工作制":
                "《劳动法》第39条：用人单位由于生产经营需要，经与工会和劳动者协商后"
                "可以延长工作时间，**延长工作时间每日不得超过三小时**；"
                "综合计算工时工作制、不定时工作制须经劳动行政部门批准，"
                "具体口径按《国务院关于职工工作时间的规定》第5条及地方审批口径核验",
            r"加班":
                "《劳动法》第44条：安排延长工作时间的支付不低于工资150%；"
                "休息日安排工作又不能安排补休的支付不低于200%；"
                "法定休假日安排工作的支付不低于300%",
            r"年休假":
                "《职工带薪年休假条例》：职工累计工作满1年不满10年的，年休假5天；"
                "满10年不满20年的，10天；满20年的，15天",
        },
    },
    "social_insurance": {
        "label": "社会保险",
        "section_headers": ["社会保险", "社保", "公积金", "五险一金"],
        "patterns": [
            r"(?:社会保险|社保)[^。；\n]{0,60}",
            r"(?:住房公积金|公积金)[^。；\n]{0,50}",
            r"(?:社保|社会保险)[^。；\n]{0,30}(?:自愿|放弃|代办|补贴)",
        ],
        "risk_high": [
            r"(?:自愿|同意|申请)[^。；\n]{0,15}(?:不缴|放弃)[^。；\n]{0,15}(?:社会保险|社保)",
            r"(?:社保|社会保险)[^。；\n]{0,30}(?:补贴|折现)[^。；\n]{0,30}元",
            r"以(?:现金|现金形式)[^。；\n]{0,20}(?:代替|折算)[^。；\n]{0,20}(?:社保|社会保险)",
        ],
        "risk_note": {
            r"自愿|同意|放弃":
                "用人单位与劳动者约定或者劳动者向用人单位承诺无需缴纳社会保险费的，"
                "人民法院应当认定该约定或者承诺**无效**（法释〔2025〕12号第19条）；"
                "用人单位未依法缴纳社会保险费，劳动者依《劳动合同法》第38条第一款第三项"
                "请求解除劳动合同并主张经济补偿的，人民法院依法予以支持。"
                "**现金补贴不能免除法定参保义务**",
            r"补贴|折现":
                "以现金补贴或折现方式替代缴纳社会保险费的约定无效；"
                "是否可用现金折抵已缴费用、如何补缴及滞纳金，按个案与地方征缴口径核验",
        },
    },
    "termination": {
        "label": "解除与终止",
        "section_headers": ["解除", "终止", "离职", "辞职", "裁员"],
        "patterns": [
            r"解除(?:劳动)?合同[^。；\n]{0,60}",
            r"终止(?:劳动)?合同[^。；\n]{0,60}",
            r"(?:员工|劳动者)(?:单方)?(?:辞职|离职)[^。；\n]{0,50}",
            r"经济(?:补偿|赔偿)[^。；\n]{0,50}",
        ],
        "risk_high": [
            r"(?:末位|排名末位|最后一名)[^。；\n]{0,30}(?:解除|辞退|终止)",
            r"试用期[^。；\n]{0,20}(?:无需|不需)[^。；\n]{0,15}(?:理由|说明|通知)",
            r"(?:孕期|产期|哺乳期|三期)[^。；\n]{0,40}(?:不得|禁止)[^。；\n]{0,20}(?:解除|辞退)[^。；\n]{0,30}任何",
            r"医疗期[^。；\n]{0,30}(?:不得|禁止)[^。；\n]{0,20}(?:解除|辞退)",
            r"(?:不需|无需)[^。；\n]{0,15}(?:提前)?30日[^。；\n]{0,15}通知",
        ],
        "risk_note": {
            r"末位|排名末位":
                "绩效考核居于末位不等于法律上的「不能胜任工作」："
                "第40条第2项要求「不能胜任工作，经过培训**或者**调整工作岗位，"
                "仍不能胜任工作」；以末位直接解除须按用工所在地裁判口径核验",
            r"三期.*不得解除任何":
                "《劳动合同法》第42条限制的是第40条、第41条；"
                "**协商解除（第36条）与过错解除（第39条）不受第42条限制**，"
                "条款表述过宽反而有风险",
            r"医疗期.*不得":
                "《劳动合同法》第42条限制第40条、第41条；"
                "医疗期满后不能从事原工作也不能从事另行安排工作的，"
                "可走第40条第1项（劳动者患病或非因工负伤）",
            r"无需30日":
                "《劳动合同法》第40条：用人单位提前**30日**书面通知劳动者本人"
                "**或者**额外支付劳动者一个月工资后，可以解除劳动合同",
        },
    },
    "delivery": {
        "label": "送达与通知",
        "section_headers": ["送达", "通知", "送达地址", "通讯地址"],
        "patterns": [
            r"送达(?:地址|方式|通知)[^。；\n]{0,60}",
            r"(?:书面|电子)?通知[^。；\n]{0,40}(?:地址|方式|邮箱|短信)",
            r"通讯地址[^。；\n]{0,50}",
        ],
        "risk_high": [
            r"(?:视为|即视为)[^。；\n]{0,30}(?:送达|收到|通知)[^。；\n]{0,40}(?:无需|不必)[^。；\n]{0,20}(?:回执|签收|确认)",
        ],
        "risk_note": {
            r"视为送达":
                "送达条款有效但需**实际履行**；仅约定而不寄送，送达争议时仍可能认定未通知",
        },
    },
    "job_change": {
        "label": "调岗与变更",
        "section_headers": ["调岗", "岗位调整", "工作地点", "变更"],
        "patterns": [
            r"调(?:整)?岗(?:位)?[^。；\n]{0,60}",
            r"(?:工作地点|岗位)[^。；\n]{0,40}(?:变更|调整|调动)",
            r"变更(?:劳动)?合同[^。；\n]{0,50}",
        ],
        "risk_high": [
            r"(?:公司|甲方|用人单位)[^。；\n]{0,25}(?:有权|可|可以)[^。；\n]{0,25}(?:随时|单方|自行)[^。；\n]{0,25}调(?:整)?岗",
            r"(?:工资|薪酬|薪资)[^。；\n]{0,20}(?:另行(?:商)?定|另行核定|不固定|浮动)",
            r"工作地点[^。；\n]{0,20}(?:全国|任意|各地|根据需要)",
            r"调岗[^。；\n]{0,30}(?:无需|不需)[^。；\n]{0,15}(?:协商|同意|征得)",
        ],
        "risk_note": {
            r"随时|单方":
                "《劳动合同法》第35条：变更劳动合同应当**协商一致**并采用书面形式，"
                "单方变更易被认定违法",
            r"全国|任意|根据需要":
                "工作地点约定过于宽泛，可能被认定为**约定不明**，反而不利于企业",
        },
    },
    "rules_validity": {
        "label": "规章制度效力",
        "section_headers": ["规章制度", "员工手册", "民主程序", "公示"],
        "patterns": [
            r"(?:规章|制度|员工手册)[^。；\n]{0,50}(?:生效|有效|制定)",
            r"职工代表(?:大会|会议)[^。；\n]{0,40}",
            r"(?:公示|告知)[^。；\n]{0,40}(?:员工|劳动者|全体)",
        ],
        "risk_high": [
            r"(?:公司|甲方)[^。；\n]{0,20}(?:有权|可)[^。；\n]{0,20}(?:随时|单方)[^。；\n]{0,20}(?:修改|修订|变更)[^。；\n]{0,20}(?:制度|规章|手册)",
            r"无需[^。；\n]{0,15}(?:民主程序|职工代表(?:大会|会议))",
        ],
        "risk_note": {
            r"随时|单方.*(?:修改|修订)":
                "《劳动合同法》第4条：涉及劳动者切身利益的规章制度须经"
                "**职工代表大会或者全体职工讨论**、与工会或者职工代表**平等协商**确定，"
                "并**公示或者告知**劳动者",
            r"无需民主程序":
                "《劳动合同法》第4条为强制性程序，未履行则制度不能作为管理依据",
        },
    },
    "recruitment": {
        "label": "录用条件",
        "section_headers": ["录用条件", "录用", "入职条件", "offer"],
        "patterns": [
            r"录用条件[^。；\n]{0,60}",
            r"(?:聘用|录用)[^。；\n]{0,40}条件[^。；\n]{0,50}",
        ],
        "risk_high": [
            r"录用条件[^。；\n]{0,40}(?:品德|操守|忠诚|德才兼备|胜任)[^。；\n]{0,20}(?:者)?[^。；\n]{0,20}(?:不予|拒绝)",
        ],
        "risk_note": {
            r"品德|忠诚|德才兼备":
                "录用条件应**具体、可衡量**并在订立劳动合同前告知；"
                "模糊表述无法举证，可能构成违法解除",
        },
    },
}

CLAUSE_ORDER = list(CLAUSE_TYPES.keys())

# --------------------------------------------------------------------------
# 数量与期限解析（中文数字 / 阿拉伯数字）
# --------------------------------------------------------------------------
_CN_DIGITS = {
    "零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
    "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
}

DURATION_RE = re.compile(
    r"(?P<num>[0-9]+(?:\.[0-9]+)?|[零〇一二两三四五六七八九十百]+)"
    r"\s*(?:个)?\s*(?P<unit>年|月|个月|日|天|周|星期)"
)


def cn_number_to_int(text):
    """解析有限非负数字；兼容旧调用名，返回整数或小数，不截断小数。"""
    text = str(text).strip()
    if re.fullmatch(r"[0-9]+(?:\.[0-9]+)?", text):
        value = float(text)
        return value if value < float("inf") else None
    if "点" in text:
        whole, fraction = text.split("点", 1)
        if not fraction or any(ch not in _CN_DIGITS for ch in fraction):
            return None
        base = cn_number_to_int(whole)
        return None if base is None else base + float("0." + "".join(str(_CN_DIGITS[ch]) for ch in fraction))
    if not text or any(ch not in _CN_DIGITS and ch not in "十百" for ch in text):
        return None
    section = number = 0
    for ch in text:
        if ch in _CN_DIGITS:
            number = _CN_DIGITS[ch]
        else:
            section += (number or 1) * (10 if ch == "十" else 100)
            number = 0
    return section + number


_NUMBER = r"(?:[0-9]+(?:\.[0-9]+)?|[零〇一二两三四五六七八九十百]+(?:点[零〇一二两三四五六七八九]+)?)"
_TERM_TOKEN = re.compile(
    r"(?:(?P<number>" + _NUMBER + r")\s*(?:个)?\s*(?P<before_half>半)?|(?P<half>半)(?:个)?)"
    r"\s*(?P<unit>年|月|日|天|周|星期)\s*(?P<after_half>半)?"
)


def read_term_after(text, head_pattern):
    """读取紧随条款主题的完整期限；不能解析时显式返回未知，不能取整前缀。"""
    head = re.search(head_pattern, text)
    if not head:
        return None
    tail = text[head.end():].split("，", 1)[0].split("。", 1)[0].split("；", 1)[0]
    # 只允许无数字的引导词，避免跳过无法识别的期限部分。
    lead = re.match(r"[^0-9零〇一二两三四五六七八九十百半]{0,16}", tail)
    pos = lead.end() if lead else 0
    first = pos
    total = 0.0
    matched = False
    while True:
        token = _TERM_TOKEN.match(tail, pos)
        if not token:
            break
        value = 0.5 if token.group("half") else cn_number_to_int(token.group("number"))
        if value is None:
            return {"months": None, "evidence": head.group(0) + tail}
        if token.group("before_half") or token.group("after_half"):
            value += 0.5
        if token.group("unit") not in ("年", "月"):
            return {"months": None, "evidence": head.group(0) + tail,
                    "reason": "按日/周的期限须结合实际起止日核验，不能固定按30天折月"}
        total += value * (12 if token.group("unit") == "年" else 1)
        matched = True
        pos = token.end()
    remainder = tail[pos:].lstrip()
    if not matched or (remainder and remainder[0] in "0123456789零〇一二两三四五六七八九十百半点年月个./~-至到"):
        return {"months": None, "evidence": head.group(0) + tail}
    return {"months": total, "evidence": head.group(0) + tail[first:pos]}


def duration_to_months(num_text, unit):
    """把期限统一折算为月数。无法判断时返回 None。"""
    value = cn_number_to_int(num_text)
    if value is None:
        return None
    unit = unit.replace("个", "")
    if unit == "年":
        return value * 12
    if unit == "月":
        return value
    if unit in ("日", "天"):
        return round(value / 30.0, 2)
    if unit in ("周", "星期"):
        return round(value * 7 / 30.0, 2)
    return None


# --------------------------------------------------------------------------
# 合同期限语境识别（文档级）
# --------------------------------------------------------------------------
CONTRACT_TERM_RE = re.compile(
    r"(?:本)?(?:劳动)?合同(?:期限|期间|年限)[^。；\n]{0,12}?"
    r"(?P<num>[0-9]+(?:\.[0-9]+)?|[零〇一二两三四五六七八九十百]+)"
    r"\s*(?:个)?\s*(?P<unit>年|月|个月|日|天)"
)

PROBATION_LEN_RE = re.compile(
    r"试用期(?:期限)?[^。；\n]{0,12}?(?:为|是|:|：)?\s*"
    r"(?P<num>[0-9]+(?:\.[0-9]+)?|[零〇一二两三四五六七八九十百]+)"
    r"\s*(?:个)?\s*(?P<unit>年|月|个月|日|天)"
)


def detect_contract_term(text):
    """合同类型与月/年期限；解析不完整时保留未知，交给人工核验。"""
    if re.search(r"以完成一定工作任务为期限", text):
        return {"kind": "task", "months": None, "evidence": "以完成一定工作任务为期限"}
    if re.search(r"无固定期限", text):
        return {"kind": "open_ended", "months": None, "evidence": "无固定期限"}
    term = read_term_after(text, r"(?:本)?(?:劳动)?合同(?:期限|期间|年限)")
    if term and term["months"] is not None:
        return {"kind": "fixed", "months": term["months"], "evidence": term["evidence"]}
    return {"kind": "unknown", "months": None, "evidence": term["evidence"] if term else ""}


def probation_cap_months(contract):
    """按合同期限计算法定试用期上限（月）。None 表示不得约定试用期。"""
    kind = contract["kind"]
    if kind == "task":
        return None
    if kind == "open_ended":
        return 6
    if kind == "fixed":
        months = contract["months"]
        if months < 3:
            return None
        if months < 12:
            return 1
        if months < 36:
            return 2
        return 6
    return None


CAP_TEXT = {
    1: "劳动合同期限三个月以上不满一年的，试用期不得超过一个月",
    2: "劳动合同期限一年以上不满三年的，试用期不得超过二个月",
    6: "三年以上固定期限和无固定期限的劳动合同，试用期不得超过六个月",
}

CAP_TEXT_NONE = ("以完成一定工作任务为期限的劳动合同或者劳动合同期限不满三个月的，"
                 "不得约定试用期")


def assess_probation(sentence, contract):
    """试用期与合同期限联合判断。

    返回 list of {trigger, note, level}
    - high：可依合同期限（或期限上限）直接判定超限
    - medium：合同期限未识别，无法判断是否超限（显式提示，不暗示无风险）
    """
    risks = []
    term = read_term_after(sentence, r"试用期(?:期限)?")
    if not term:
        return risks
    trigger = term["evidence"]
    months = term["months"]
    if months is None:
        return [{"trigger": trigger, "level": "medium",
                 "note": "试用期限未能完整按月/年解析，须核验原文和实际起止日，不得据此视为无风险。"}]

    cap = probation_cap_months(contract)
    kind = contract["kind"]

    if cap is None and kind in ("task", "fixed"):
        risks.append({
            "trigger": trigger,
            "note": "《劳动合同法》第19条第3款：%s。本条约定试用期%s。"
                    % (CAP_TEXT_NONE,
                       "与合同期限类型冲突" if kind == "task" else "超过法定上限"),
            "level": "high",
        })
        return risks

    if cap is not None:
        if months > cap:
            risks.append({
                "trigger": trigger,
                "note": "《劳动合同法》第19条第1款：%s（本案合同期限：%s）。"
                        "本条约定试用期%s，超过法定上限%s个月。"
                        % (CAP_TEXT[cap],
                           contract["evidence"] or kind,
                           trigger,
                           _fmt_months(months - cap)),
                "level": "high",
            })
        return risks

    # 合同期限未识别：上限不可判定（1/2/6 个月，或根本不得约定）
    if months > 6:
        risks.append({
            "trigger": trigger,
            "note": "《劳动合同法》第19条第1款：三年以上固定期限和无固定期限的劳动合同，"
                    "试用期不得超过六个月；本条约定试用期%s，**超过任何合同类型下的法定上限**，"
                    "无须再核合同期限即可判定违法约定。" % trigger,
            "level": "high",
        })
    else:
        risks.append({
            "trigger": trigger,
            "note": "未在文本中识别劳动合同期限，**无法判断**本条试用期是否超过法定上限"
                    "（1个月 / 2个月 / 6个月，取决于合同期限；"
                    "以完成一定工作任务为期限或期限不满3个月的还不得约定试用期）。"
                    "须补核合同期限后判定，不得据此认为无风险。",
            "level": "medium",
        })
    return risks


# 竞业限制期限的法定上限：二年 = 24 个月（《劳动合同法》第24条）
NONCOMPETE_MAX_MONTHS = 24

NONCOMPETE_TERM_HEAD_RE = re.compile(
    r"竞业(?:限制|禁止)[^。；\n]{0,20}?(?:期限|年限|期)"
)

# 期限数字串："两年零一个月" / "36个月" / "二年" / "5年"
_DURATION_TOKEN_RE = re.compile(
    r"(?P<y_num>[0-9]+(?:\.[0-9]+)?|[零〇一二两三四五六七八九十百]+)\s*(?:个)?\s*年"
    r"|"
    r"(?P<m_num>[0-9]+(?:\.[0-9]+)?|[零〇一二两三四五六七八九十百]+)\s*(?:个)?\s*(?P<m_unit>月|个月)"
)


def _parse_compound_months(y_num, m_num):
    """解析"X年Y个月""Y个月""X年"等期限，返回月数或 None。"""
    months = 0.0
    matched = False
    if y_num:
        v = cn_number_to_int(y_num)
        if v is None:
            return None
        months += v * 12
        matched = True
    if m_num:
        v = cn_number_to_int(m_num)
        if v is None:
            return None
        months += v
        matched = True
    return months if matched else None

IN_SERVICE_RE = re.compile(
    r"在职(?:期间|期间内)?[^。；\n]{0,25}(?:竞业|限制)|"
    r"(?:竞业|限制)[^。；\n]{0,25}在职"
)

POST_SERVICE_RE = re.compile(
    r"离职后|解除(?:或|或者)终止劳动合同后|终止劳动合同后|劳动合同解除(?:后|终止后)"
)


def assess_noncompete(sentence):
    """竞业限制期限的数值判定 + 在职/离职后分流。

    解决两类缺陷：
    1. 期限只按"三年/五年/十年"字样匹配，会漏掉"30个月""两年零一个月"等
       已超二年上限的表述——改为按月数数值比较。
    2. "不另行支付经济补偿"在未指明在职/离职后时，提示会退回通用兜底文案，
       无法给出适用条件——此处按语境分流，未指明时明确提示须先确定义务指向。
    """
    risks = []

    # 第24条第二款的二年上限仅针对离职后；混合/未知时间条件不直接定超限。
    in_service = bool(IN_SERVICE_RE.search(sentence))
    post_service = bool(POST_SERVICE_RE.search(sentence))
    term = read_term_after(sentence, r"竞业(?:限制|禁止)[^。；\n]{0,20}?(?:期限|年限|期)")
    if term:
        months = term["months"]
        if months is None:
            risks.append({"trigger": term["evidence"], "level": "medium",
                          "note": "竞业期限无法完整解析，须核验原文及起止日，不能据此判无风险。"})
        elif in_service and not post_service:
            risks.append({"trigger": term["evidence"], "level": "medium",
                          "note": "本条指向在职期间竞业，不直接适用离职后竞业的二年上限；须核验人员范围、秘密接触及约定合理性（法释〔2025〕12号第13、14条）。"})
        elif months > NONCOMPETE_MAX_MONTHS:
            if post_service and not in_service:
                note = "《劳动合同法》第24条第二款：解除或者终止劳动合同后的竞业限制期限不得超过二年；本条折合%s个月，超出24个月上限。" % _fmt_months(months)
                level = "high"
            else:
                note = "须先区分在职和离职后义务。若本条指向解除或者终止劳动合同后，第24条第二款规定竞业限制期限不得超过二年；当前不得跨时间条件直接作超限定性。"
                level = "medium"
            risks.append({"trigger": term["evidence"], "level": level, "note": note})

    # 2) 补偿义务的在职/离职后分流
    if re.search(r"(?:不[^。；\n]{0,4}(?:支付|发放|另付)|无需[^。；\n]{0,4}支付|"
                 r"没有[^。；\n]{0,6}(?:约定|支付)|未[^。；\n]{0,4}约定)[^。；\n]{0,20}"
                 r"(?:经济)?(?:补偿|报酬|费用|津贴)", sentence):
        in_service = bool(IN_SERVICE_RE.search(sentence))
        post_service = bool(POST_SERVICE_RE.search(sentence))
        if in_service and not post_service:
            risks.append({
                "trigger": "在职期间竞业限制未约定/未支付经济补偿",
                "note": "本条指向**在职期间**的竞业限制义务。《劳动合同法》第23条的"
                        "按月经济补偿义务指向「**在解除或者终止劳动合同后**，"
                        "在竞业限制期限内」按月给予；"
                        "劳动者以不得约定在职期间竞业限制、未支付经济补偿为由"
                        "请求确认在职竞业条款无效的，人民法院不予支持"
                        "（法释〔2025〕12号第14条）。"
                        "在职约定本身是否有效，仍须核验是否属第24条限定的竞业限制人员范围、"
                        "以及是否与劳动者实际接触的商业秘密相适应"
                        "（法释〔2025〕12号第13条）。"
                        "**不得据此推定离职后的竞业限制补偿义务已免除。**",
                "level": "high",
            })
        elif post_service and not in_service:
            risks.append({
                "trigger": "离职后竞业限制未约定/未支付经济补偿",
                "note": "本条指向**解除或终止劳动合同后**的竞业限制期间。"
                        "《劳动合同法》第23条要求用人单位在该竞业限制期限内"
                        "按月给予劳动者经济补偿；"
                        "未约定补偿标准而劳动者已履行竞业限制义务的，"
                        "可请求按劳动合同解除或者终止前十二个月平均工资的30%按月支付，"
                        "低于劳动合同履行地最低工资标准的按最低工资标准支付"
                        "（法释〔2020〕26号第36条）；"
                        "解除或终止后因用人单位原因导致三个月未支付经济补偿，"
                        "劳动者请求解除竞业限制约定的，人民法院应予支持"
                        "（法释〔2020〕26号第38条）。",
                "level": "high",
            })
        else:
            risks.append({
                "trigger": "竞业限制未约定/未支付经济补偿（义务指向未指明）",
                "note": "本条**未指明**竞业限制义务发生在在职期间还是解除或终止劳动合同后，"
                        "两种情形的适用条件不同，不能一概处理："
                        "若为**在职期间**竞业限制，《劳动合同法》第23条的按月补偿义务"
                        "并不当然适用（法释〔2025〕12号第14条）；"
                        "若为**离职后**竞业限制，则应按月给予经济补偿，"
                        "未约定补偿标准而劳动者已履行的，可请求按前十二个月平均工资的30%"
                        "按月支付（法释〔2020〕26号第36条），"
                        "解除或终止后三个月未支付的，劳动者可请求解除竞业限制约定"
                        "（法释〔2020〕26号第38条）。"
                        "**须先确定义务指向的时间点再判断，不得笼统适用。**",
                "level": "high",
            })

    return risks


def _fmt_months(value):
    if value is None:
        return ""
    if abs(value - round(value)) < 1e-6:
        return "%d" % int(round(value))
    return "%.2f" % value


# --------------------------------------------------------------------------
# 文件读取
# --------------------------------------------------------------------------
def read_document(path):
    """读取文档为纯文本。返回 (text, fmt)；失败抛 RuntimeError（受控错误）。"""
    ext = os.path.splitext(path)[1].lower()

    if ext in (".md", ".txt", ".markdown", ".text"):
        try:
            with io.open(path, encoding="utf-8", errors="replace") as f:
                return f.read(), "text"
        except (IOError, UnicodeDecodeError) as exc:
            raise RuntimeError("读取文件失败：%s" % exc)

    if ext == ".docx":
        try:
            import docx  # noqa
        except ImportError:
            raise RuntimeError(
                "读取 .docx 需要 python-docx 库。\n"
                "  安装：pip install python-docx\n"
                "  或：将文档另存为 .txt / .md 后重试（不影响主流程）")
        try:
            d = docx.Document(path)
        except Exception as exc:  # 损坏 / 非 OPC 包 / 加密文档
            raise RuntimeError(
                "文档无法解析（可能已损坏、不是有效的 .docx，或已加密）：%s" % exc)
        try:
            parts = [p.text for p in d.paragraphs]
            for t in d.tables:
                for row in t.rows:
                    parts.append("\t".join(c.text for c in row.cells))
        except Exception as exc:
            raise RuntimeError("文档内容提取失败：%s" % exc)
        return "\n".join(parts), "docx"

    raise RuntimeError("暂不支持的文件格式：%s（支持 .md/.txt/.docx）" % ext)


def split_sentences(text):
    """按中文标点切句，保留所属章节标题。"""
    items = []
    current_section = ""
    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        is_header = line.startswith("#") or (
            len(line) <= 30 and
            any(k in line for k in ("第", "条", "章", "节")) and
            not line.endswith(("。", "；", "，"))
        )
        if is_header:
            current_section = line.lstrip("#").strip()
            continue

        for sent in re.split(r"[。；\n]", line):
            sent = sent.strip()
            if len(sent) >= 6:
                items.append((sent, current_section))
    return items


# --------------------------------------------------------------------------
# 抽取
# --------------------------------------------------------------------------
def extract(text, only_types=None, min_risk="all"):
    """抽取条款。min_risk: all / medium / high

    同一句可同时命中多个条款类型（多标签），不再因首个命中而 break。
    """
    sentences = split_sentences(text)
    contract = detect_contract_term(text)
    results = {k: [] for k in CLAUSE_ORDER}
    scanned = only_types or CLAUSE_ORDER

    for sent, section in sentences:
        for ctype in scanned:
            spec = CLAUSE_TYPES[ctype]

            matched = False
            for pat in spec["patterns"]:
                if re.search(pat, sent, re.I):
                    matched = True
                    break
            if not matched:
                continue

            risks = []
            for pat in spec["risk_high"]:
                m = re.search(pat, sent, re.I)
                if m:
                    note = ""
                    for note_pat, note_text in spec["risk_note"].items():
                        if re.search(note_pat, sent, re.I):
                            note = note_text
                            break
                    risks.append({
                        "trigger": m.group(0),
                        "note": note or "建议核验现行法条要求",
                        "level": "high",
                    })

            if ctype == "probation":
                risks.extend(assess_probation(sent, contract))
            elif ctype == "noncompete":
                risks.extend(assess_noncompete(sent))

            if ctype in ("confidentiality", "liquidated_damages") and "违约金" in sent:
                permitted_topic = bool(re.search(r"竞业(?:限制|禁止)|专项培训|专业技术培训|培训服务期", sent))
                if not permitted_topic:
                    risks = [{"trigger": "劳动者承担违约金的约定", "level": "high",
                              "note": "须先核验能否约定违约金：《劳动合同法》第25条规定，除第22条专项培训服务期和第23条竞业限制情形外，用人单位不得与劳动者约定由劳动者承担违约金。单独保密义务不因此成为允许约定违约金的例外；不能仅按金额过高建议调低。"}]

            levels = {r["level"] for r in risks}
            if min_risk == "high" and "high" not in levels:
                continue
            if min_risk == "medium" and not levels & {"high", "medium"}:
                continue

            results[ctype].append({
                "section": section,
                "text": sent,
                "risks": risks,
            })

    return {k: v for k, v in results.items() if v}


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------
LEVEL_MARK = {
    "high": "🔴 高风险线索",
    "medium": "🟡 待人工核验线索",
    "info": "🔵 提示",
}


def print_human(results, meta):
    print("=" * 66)
    print("  用工合同关键条款抽取 · labor-compliance-assistant")
    print("=" * 66)
    print("文件：%s" % meta["file"])
    print("格式：%s ｜ 抽取时间：%s" % (meta["format"], meta["extracted_at"]))
    print("命中条款类型：%d 个" % len(results))

    total_risks = 0
    for ctype, items in results.items():
        spec = CLAUSE_TYPES[ctype]
        risk_n = sum(len(i["risks"]) for i in items)
        total_risks += risk_n
        print("\n" + "-" * 66)
        print("【%s】%s（%d 处，风险线索 %d 条）"
              % (spec["label"], ctype, len(items), risk_n))
        print("-" * 66)
        for i, item in enumerate(items, 1):
            loc = item["section"] or "（正文）"
            print("\n  %d. [%s] %s" % (i, loc, item["text"][:110]))
            for r in item["risks"]:
                print("     %s  命中：%s"
                      % (LEVEL_MARK.get(r["level"], "[?]"), r["trigger"]))
                if r["note"]:
                    print("       %s" % r["note"])

    print("\n" + "=" * 66)
    print("风险线索合计：%d 条" % total_risks)
    if total_risks:
        print("\n⚠️ 以上为**抽取线索**，不是法律结论。")
        print("   每一项都必须回 A1 权威源核验法条原文后方可定性，")
        print("   并按 knowledge-integrity.md 记录差异。")
    else:
        print("\n⚠️ 未抽取到风险线索，**不等于**该文档没有法律风险。")
        print("   表格、扫描件、图片与复杂句式均可能漏抽，请回原文逐条阅读。")
    print("=" * 66)


def build_payload(results, meta, status, error=None):
    """统一的 JSON 输出契约。"""
    return {
        "meta": meta,
        "status": status,
        "clauses": [
            {
                "type": ctype,
                "label": CLAUSE_TYPES[ctype]["label"],
                "items": items,
            }
            for ctype, items in results.items()
        ],
        "error": error,
        "disclaimer": "抽取线索非法律结论，须回权威源核验法条后定性",
    }


def main():
    ap = argparse.ArgumentParser(
        description="用工合同关键条款抽取（可选增强，失败不影响主流程）")
    ap.add_argument("path", help="文档路径（.md/.txt/.docx）")
    ap.add_argument("--json", action="store_true", dest="as_json",
                    help="以 JSON 输出")
    ap.add_argument("--type", help="仅抽取指定类型，逗号分隔。可选：%s"
                                  % ",".join(CLAUSE_ORDER))
    ap.add_argument("--min-risk", choices=["all", "medium", "high"],
                    default="all", help="只保留含风险线索的条款")
    args = ap.parse_args()

    meta = {
        "file": os.path.basename(args.path),
        "path": os.path.abspath(args.path),
        "format": "unknown",
        "extracted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "types_scanned": [],
    }

    def emit_json(status, error=None):
        print(json.dumps(build_payload({}, meta, status, error),
                         ensure_ascii=False, indent=2))

    if not os.path.isfile(args.path):
        message = "文件不存在：%s" % args.path
        if args.as_json:
            emit_json("error", {"code": "FILE_NOT_FOUND", "message": message,
                                "degraded": True,
                                "hint": "降级：直接以文本方式阅读原文继续评估，不阻塞主流程"})
        else:
            print("错误：%s" % message, file=sys.stderr)
            print("降级建议：直接以文本方式阅读原文继续评估，不阻塞主流程。",
                  file=sys.stderr)
        return 1

    only_types = None
    if args.type is not None:
        only_types = [t.strip() for t in args.type.split(",") if t.strip()]
        if not only_types:
            message = "--type 未指定任何有效条款类型（仅包含分隔符或空白）"
            if args.as_json:
                emit_json("error", {"code": "EMPTY_TYPE_SET", "message": message,
                                    "degraded": False,
                                    "hint": "可选类型：%s" % ",".join(CLAUSE_ORDER)})
            else:
                print("错误：%s" % message, file=sys.stderr)
                print("可选类型：%s" % ",".join(CLAUSE_ORDER), file=sys.stderr)
            return 2
        unknown = [t for t in only_types if t not in CLAUSE_TYPES]
        if unknown:
            message = "未知条款类型：%s" % ", ".join(unknown)
            if args.as_json:
                emit_json("error", {"code": "UNKNOWN_TYPE", "message": message,
                                    "degraded": False,
                                    "hint": "可选类型：%s" % ",".join(CLAUSE_ORDER)})
            else:
                print("错误：%s" % message, file=sys.stderr)
                print("可选类型：%s" % ",".join(CLAUSE_ORDER), file=sys.stderr)
            return 2

    meta["types_scanned"] = only_types or CLAUSE_ORDER

    try:
        text, fmt = read_document(args.path)
    except RuntimeError as exc:
        message = str(exc)
        if args.as_json:
            emit_json("error", {"code": "DOCUMENT_READ_FAILED", "message": message,
                                "degraded": True,
                                "hint": "降级：直接以文本方式阅读原文继续评估，不阻塞主流程"})
        else:
            print("无法读取文档：%s" % message, file=sys.stderr)
            print("\n降级建议：直接以文本方式阅读原文继续评估，不阻塞主流程。",
                  file=sys.stderr)
        return 1
    except (IOError, UnicodeDecodeError) as exc:
        message = "读取失败：%s" % exc
        if args.as_json:
            emit_json("error", {"code": "DOCUMENT_READ_FAILED", "message": message,
                                "degraded": True,
                                "hint": "降级：直接以文本方式阅读原文继续评估，不阻塞主流程"})
        else:
            print("读取失败：%s" % message, file=sys.stderr)
        return 1

    meta["format"] = fmt

    results = extract(text, only_types, args.min_risk)

    if not results:
        if args.as_json:
            emit_json("no_match")
        else:
            print("未抽取到已知类型的用工条款。")
            print("提示：可尝试去掉 --type 与 --min-risk 参数，或确认文档内容是否相关。")
            print("⚠️ 未命中不等于无风险：表格、扫描件、图片与复杂句式均可能漏抽。")
        return 0

    if args.as_json:
        print(json.dumps(build_payload(results, meta, "ok"),
                         ensure_ascii=False, indent=2))
    else:
        print_human(results, meta)
    return 0


if __name__ == "__main__":
    sys.exit(main())
