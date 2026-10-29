#!/usr/bin/env python3
"""
labor-compliance-assistant · 知识层元数据体检脚本

用途
----
本脚本对 references/ 下的知识层文件做**元数据体检**，检查：

  1. 元数据完整性  —— 缺 layer / last-checked / confidence 等关键字段
  2. 核验时效       —— last-checked 距今超过设定天数（STALE / AGING）
  3. 日期合法性     —— last-checked 无法解析 / 晚于今日
  4. 枚举合法性     —— layer / confidence 是否在允许取值内
  5. 失效触发登记   —— E1 文件是否登记 staleness-trigger
  6. 草稿标记       —— confidence 为 draft / medium 的高风险内容
  7. 交叉引用       —— E1 文件之间是否互相引用（静态启发式）
  8. 空目录         —— references 目录下是否一个 .md 都没有

【能力边界（重要）】
本脚本**只做元数据体检**。它：
  · 不联网、不抓取法规、不比对法律正文语义；
  · **不能**自动发现法律修订，也不能证明知识层与现行法一致；
  · 「体检通过」只表示元数据自洽，**不等于**内容正确或法律有效。
  法律变更的发现与核验，只能由使用者在检索阶段按
  `kb-search-protocol.md` 的取源顺序回 A1 权威源完成。

输出
----
- 终端摘要
- 可选 Markdown 报告（--report 路径）
- 可选 JSON（--json）
退出码：0 无阻塞项；1 存在阻塞项；2 参数错误

用法
----
python knowledge_sync_check.py
python knowledge_sync_check.py --days 90 --report ./sync-report.md
python knowledge_sync_check.py --json
"""

import argparse
import datetime as dt
import io
import json
import os
import re
import sys

DEFAULT_MAX_AGE_DAYS = 180
BLOCKING_CONFIDENCE = ("draft",)

# 允许的元数据枚举值
VALID_CONFIDENCE = ("high", "medium", "draft")
# 本地知识层允许声明的信任层级。
# 注意：A1 指权威源（法条原文），本地文件无权自行声明为 A1——
# 自称 A1 会造成"用自己写的总结证明自己"的循环，因此视为非法声明。
VALID_LOCAL_LAYER = ("E1", "E2")

# 方法层文件：本技能的操作方法与体例，不承载法律实体内容，
# 因此不要求 frontmatter，也不参与 E1 核验与失效检测。
# 只有列在此白名单之外的 .md，才被当作「待确认的知识层候选」要求元数据。
METHOD_LAYER_FILES = {
    "kb-search-protocol.md",
    "public-source-policy.md",
    "labor-domains.md",
    "labor-law-registry.md",
    "dimensions.md",
    "legal-issue-tree.md",
    "question-bank.md",
    "risk-grading.md",
    "report-template.md",
    "routing-table.md",
    "glossary-external.md",
    "review-cycle.md",
    "knowledge-integrity.md",
    "data-handling.md",
    "quality-checklist.md",
    "document-processing.md",
    "jurisdiction-scope.md",
    "organization-policy-interface.md",
}

FRONTMATTER_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n", re.S)
LIST_ITEM_RE = re.compile(r"^\s*-\s+(.+?)\s*$")


# --------------------------------------------------------------------------
# 轻量 frontmatter 解析（标准库，不依赖 PyYAML）
# --------------------------------------------------------------------------
def parse_frontmatter(text):
    """解析 YAML frontmatter，返回 (dict, bool)。无 frontmatter 时返回 ({}, False)。"""
    m = FRONTMATTER_RE.match(text)
    if not m:
        return {}, False

    data = {}
    current_key = None
    current_list = None

    for line in m.group(1).split("\n"):
        if not line.strip() or line.strip().startswith("#"):
            continue

        li = LIST_ITEM_RE.match(line)
        if li and current_list is not None:
            current_list.append(li.group(1).strip().strip("'\""))
            continue

        if ":" not in line:
            continue

        key, _, val = line.partition(":")
        key = key.strip()
        val = val.strip().strip("'\"")

        if val == "":
            current_key = key
            current_list = []
            data[key] = current_list
        else:
            data[key] = val
            current_key = None
            current_list = None

    # 标量字段若被误解析成空列表，回退为 ""
    for k, v in list(data.items()):
        if isinstance(v, list) and not v:
            data[k] = ""

    return data, True


def parse_date(value):
    if not value:
        return None
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", str(value))
    if not m:
        return None
    try:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


# --------------------------------------------------------------------------
# 扫描
# --------------------------------------------------------------------------
def scan_references(ref_dir):
    findings = []
    files_meta = []

    if not os.path.isdir(ref_dir):
        return findings, files_meta

    md_files = [fn for fn in sorted(os.listdir(ref_dir))
                if fn.endswith(".md") and os.path.isfile(os.path.join(ref_dir, fn))]

    if not md_files:
        findings.append({
            "level": "BLOCK", "file": "(目录)", "code": "EMPTY-REFERENCES",
            "message": "references 目录下没有任何 .md 文件，无法体检",
            "action": "确认目录路径是否正确；若知识层已被移除或移动，请补回文件或登记为方法层",
        })
        return findings, files_meta

    for fn in md_files:
        path = os.path.join(ref_dir, fn)
        try:
            text = io.open(path, encoding="utf-8").read()
        except (IOError, UnicodeDecodeError) as exc:
            findings.append({
                "level": "WARN", "file": fn, "code": "READ-ERROR",
                "message": "无法读取文件：%s" % exc,
                "action": "检查文件编码（应为 UTF-8）与权限",
            })
            continue

        meta, has_fm = parse_frontmatter(text)
        files_meta.append({"file": fn, "meta": meta, "has_fm": has_fm})

        if not has_fm:
            if fn in METHOD_LAYER_FILES:
                # 方法层文件：设计如此，不报提示
                pass
            else:
                # 未登记的方法层文件 + 无元数据 = 疑似新增的知识层内容漏登记
                findings.append({
                    "level": "WARN", "file": fn, "code": "META-MISSING",
                    "message": "非白名单文件但无元数据头，无法判定其信任层级与核验状态",
                    "action": "补 frontmatter（layer/last-checked/confidence/staleness-trigger），"
                              "或在 knowledge_sync_check.py 的 METHOD_LAYER_FILES 中登记为方法层",
                })
            continue

        layer = meta.get("layer", "")

        # 3.1 元数据完整性
        required = ["layer", "last-checked", "confidence"]
        missing = [k for k in required if not meta.get(k)]
        if missing:
            findings.append({
                "level": "BLOCK", "file": fn, "code": "META-INCOMPLETE",
                "message": "缺少必填字段：%s" % ", ".join(missing),
                "action": "补齐 frontmatter 字段，参考 knowledge-integrity.md §3.1",
            })

        # 3.2 核验时效
        checked = parse_date(meta.get("last-checked"))
        if checked is None and meta.get("last-checked"):
            findings.append({
                "level": "BLOCK", "file": fn, "code": "DATE-UNPARSEABLE",
                "message": "last-checked 无法解析：%s" % meta.get("last-checked"),
                "action": "统一使用 YYYY-MM-DD 格式；日期不可解析即无法判定核验时效，"
                          "不得按「已核验」对待",
            })

        # 3.3 失效触发条件
        triggers = meta.get("staleness-trigger")
        if isinstance(triggers, str):
            triggers = []
        if layer == "E1" and not triggers:
            findings.append({
                "level": "BLOCK", "file": fn, "code": "NO-STALENESS-TRIGGER",
                "message": "E1 知识层文件未登记 staleness-trigger，失效无法感知",
                "action": "补 staleness-trigger，至少含 statute-amended 与 judicial-interpretation-updated",
            })

        # 3.4 草稿与低置信
        conf = str(meta.get("confidence", "")).lower()
        if conf and conf not in VALID_CONFIDENCE:
            findings.append({
                "level": "BLOCK", "file": fn, "code": "CONFIDENCE-INVALID",
                "message": "confidence 取值非法：%s（允许：%s）"
                           % (meta.get("confidence"), " / ".join(VALID_CONFIDENCE)),
                "action": "改为 high / medium / draft 之一；非法取值不得被当作已核验",
            })
        elif conf in BLOCKING_CONFIDENCE:
            findings.append({
                "level": "BLOCK", "file": fn, "code": "DRAFT-CONFIDENCE",
                "message": "confidence=draft，尚未系统核验",
                "action": "逐条回 A1 权威源核验后上调置信度；引用时必须披露未核验",
            })
        elif conf == "medium":
            findings.append({
                "level": "WARN", "file": fn, "code": "MEDIUM-CONFIDENCE",
                "message": "confidence=medium，含未完全核验内容",
                "action": "引用该文件时须向用户提示未核验范围",
            })

        # 3.5 枚举合法性：layer
        if layer and layer not in VALID_LOCAL_LAYER:
            level_msg = "layer 取值非法：%s" % layer
            action = "本地知识层只能声明 %s；A1 指权威源原文，本地文件无权自封 A1" \
                     % " / ".join(VALID_LOCAL_LAYER)
            if layer == "A1":
                level_msg += "（自称 A1 会形成自证循环，须改为 E1/E2 并登记核验依据）"
            findings.append({
                "level": "BLOCK", "file": fn, "code": "LAYER-INVALID",
                "message": level_msg,
                "action": action,
            })

        files_meta[-1]["checked"] = checked
        files_meta[-1]["confidence"] = conf

    return findings, files_meta


def check_age(findings, files_meta, max_age_days, today):
    """时效检查需要 max_age_days，单独跑以便 CLI 控制。"""
    for item in files_meta:
        checked = item.get("checked")
        if checked is None:
            continue
        age = (today - checked).days
        item["age_days"] = age
        if age < 0:
            findings.append({
                "level": "BLOCK", "file": item["file"], "code": "DATE-IN-FUTURE",
                "message": "last-checked 为未来日期（%s），晚于今日 %s"
                           % (checked.isoformat(), today.isoformat()),
                "action": "核对核验日期是否填错；未来日期不得被当作「已核验且在有效期内」",
            })
        elif age > max_age_days:
            findings.append({
                "level": "WARN", "file": item["file"], "code": "STALE",
                "message": "距上次核验已 %d 天（阈值 %d 天）" % (age, max_age_days),
                "action": "重新核验依据法条，更新 last-checked 与 confidence",
            })
        elif age > max_age_days * 0.7:
            findings.append({
                "level": "INFO", "file": item["file"], "code": "AGING",
                "message": "距上次核验已 %d 天，接近阈值" % age,
                "action": "纳入下一轮核验计划",
            })


# --------------------------------------------------------------------------
# 交叉引用一致性（静态启发式）
# --------------------------------------------------------------------------
CROSS_REF_HINTS = {
    "dismissal-pathway.md": ["evidence-checklist.md", "risk-grading.md",
                             "kb-search-protocol.md", "knowledge-integrity.md"],
    "evidence-checklist.md": ["dismissal-pathway.md", "kb-search-protocol.md",
                              "risk-grading.md"],
    "kb-search-protocol.md": ["dismissal-pathway.md", "evidence-checklist.md",
                              "public-source-policy.md", "labor-law-registry.md"],
}


def check_cross_refs(ref_dir, files_meta):
    """交叉引用一致性（静态启发式：仅检查正文是否提及对方文件名）。

    只在「两者都承载实质法律内容」时才要求互相引用。
    方法层文件（如 kb-search-protocol）本就是被引用的对象，
    反过来要求它引用别人没有意义，会产生大量误报，故排除。
    """
    present = {m["file"] for m in files_meta}
    notes = []

    for src, targets in CROSS_REF_HINTS.items():
        if src not in present or src in METHOD_LAYER_FILES:
            continue
        fp = os.path.join(ref_dir, src)
        try:
            text = io.open(fp, encoding="utf-8").read()
        except (IOError, UnicodeDecodeError):
            continue
        for tgt in targets:
            if tgt not in present:
                continue
            # 方法层目标本就无需被反向引用
            if tgt in METHOD_LAYER_FILES:
                continue
            # 去掉 frontmatter 后再判断正文是否提及
            body = FRONTMATTER_RE.sub("", text)
            if tgt not in body:
                notes.append((src, tgt,
                              "同为 E1 知识层但正文未互相引用，"
                              "修改其一可能造成两份经验总结相互矛盾"))
    return notes


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------
LEVEL_ICON = {"BLOCK": "[阻塞]", "WARN": "[警告]", "INFO": "[提示]"}


def print_summary(findings, files_meta, cross_notes, today):
    print("=" * 62)
    print("  知识层元数据体检 · labor-compliance-assistant")
    print("  检测日期：%s" % today.isoformat())
    print("=" * 62)

    e1 = [m for m in files_meta if str(m.get("meta", {}).get("layer")) == "E1"]
    print("\n已扫描 reference 文件：%d 个（其中 E1 知识层 %d 个）"
          % (len(files_meta), len(e1)))

    if e1:
        print("\n--- E1 知识层状态 ---")
        for m in e1:
            meta = m["meta"]
            age = m.get("age_days")
            age_txt = "未记录" if age is None else ("%d 天前" % age)
            print("  · %-26s conf=%-7s 核验=%s"
                  % (m["file"], meta.get("confidence", "?"), age_txt))
            va = meta.get("verified-against")
            if isinstance(va, list) and va:
                print("      已核验法条 %d 项，首项：%s" % (len(va), va[0]))

    if not findings:
        print("\n✅ 元数据体检未发现问题。")
    else:
        order = {"BLOCK": 0, "WARN": 1, "INFO": 2}
        findings.sort(key=lambda f: order.get(f["level"], 3))
        print("\n--- 发现 %d 项 ---" % len(findings))
        for f in findings:
            print("\n%s %s  [%s]" % (LEVEL_ICON.get(f["level"], "[?]"),
                                       f["file"], f["code"]))
            print("    问题：%s" % f["message"])
            print("    建议：%s" % f["action"])

    if cross_notes:
        print("\n--- 交叉引用提示（%d 项）---" % len(cross_notes))
        for src, tgt, msg in cross_notes:
            print("  · %s → %s：%s" % (src, tgt, msg))

    blocks = sum(1 for f in findings if f["level"] == "BLOCK")
    warns = sum(1 for f in findings if f["level"] == "WARN")
    print("\n" + "=" * 62)
    print("汇总：阻塞 %d 项 ｜ 警告 %d 项" % (blocks, warns))
    if blocks:
        print("结论：存在阻塞项，知识层不可作为分析依据，须先处理。")
    else:
        print("结论：元数据自洽，知识层可用于分析（结论仍须回 A1 权威源支撑）。")
    print("⚠️ 本脚本只做元数据体检，**不检测法律修订、不比对法条正文**；")
    print("   「体检通过」不等于「内容与现行法一致」。")
    print("=" * 62)
    return blocks


def write_report(path, findings, files_meta, cross_notes, today, max_age_days):
    lines = []
    lines.append("# 知识层元数据体检报告")
    lines.append("")
    lines.append("- 检测日期：%s" % today.isoformat())
    lines.append("- 时效阈值：%d 天" % max_age_days)
    lines.append("- 扫描文件：%d 个" % len(files_meta))
    lines.append("")
    lines.append("> 本报告只反映**元数据**状态（字段完整性、日期合法性、枚举取值、"
                 "时效、交叉引用），**不反映法律内容是否正确**。")
    lines.append("> 本脚本不联网、不比对法条正文，不能自动发现法律修订。")
    lines.append("")

    lines.append("## E1 知识层状态")
    lines.append("")
    lines.append("| 文件 | 置信度 | 核验日期 | 距今天数 | 已核验法条数 |")
    lines.append("|---|---|---|---|---|")
    for m in files_meta:
        if str(m.get("meta", {}).get("layer")) != "E1":
            continue
        meta = m["meta"]
        age = m.get("age_days")
        va = meta.get("verified-against")
        va_n = len(va) if isinstance(va, list) else 0
        lines.append("| %s | %s | %s | %s | %d |" % (
            m["file"], meta.get("confidence", "?"),
            meta.get("last-checked", "?"),
            "-" if age is None else str(age), va_n))
    lines.append("")

    lines.append("## 检测发现")
    lines.append("")
    if not findings:
        lines.append("未发现问题。")
    else:
        order = {"BLOCK": 0, "WARN": 1, "INFO": 2}
        findings.sort(key=lambda f: order.get(f["level"], 3))
        lines.append("| 级别 | 文件 | 代码 | 问题 | 建议动作 |")
        lines.append("|---|---|---|---|---|")
        for f in findings:
            lines.append("| %s | %s | %s | %s | %s |" % (
                LEVEL_ICON.get(f["level"], "?"), f["file"], f["code"],
                f["message"].replace("|", "\\|"),
                f["action"].replace("|", "\\|")))
    lines.append("")

    if cross_notes:
        lines.append("## 交叉引用提示")
        lines.append("")
        for src, tgt, msg in cross_notes:
            lines.append("- `%s` → `%s`：%s" % (src, tgt, msg))
        lines.append("")

    lines.append("---")
    lines.append("")
    lines.append("> 本报告由 `knowledge_sync_check.py` 生成。")
    lines.append("> 阻塞项须处理后方可将对应知识层用于分析。")
    lines.append("> 任何确定性结论仍须以 A1 权威源（法条原文）为依据。")

    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d)
    io.open(path, "w", encoding="utf-8").write("\n".join(lines))
    print("\n报告已写入：%s" % path)


def main():
    ap = argparse.ArgumentParser(
        description="labor-compliance-assistant 知识层元数据体检"
                    "（只检查元数据，不检测法律修订、不比对法条正文）")
    ap.add_argument("--days", type=int, default=DEFAULT_MAX_AGE_DAYS,
                    help="核验时效阈值（天），默认 %d" % DEFAULT_MAX_AGE_DAYS)
    ap.add_argument("--report", metavar="PATH",
                    help="同时输出 Markdown 报告到指定路径")
    ap.add_argument("--json", action="store_true", dest="as_json",
                    help="以 JSON 输出，便于被其他流程消费")
    ap.add_argument("--ref-dir", metavar="DIR",
                    help="references 目录路径，默认为脚本上级目录下的 references/")
    args = ap.parse_args()

    if args.days <= 0:
        print("错误：--days 必须为正整数", file=sys.stderr)
        return 2

    here = os.path.dirname(os.path.abspath(__file__))
    ref_dir = args.ref_dir or os.path.join(os.path.dirname(here), "references")

    if not os.path.isdir(ref_dir):
        print("错误：找不到 references 目录：%s" % ref_dir, file=sys.stderr)
        return 2

    today = dt.date.today()
    findings, files_meta = scan_references(ref_dir)
    check_age(findings, files_meta, args.days, today)
    cross_notes = check_cross_refs(ref_dir, files_meta)

    if args.as_json:
        print(json.dumps({
            "checked_at": today.isoformat(),
            "max_age_days": args.days,
            "scope": "metadata-only",
            "capability_note": "本工具只做元数据体检：不联网、不比对法条正文、"
                               "不能自动发现法律修订；体检通过不等于内容与现行法一致",
            "files": len(files_meta),
            "findings": findings,
            "cross_refs": [{"src": s, "tgt": t, "note": n}
                           for s, t, n in cross_notes],
            "blocking": sum(1 for f in findings if f["level"] == "BLOCK"),
        }, ensure_ascii=False, indent=2))
        return 1 if any(f["level"] == "BLOCK" for f in findings) else 0

    blocks = print_summary(findings, files_meta, cross_notes, today)

    if args.report:
        write_report(args.report, findings, files_meta, cross_notes, today, args.days)

    return 1 if blocks else 0


if __name__ == "__main__":
    sys.exit(main())
