# -*- coding: utf-8 -*-
"""
maintain.py — 清单动态维护一键脚本。

改完 references/ 下的 .md 后，用一条命令重建索引与要求 JSON，避免漏跑某个脚本。

依赖关系（一个 JSON ← 一个或多个 .md 源）：
  references_index.json  ← 水利标准.md + 法律法规.md   （build_reference_index.build）
  requirements.json      ← 检查要求.md                  （parse_requirements.parse）

用法：
  python scripts/maintain.py                 # 智能重建：仅重建"源 .md 比 JSON 新"或 JSON 缺失的项
  python scripts/maintain.py --force         # 强制重建全部（忽略时间戳）
  python scripts/maintain.py --only index    # 只重建标准/法规索引
  python scripts/maintain.py --only req      # 只重建检查要求
  python scripts/maintain.py --check         # 仅检查是否过期、不重建（退出码 0=最新, 1=需重建）
  python scripts/maintain.py --backup        # 重建前把旧 JSON 备份为 .bak
  python scripts/maintain.py --force --backup # 强制重建并备份
"""

from __future__ import annotations
import argparse
import json
import os
import shutil
import sys

# GBK 控制台自愈：打印 ✓/✗ 等会 UnicodeEncodeError 且产物不落盘（2026-09-30 实测）
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass
from collections import OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_reference_index as bri   # noqa: E402
import parse_requirements as pr       # noqa: E402


def _here():
    return os.path.dirname(os.path.abspath(__file__))


def _skill_root():
    return os.path.normpath(os.path.join(_here(), ".."))


# 依赖表：每个产物 JSON 依赖一份或多份 .md 源
SOURCES = {
    "index": {
        "json": os.path.join(_here(), "references_index.json"),
        "mds": [os.path.join(_skill_root(), "references", "水利标准.md"),
                os.path.join(_skill_root(), "references", "法律法规.md")],
        "label": "标准/法规索引",
    },
    "req": {
        "json": os.path.join(_here(), "requirements.json"),
        "mds": [os.path.join(_skill_root(), "references", "检查要求.md")],
        "label": "检查要求",
    },
}


def is_stale(key):
    """源 .md 比产物 JSON 新、或产物/源缺失，则视为需要重建。"""
    cfg = SOURCES[key]
    if not os.path.exists(cfg["json"]):
        return True
    jtime = os.path.getmtime(cfg["json"])
    for md in cfg["mds"]:
        if not os.path.exists(md):
            return True  # 源缺失：交给 rebuild 报错，这里先判为"需处理"
        if os.path.getmtime(md) > jtime:
            return True
    return False


def _missing_md(cfg):
    return [md for md in cfg["mds"] if not os.path.exists(md)]


def _backup(path):
    bak = path + ".bak"
    if os.path.exists(path):
        shutil.copy2(path, bak)
        print(f"[maintain] 备份 {os.path.basename(path)} -> {os.path.basename(bak)}")


def _old_counts(key):
    """读取旧 JSON 的条目数，用于重建后比对骤降。"""
    cfg = SOURCES[key]
    try:
        d = json.load(open(cfg["json"], encoding="utf-8"))
    except Exception:
        return None
    if key == "index":
        return {"standards": len(d.get("standards", [])), "laws": len(d.get("laws", []))}
    return {"requirements": len(d) if isinstance(d, list) else 0}


def rebuild(key, backup=False):
    """重建指定项。返回 True/False 表示成功与否。"""
    cfg = SOURCES[key]
    missing = _missing_md(cfg)
    if missing:
        print(f"[maintain] ✗ {cfg['label']}：源文件缺失，无法重建：")
        for m in missing:
            print(f"          - {m}")
        return False

    if backup:
        _backup(cfg["json"])
    old = _old_counts(key)

    try:
        if key == "index":
            idx = bri.build(cfg["mds"][0], cfg["mds"][1], cfg["json"])
            new = {"standards": idx["stats"]["standards"], "laws": idx["stats"]["laws"]}
            print(f"[maintain] ✓ 重建 {cfg['label']}：标准 {new['standards']} 条，"
                  f"法规 {new['laws']} 条 -> {os.path.basename(cfg['json'])}")
        else:
            reqs = pr.parse(cfg["mds"][0])
            json.dump(reqs, open(cfg["json"], "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
            new = {"requirements": len(reqs)}
            by_cat = OrderedDict()
            for r in reqs:
                by_cat.setdefault(r["category"], []).append(r["id"])
            print(f"[maintain] ✓ 重建 {cfg['label']}：共 {new['requirements']} 条 "
                  f"-> {os.path.basename(cfg['json'])}")
            for cat, ids in by_cat.items():
                print(f"          {cat}: {len(ids)} 条 ({ids[0]}..{ids[-1]})")
    except Exception as e:
        print(f"[maintain] ✗ 重建 {cfg['label']} 失败：{e}")
        return False

    # 骤降告警：条目数较旧值腰斩，疑似 .md 表格格式被破坏
    if old:
        for k, v in new.items():
            ov = old.get(k)
            if ov and v < ov * 0.5:
                print(f"[maintain] ⚠ 警告：{k} 由 {ov} 骤降至 {v}，"
                      f"请检查源 .md 表格格式是否被破坏")
    return True


def run_check(keys):
    """仅检查是否过期，不重建。退出码 0=全部最新，1=存在过期。"""
    stale = [k for k in keys if is_stale(k)]
    for k in keys:
        cfg = SOURCES[k]
        status = "需重建" if is_stale(k) else "最新"
        print(f"[maintain] {cfg['label']:<8} {status}  ({os.path.basename(cfg['json'])})")
    if stale:
        print(f"[maintain] 检测到 {len(stale)} 项过期 → 运行: python scripts/maintain.py")
        sys.exit(1)
    print("[maintain] 全部最新，无需重建")
    sys.exit(0)


def main():
    ap = argparse.ArgumentParser(
        description="清单动态维护一键重建脚本（改完 .md 后重建索引/要求 JSON）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例:\n"
               "  python scripts/maintain.py            # 智能重建过期项\n"
               "  python scripts/maintain.py --force    # 强制重建全部\n"
               "  python scripts/maintain.py --check    # 只检查是否过期\n"
               "  python scripts/maintain.py --backup   # 重建前备份旧 JSON\n")
    ap.add_argument("--force", action="store_true", help="强制重建全部（忽略时间戳）")
    ap.add_argument("--only", choices=["index", "req"],
                    help="只重建某一项：index(标准/法规索引) 或 req(检查要求)")
    ap.add_argument("--check", action="store_true",
                    help="仅检查是否过期、不重建（退出码 0=最新, 1=需重建）")
    ap.add_argument("--backup", action="store_true",
                    help="重建前把旧 JSON 备份为 .bak")
    args = ap.parse_args()

    keys = [args.only] if args.only else ["index", "req"]

    if args.check:
        run_check(keys)

    did, failed, skipped = [], [], []
    for k in keys:
        if not (args.force or is_stale(k)):
            print(f"[maintain] 跳过 {SOURCES[k]['label']}（已是最新，--force 可强制）")
            skipped.append(k)
            continue
        ok = rebuild(k, backup=args.backup)
        (did if ok else failed).append(k)

    print("-" * 52)
    if did:
        print(f"[maintain] 已重建: {', '.join(SOURCES[k]['label'] for k in did)}")
    if skipped:
        print(f"[maintain] 已跳过: {', '.join(SOURCES[k]['label'] for k in skipped)}")
    if failed:
        print(f"[maintain] 失败:   {', '.join(SOURCES[k]['label'] for k in failed)}")
        sys.exit(1)
    if not did:
        print("[maintain] 全部最新，未重建（加 --force 可强制）")
    else:
        print("[maintain] 提示：已生成的旧校审报告不会自动重审，"
              "需重跑校审流程才按新规则复评")


if __name__ == "__main__":
    main()
