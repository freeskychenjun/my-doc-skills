# -*- coding: utf-8 -*-
"""
json_append.py — subagent 增量写盘助手（groupX.json 骨架创建 + 逐条追加 + 自检）

背景（2026-09-30 赣东北可研实战）：大文档语义审查的单次 subagent 运行可能中途
死于输出超长（stopReason=length）/ turn 上限 / 30 分钟超时。"全部判完后一次性写
groupX.json"会让已完成的判定随运行一起丢失。纪律：先写骨架，每判完 1 条立即追加，
运行死亡时已写条目全部可抢救（实战抢回 2/3 以上条目）。
另：禁止用 Edit 工具追加多行中文 JSON——锚点易失配（实测），一律走本脚本。

用法：
  1) 判定开始前创建骨架（文件已存在且有结果时拒绝，防止覆盖可抢救成果）：
     python "$SKILL/scripts/json_append.py" <group.json> init "强制性条文检查"

  2) 每判完 1 条，把该条 JSON 从 stdin 传入追加（heredoc 防引号地狱）：
     python "$SKILL/scripts/json_append.py" <group.json> <<'EOF'
     {"id":"mp_001","name":"设计洪水计算过程检查","verdict":"部分符合",
      "conclusion":"……","evidence":[{"loc":"P0123","quote":"……"}],
      "issues":[{"loc":"P0238","original":"……","type":"错字","suggestion":"……"}]}
     EOF

  3) 全部写完后自检（打印条数与 id 清单；可带期望条数，不足即报错）：
     python "$SKILL/scripts/json_append.py" <group.json> verify [期望条数]

行为：追加时 id 重复默认报错（--replace 覆盖重写该条）；verdict 非法仅警告
（collect 阶段会归一）；目标文件损坏（非法 JSON）时拒绝写入，提示人工检查。
退出码：0 正常 / 1 错误。
"""

from __future__ import annotations
import json
import sys

VERDICTS = {"符合", "部分符合", "不符合", "不适用"}


def _console_safe():
    # GBK 控制台下打印 ✓/✗ 会 UnicodeEncodeError 且中断写入，errors=replace 自愈
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(errors="replace")
        except Exception:
            pass


def _die(msg):
    print(f"[json_append] ✗ {msg}")
    sys.exit(1)


def _load(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    except (json.JSONDecodeError, OSError) as e:
        _die(f"{path} 不是合法 JSON（{e}）——拒绝写入。请人工检查该文件，勿再盲试。")


def _save(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def main():
    _console_safe()
    args = sys.argv[1:]
    if not args:
        print("用法: python json_append.py <group.json> init <category>")
        print("      python json_append.py <group.json> [--replace] <<'EOF' …一条 result JSON… EOF   # 默认 append")
        print("      python json_append.py <group.json> verify [期望条数]")
        sys.exit(1)
    # mode 可省略：仅给文件名（stdin 有内容）默认 append，最贴合 heredoc 用法
    path = args[0]
    if len(args) >= 2 and args[1] in ("init", "append", "verify"):
        mode, rest = args[1], args[2:]
    else:
        mode, rest = "append", args[1:]

    if mode == "init":
        category = rest[0] if rest else ""
        if not category:
            _die("init 需要给出 category，如：init \"强制性条文检查\"")
        old = _load(path)
        if old is not None:
            n = len(old.get("results", [])) if isinstance(old, dict) else "?"
            _die(f"{path} 已存在（{n} 条结果）——为保护可抢救成果拒绝覆盖；确要重建请先删文件")
        _save(path, {"category": category, "results": []})
        print(f"[json_append] ✓ 骨架已创建 {path}（category={category}）")

    elif mode == "verify":
        data = _load(path)
        if not isinstance(data, dict) or not isinstance(data.get("results"), list):
            _die(f"{path} 缺 {{category, results}} 结构")
        results = data["results"]
        ids = [r.get("id", "?") for r in results]
        print(f"[json_append] ✓ {path} 合法，共 {len(results)} 条: {ids}")
        if rest:
            try:
                expect = int(rest[0])
            except ValueError:
                _die("期望条数应为整数")
            if len(results) < expect:
                _die(f"条数不足：已写 {len(results)} < 应写 {expect}——继续补判追加")

    elif mode == "append":
        replace = "--replace" in rest
        raw = sys.stdin.read().strip()
        if not raw:
            _die("stdin 为空——应经 heredoc 传入一条 result 的 JSON")
        try:
            item = json.loads(raw)
        except json.JSONDecodeError as e:
            _die(f"stdin 不是合法 JSON（{e}）。收到的开头 200 字：{raw[:200]!r}")
        if not isinstance(item, dict) or not item.get("id"):
            _die("result 缺 id 字段，拒绝追加")
        if item.get("verdict") not in VERDICTS:
            print(f"[json_append] ⚠ verdict={item.get('verdict')!r} 非标准"
                  f"（符合/部分符合/不符合/不适用），仍写入，collect 阶段将归一")
        data = _load(path)
        if not isinstance(data, dict) or not isinstance(data.get("results"), list):
            _die(f"{path} 缺 {{category, results}} 骨架——先运行 init")
        results = data["results"]
        ids = [r.get("id") for r in results]
        if item["id"] in ids:
            if not replace:
                _die(f"id={item['id']} 已存在；确要覆盖重写该条请加 --replace")
            results[ids.index(item["id"])] = item
            _save(path, data)
            print(f"[json_append] ✓ {item['id']} 已覆盖重写（--replace），当前 {len(results)} 条")
        else:
            results.append(item)
            _save(path, data)
            print(f"[json_append] ✓ {item['id']} 已写入，当前共 {len(results)} 条: {ids + [item['id']]}")

    else:
        _die(f"未知模式 {mode!r}（可用：init / append / verify）")


if __name__ == "__main__":
    main()
