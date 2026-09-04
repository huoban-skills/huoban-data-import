#!/usr/bin/env python3
"""源文件评估：列清单、行数、空值率，以及按主键判重。

只读源文件，不写伙伴云。用于阶段一算数据量和阶段二的源表层检查。

  python3 get_source_profile.py --source data.xlsx
  python3 get_source_profile.py --source data.xlsx --key '单位类型,单位名称'
  python3 get_source_profile.py --source data.xlsx --key '协议编号' --dump-dup dup.xlsx
"""
import argparse
import sys
from collections import defaultdict

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from _common import KEY_SEP, build_key, check_columns, has_blank, normalize, raw_key, read_rows, split_cols, write_xlsx


def profile_columns(header, rows):
    total = len(rows)
    out = []
    for col in header:
        values = [r.get(col) for r in rows]
        filled = [v for v in values if normalize(v)]
        distinct = len({normalize(v) for v in filled})
        samples = []
        for v in filled:
            s = str(v).strip()
            if s not in samples:
                samples.append(s)
            if len(samples) == 3:
                break
        out.append(
            {
                "列名": col,
                "非空": len(filled),
                "空值率": f"{(total - len(filled)) / total * 100:.1f}%" if total else "-",
                "去重值": distinct,
                "样例": " / ".join(s[:20] for s in samples),
            }
        )
    return out


def find_duplicates(rows, key_cols):
    """返回 (重复组, 主键含空值的行, 去重后条数)。

    空值参与拼接：两行同样空着就是真重复。但含空值的行数要单独报，
    因为主键里有空列意味着这一列作为关联依据时匹配不上，主键可能定错了。
    """
    groups = defaultdict(list)
    blank = []
    for i, r in enumerate(rows):
        groups[build_key(r, key_cols)].append(i)
        if has_blank(r, key_cols):
            blank.append(i)
    dups = {k: v for k, v in groups.items() if len(v) > 1}
    return dups, blank, len(groups)


def classify(rows, idxs, key_cols, header):
    """区分整行完全重复与主键相同但其他列不同。"""
    others = [c for c in header if c not in key_cols]
    seen = {tuple(normalize(rows[i].get(c)) for c in others) for i in idxs}
    return "整行重复" if len(seen) == 1 else "主键冲突"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True, help="源文件绝对路径")
    ap.add_argument("--key", help="主键列名，逗号分隔")
    ap.add_argument("--sheet", type=int, default=1)
    ap.add_argument("--title-row", type=int, default=1)
    ap.add_argument("--encoding", default="utf-8")
    ap.add_argument("--dump-dup", help="把重复行明细写到这个 xlsx，重复行标黄")
    args = ap.parse_args()

    header, rows = read_rows(args.source, args.sheet, args.title_row, args.encoding)
    print(f"总行数：{len(rows)}    列数：{len(header)}\n")

    print("列清单")
    print(f"{'列名':<20}{'非空':>8}{'空值率':>10}{'去重值':>10}  样例")
    for c in profile_columns(header, rows):
        print(f"{c['列名']:<20}{c['非空']:>8}{c['空值率']:>10}{c['去重值']:>10}  {c['样例']}")

    if not args.key:
        print("\n未指定 --key，跳过判重")
        return

    key_cols = split_cols(args.key)
    check_columns(header, key_cols, "源文件")
    dups, blank, distinct = find_duplicates(rows, key_cols)

    print(f"\n按主键 {' | '.join(key_cols)} 判重")
    print(f"  去重后条数：{distinct}")
    print(f"  重复组：{len(dups)}    涉及行数：{sum(len(v) for v in dups.values())}")
    if blank:
        ratio = len(blank) / len(rows)
        print(f"  主键含空列的行：{len(blank)}（{ratio*100:.1f}%）")
        for c in key_cols:
            n = sum(1 for r in rows if not normalize(r.get(c)))
            if n:
                print(f"    {c} 为空：{n} 行")
        if ratio > 0.2:
            print("  这一列空值太多，作为关联依据会大面积匹配不上，回头确认主键是不是定宽了")

    if not dups:
        return

    kinds = defaultdict(int)
    detail = []
    for k, idxs in dups.items():
        if not k.replace(KEY_SEP, ""):
            kind = "主键全空"
        else:
            kind = classify(rows, idxs, key_cols, header)
        kinds[kind] += 1
        detail.append((kind, raw_key(rows[idxs[0]], key_cols), [i + args.title_row + 1 for i in idxs]))

    for kind, n in kinds.items():
        print(f"  {kind}：{n} 组")

    print("\n样例（前 5 组）")
    for kind, k, lines in detail[:5]:
        shown = ", ".join(map(str, lines[:8])) + (f" … 共 {len(lines)} 行" if len(lines) > 8 else "")
        print(f"  [{kind}] {k}  行号 {shown}")

    if args.dump_dup:
        out_rows = []
        highlight = set()
        for kind, k, lines in detail:
            for ln in lines:
                r = rows[ln - args.title_row - 1]
                if kind == "主键冲突":
                    highlight.add(len(out_rows))
                out_rows.append([kind, k, ln] + [r.get(c) for c in header])
        write_xlsx(args.dump_dup, ["重复类型", "主键", "行号"] + header, out_rows, highlight)
        print(f"\n重复明细已写入 {args.dump_dup}，主键冲突行已标黄")


if __name__ == "__main__":
    main()
