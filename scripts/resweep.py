#!/usr/bin/env python3
"""导入后复扫：把本批新建的关联表记录与表里其它记录再比一遍，找出评估时漏掉的相似记录。

关联表快照先按 list-items 全量落成 CSV（含主键字段、创建人/创建时间列），再喂给本脚本。只读文件，不写伙伴云。

  python3 resweep.py --snapshot 药品快照.csv --key '药品名称,规格' \
    --mine-col 创建人 --mine 詹达富 --since-col 创建时间 --since 2026-09-03 \
    --aux '生产厂家（上市持有人）' --outdir ./out --label 药品

本批 = 满足 --mine / --since 任一组条件的记录；分几天导的同一批一起算。系统侧 = 其余全部记录。
"""
import argparse
import os
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from _common import KEY_SEP, check_columns, keys_similar, raw_key, read_rows, split_cols, write_xlsx_sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--snapshot", required=True, help="关联表当前全量快照（csv/xlsx），需含主键列与创建人/创建时间列")
    ap.add_argument("--key", required=True, help="主键字段，逗号分隔，与评估阶段一致")
    ap.add_argument("--id-col", default="item_id")
    ap.add_argument("--mine-col", help="创建人列名")
    ap.add_argument("--mine", help="本批创建人，逗号分隔可多个")
    ap.add_argument("--since-col", help="创建时间列名")
    ap.add_argument("--since", help="本批起始时间（含），如 2026-09-03")
    ap.add_argument("--aux", help="两侧并排展示的对照列，逗号分隔；只供人判断")
    ap.add_argument("--cutoff", type=float, default=0.75)
    ap.add_argument("--outdir", default=".")
    ap.add_argument("--label", default="关联表")
    ap.add_argument("--sheet", type=int, default=1)
    ap.add_argument("--title-row", type=int, default=1)
    args = ap.parse_args()
    if not (args.mine or args.since):
        raise SystemExit("至少给 --mine 或 --since 之一来圈定本批记录")

    header, rows = read_rows(args.snapshot, args.sheet, args.title_row)
    keys = split_cols(args.key)
    aux = split_cols(args.aux) if args.aux else []
    check_columns(header, keys + aux + [args.id_col] + [c for c in (args.mine_col, args.since_col) if c], "快照")
    mine_names = set(split_cols(args.mine)) if args.mine else set()

    def is_mine(r):
        by_name = bool(mine_names) and str(r.get(args.mine_col) or "") in mine_names
        by_time = bool(args.since) and str(r.get(args.since_col) or "") >= args.since
        return by_name or by_time

    mine = [r for r in rows if is_mine(r)]
    others = [r for r in rows if not is_mine(r)]
    print(f"快照 {len(rows)} 条：本批新建 {len(mine)}，系统侧 {len(others)}")

    out = []
    for m in mine:
        mp = [str(m.get(c) or "") for c in keys]
        best, score = None, 0.0
        for o in others:
            ok, sc = keys_similar(mp, [str(o.get(c) or "") for c in keys], args.cutoff)
            if ok and sc > score:
                best, score = o, sc
        if best:
            out.append((m, best, score))
    out.sort(key=lambda x: -x[2])
    print(f"疑似与系统重复：{len(out)} 条")
    for m, o, sc in out:
        print(f"  {sc:.2f}  {raw_key(m, keys)}  ⇔  {raw_key(o, keys)}")

    os.makedirs(args.outdir, exist_ok=True)
    path = os.path.join(args.outdir, f"{args.label}-复扫.xlsx")
    aux_header = [h for c in aux for h in (f"本批·{c}", f"系统·{c}")]
    hdr = ["本批新建（" + KEY_SEP.join(keys) + "）", "本批item_id", "系统最接近的（" + KEY_SEP.join(keys) + "）", "系统item_id", "相似度"] + aux_header + ["确认结果"]
    body = [[raw_key(m, keys), m.get(args.id_col), raw_key(o, keys), o.get(args.id_col), f"{sc:.2f}"]
            + [v for c in aux for v in (m.get(c), o.get(c))] + [""] for m, o, sc in out]
    write_xlsx_sheets(path, [("结论", hdr, body, None, {"确认结果": ["是同一个", "不是同一个"]}, ["系统最接近的（" + KEY_SEP.join(keys) + "）"])])
    print(f"复扫清单已写入 {path}；判「是同一个」的用 relink.py 把本批导入记录的关联改指系统记录，再删本批新建的那条")


if __name__ == "__main__":
    main()
