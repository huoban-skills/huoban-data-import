#!/usr/bin/env python3
"""按裁决结果把本批已导入记录的关联字段改指到别的记录，逐条 update-item，落记录文件。

  python3 relink.py --table 2100000080269274 --field '对应药品' \
    --links 待判药品-删除前关联备份.json --links-key links --item-key 协议item_id --old-key 药品item_id \
    --map 映射.json --out 回填记录.json [--verify --field-id 2200000615218546] [--preview]

links：JSON，列表或含列表的对象（--links-key 指定键），每项至少有 item-key（要改的记录）和 old-key（原来指向的关联记录）。
map：JSON 对象 {原关联记录id: 新关联记录id}；不在 map 里的链接跳过。
只处理 links 里列出的记录，即本批导入建的；不扫描表里其它数据。
用 update-item 而不用 update-items-by-ids：后者整条校验必填，其它必填为空的行会被拒。
hac 认证走环境变量 HUOBAN_ACCESS_TOKEN / HUOBAN_COMPANY_ID。
"""
import argparse
import json
import subprocess
import sys


def hac(args, timeout=120):
    o = subprocess.run(["hac", "table"] + args + ["--format", "json"], capture_output=True, text=True, timeout=timeout)
    raw = (o.stdout or "") + (o.stderr or "")
    try:
        return json.loads(o.stdout), raw
    except Exception:
        return None, raw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--table", required=True)
    ap.add_argument("--field", required=True, help="关联字段名")
    ap.add_argument("--links", required=True)
    ap.add_argument("--links-key", help="links 是对象时取哪个键下的列表")
    ap.add_argument("--item-key", required=True)
    ap.add_argument("--old-key", required=True)
    ap.add_argument("--map", required=True)
    ap.add_argument("--out", required=True, help="改指记录 JSON，追加写")
    ap.add_argument("--preview", action="store_true", help="只跑 dry-run，不写入")
    ap.add_argument("--verify", action="store_true", help="改完全表回读，确认没有记录再指向旧 id（需 --field-id）")
    ap.add_argument("--field-id")
    args = ap.parse_args()

    links = json.load(open(args.links, encoding="utf-8"))
    if args.links_key:
        links = links[args.links_key]
    mapping = {str(k): str(v) for k, v in json.load(open(args.map, encoding="utf-8")).items()}
    tasks = [(str(l[args.item_key]), mapping[str(l[args.old_key])]) for l in links if str(l.get(args.old_key)) in mapping]
    print(f"links {len(links)} 条，命中映射 {len(tasks)} 条")

    res = []
    for pid, t in tasks:
        f = json.dumps({args.field: [t]}, ensure_ascii=False)
        dr, raw = hac(["update-item", "--table-id", args.table, "--item-id", pid, "--fields", f, "--dry-run"])
        if not (dr and dr.get("ok")):
            res.append({"item_id": pid, "目标": t, "ok": False, "stage": "dry-run", "err": raw[:300]})
            continue
        if args.preview:
            res.append({"item_id": pid, "目标": t, "ok": True, "stage": "preview"})
            continue
        ap_, raw = hac(["update-item", "--table-id", args.table, "--item-id", pid, "--fields", f, "--yes"])
        res.append({"item_id": pid, "目标": t, "ok": bool(ap_ and ap_.get("ok")), "stage": "apply", "err": None if ap_ and ap_.get("ok") else raw[:300]})
    ok = sum(1 for r in res if r["ok"])
    print(f"{'预览' if args.preview else '改指'} {ok}/{len(res)}")
    for r in res:
        if not r["ok"]:
            print("  失败", r["item_id"], r["err"])

    try:
        prev = json.load(open(args.out, encoding="utf-8"))
        if not isinstance(prev, list):
            prev = [prev]
    except Exception:
        prev = []
    json.dump(prev + res, open(args.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    if args.verify and not args.preview:
        if not args.field_id:
            raise SystemExit("--verify 需要 --field-id")
        olds = set(mapping)
        items, off = [], 0
        while True:
            d, _ = hac(["list-items", "--table-id", args.table, "--limit", "50", "--offset", str(off),
                        "--order", '{"item_id":"asc"}', "--output-mode", "full"], timeout=300)
            batch = (d.get("data") or d)["items"]
            items += batch
            off += len(batch)
            if len(batch) < 50:
                break
        still = [it["item_id"] for it in items for v in (it["fields"].get(args.field_id) or []) if str(v.get("item_id")) in olds]
        print(f"回读 {len(items)} 条，仍指向旧记录的：{len(still)} {still[:10]}")
        sys.exit(1 if still else 0)


if __name__ == "__main__":
    main()
