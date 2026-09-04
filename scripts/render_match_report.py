#!/usr/bin/env python3
"""源文件主键与关联表现有数据比对，输出可匹配 / 相似待确认 / 完全无 三档。

关联表数据先用 hac table +export 导到本地再喂给本脚本。只读文件，不写伙伴云。

  python3 render_match_report.py \
    --source 协议.xlsx --source-key '医疗机构名称' \
    --target 往来单位.xlsx --target-key '单位类型,单位名称' \
    --outdir ./out --label 医疗机构
"""
import argparse
import os
import sys
from collections import defaultdict
from difflib import SequenceMatcher

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from _common import KEY_SEP, build_key, check_columns, has_blank, keys_similar, normalize, raw_key, read_rows, split_cols, write_xlsx_sheets


def best_similar(src_key, candidates, cutoff):
    """在候选真实主键里找最像的一个。口径见 _common.keys_similar：逐段比对，与业务字段无关。"""
    parts = src_key.split(KEY_SEP)
    best, best_score = None, 0.0
    for cand in candidates:
        ok, score = keys_similar(parts, cand.split(KEY_SEP), cutoff)
        if ok and score > best_score:
            best, best_score = cand, score
    return best, best_score


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    ap.add_argument("--source-key", required=True, help="源文件里构成主键的列，逗号分隔")
    ap.add_argument("--target", required=True, help="关联表导出文件")
    ap.add_argument("--target-key", required=True, help="关联表里构成主键的字段，逗号分隔")
    ap.add_argument("--outdir", default=".")
    ap.add_argument("--label", default="关联表", help="用于命名输出文件")
    ap.add_argument("--cutoff", type=float, default=0.75, help="单段相似度阈值，默认 0.75；等价段（去括号后相同、数字等价）不受阈值影响")
    ap.add_argument("--const", action="append", default=[], help="给源文件注入常量列，格式 列名=值，可重复；用于源表没有但主键需要的维度（如单位类型）")
    ap.add_argument("--drop-blank-keys", action="store_true", help="跳过主键含空列的源行；用于该关联本身选填、空值行不参与匹配的场景")
    ap.add_argument("--carry", help="待导入结论 sheet 随主键一起带出的源列，逗号分隔；写 源列=关联表字段 可改列名，写 all 带出与关联表同名的全部非主键源列；同一主键多值时用；连接。凡源文件有、关联表也有的列都应带上，只导主键会漏字段")
    ap.add_argument("--aux", help="疑似匹配文件里两侧并排展示的对照列，逗号分隔，格式 源列=关联表字段（如 生产厂家=生产厂家（上市持有人））；只供人判断，不参与相似判定")
    ap.add_argument("--judged-new", help="已裁决为「不是同一个」的源主键清单文件（每行一个，含常量维度的完整键），这些键跳过相似判断直接归入完全无，避免复检时再被抓成疑似")
    ap.add_argument("--sheet", type=int, default=1)
    ap.add_argument("--title-row", type=int, default=1)
    ap.add_argument("--target-sheet", type=int, default=1)
    ap.add_argument("--target-title-row", type=int, default=1)
    args = ap.parse_args()

    s_header, s_rows = read_rows(args.source, args.sheet, args.title_row)
    t_header, t_rows = read_rows(args.target, args.target_sheet, args.target_title_row)
    const_cols = []
    for item in args.const:
        col, _, val = item.partition("=")
        const_cols.append(col)
        s_header = list(s_header) + [col]
        for r in s_rows:
            r[col] = val
    s_cols, t_cols = split_cols(args.source_key), split_cols(args.target_key)
    check_columns(s_header, s_cols, "源文件")
    check_columns(t_header, t_cols, "关联表导出文件")
    if len(s_cols) != len(t_cols):
        raise SystemExit(f"两侧主键列数不一致：源 {len(s_cols)} 列，关联表 {len(t_cols)} 列")

    def parse_pairs(text):
        out = []
        for item in split_cols(text or ""):
            src, _, tgt = item.partition("=")
            out.append((src, tgt or src))
        return out

    aux_pairs = parse_pairs(args.aux)
    for sc, tc in aux_pairs:
        check_columns(s_header, [sc], "源文件")
        check_columns(t_header, [tc], "关联表导出文件")

    # 关联表侧建索引，顺带查它自己有没有重复。空列以空串参与拼接，两侧同空才算配上
    target_index = defaultdict(list)
    for r in t_rows:
        target_index[build_key(r, t_cols)].append(r)
    t_dups = {k: v for k, v in target_index.items() if len(v) > 1}

    print(f"关联表 {args.label}：{len(t_rows)} 条，去重后 {len(target_index)} 个主键")
    if t_dups:
        print(f"  警告：关联表自身有 {len(t_dups)} 个主键重复，匹配会随机命中其中一条")
        for k in list(t_dups)[:3]:
            print(f"    {k}  {len(t_dups[k])} 条")

    # 源侧按主键去重后再比对，同一个值不重复算
    source_groups = defaultdict(list)
    blank = 0
    for i, r in enumerate(s_rows):
        if has_blank(r, s_cols):
            blank += 1
            if args.drop_blank_keys:
                continue
        source_groups[build_key(r, s_cols)].append(i)

    # 相似度只算真实主键部分：注入的常量维度两边同串，参与打分会把不相干的名字抬过阈值
    const_vals = {}
    for item in args.const:
        col, _, val = item.partition("=")
        const_vals[col] = normalize(val)
    const_pos = {i: const_vals[c] for i, c in enumerate(s_cols) if c in const_vals}

    def real_part(key):
        parts = key.split(KEY_SEP)
        return KEY_SEP.join(p for i, p in enumerate(parts) if i not in const_pos)

    def const_match(key):
        parts = key.split(KEY_SEP)
        return all(parts[i] == v for i, v in const_pos.items())

    # 同一真实主键有多条候选时，优先常量维度也一致的那条
    cand_map = {}
    for k in target_index:
        rp = real_part(k)
        if rp not in cand_map or const_match(k):
            cand_map[rp] = k

    judged_new = set()
    if args.judged_new:
        with open(args.judged_new, encoding="utf-8") as f:
            judged_new = {KEY_SEP.join(normalize(p) for p in line.rstrip("\n").split(KEY_SEP)) for line in f if line.strip()}

    matched, similar, missing = [], [], []
    candidates = list(cand_map.keys())
    for k, idxs in source_groups.items():
        if k in target_index:
            matched.append((k, idxs))
            continue
        if k in judged_new:
            missing.append((k, idxs))
            continue
        best_rp, score = best_similar(real_part(k), candidates, args.cutoff)
        if best_rp:
            similar.append((k, idxs, cand_map[best_rp], score))
        else:
            missing.append((k, idxs))

    # 百分比分母是参与匹配的行：drop-blank-keys 时空键行不算
    total_rows = len(s_rows) - (blank if args.drop_blank_keys else 0)
    def pct(n):
        return f"{n / total_rows * 100:.1f}%" if total_rows else "-"

    m_rows = sum(len(v) for _, v in matched)
    s_rows_n = sum(len(v) for _, v, _, _ in similar)
    x_rows = sum(len(v) for _, v in missing)

    print(f"\n源文件参与匹配 {total_rows} 行，主键去重后 {len(source_groups)} 个")
    if blank:
        print(f"  主键含空列的行：{blank}（空列以空串参与匹配，需系统侧同为空才配得上）")
    print(f"  可匹配      {len(matched):>5} 个主键 / {m_rows:>5} 行  {pct(m_rows)}")
    print(f"  相似待确认  {len(similar):>5} 个主键 / {s_rows_n:>5} 行  {pct(s_rows_n)}")
    print(f"  完全无      {len(missing):>5} 个主键 / {x_rows:>5} 行  {pct(x_rows)}")

    os.makedirs(args.outdir, exist_ok=True)

    if similar:
        path = os.path.join(args.outdir, f"{args.label}-疑似匹配.xlsx")
        ordered = sorted(similar, key=lambda x: -x[3])
        # 展示时剥掉注入的常量维度（s_cols 与 t_cols 按位对应），客户只看真实差异
        real_key_cols = [c for c in s_cols if c not in const_cols]
        disp_t_cols = [t for s, t in zip(s_cols, t_cols) if s not in const_cols] or t_cols

        def disp_target(best):
            # 疑似匹配列展示关联表完整主键，让人看清对到的是哪一条（含类型维度）
            return KEY_SEP.join(str(target_index[best][0].get(c) or "") for c in t_cols)

        def aux_vals(idxs, best):
            out = []
            for sc, tc in aux_pairs:
                vals = []
                for i in idxs:
                    v = str(s_rows[i].get(sc) or "").strip()
                    if v and v not in vals: vals.append(v)
                out.append("；".join(vals))
                out.append(str(target_index[best][0].get(tc) or ""))
            return out
        aux_header = [h for sc, tc in aux_pairs for h in (f"源文件·{sc}", f"系统·{tc}")]

        # sheet1 结论：一个主键一行，客户在确认结果列作答；两侧都展示完整主键，对照列并排
        summary_rows = [
            [raw_key(s_rows[idxs[0]], s_cols),
             disp_target(best),
             f"{ratio:.2f}", len(idxs)] + aux_vals(idxs, best) + [""]
            for _, idxs, best, ratio in ordered
        ]
        # sheet2 明细：把涉及的源表行原样抽出来，主键字段旁插疑似匹配列供对照
        insert_at = max(s_header.index(c) for c in real_key_cols) if real_key_cols else len(s_header) - 1
        orig_header = [c for c in s_header if c not in const_cols]
        left = [c for c in orig_header if s_header.index(c) <= insert_at]
        right = [c for c in orig_header if s_header.index(c) > insert_at]
        sys_key_label = f"疑似匹配·系统现有（{KEY_SEP.join(t_cols)}）"
        sim_label = f"相似度（按{KEY_SEP.join(disp_t_cols)}）"
        sys_aux_header = [f"系统·{tc}" for _, tc in aux_pairs]
        detail_header = left + [sys_key_label, sim_label] + sys_aux_header + right
        detail_rows = []
        for k, idxs, best, ratio in ordered:
            raw_target = disp_target(best)
            sys_aux = [str(target_index[best][0].get(tc) or "") for _, tc in aux_pairs]
            for i in idxs:
                r = s_rows[i]
                detail_rows.append(
                    [r.get(c) for c in left] + [raw_target, f"{ratio:.2f}"] + sys_aux + [r.get(c) for c in right]
                )
        write_xlsx_sheets(path, [
            ("明细", detail_header, detail_rows, None, None, [sys_key_label]),
            ("结论", ["源文件值", "系统里最接近的", sim_label, "涉及行数"] + aux_header + ["确认结果"], summary_rows, None,
             {"确认结果": ["是同一个", "不是同一个"]}),
        ])
        print(f"\n疑似匹配已写入 {path}（明细 sheet 对照源表行、疑似匹配列标黄；结论 sheet 下拉选确认结果）")

    if missing:
        path = os.path.join(args.outdir, f"{args.label}-待导入.xlsx")
        ordered_missing = sorted(missing, key=lambda x: -len(x[1]))
        real_key_cols = [c for c in s_cols if c not in const_cols]
        orig_header = [c for c in s_header if c not in const_cols]
        # 「导入索引」与关联表补索引字段同名同值；--carry 的源列一并带出，同键多值用；连接
        if args.carry and args.carry.strip() == "all":
            carry_pairs = [(c, c) for c in orig_header if c not in s_cols and c in t_header]
        else:
            carry_pairs = parse_pairs(args.carry)
        carry_cols = [tc for _, tc in carry_pairs]
        def carry_vals(idxs):
            out = []
            for sc, _ in carry_pairs:
                vals = []
                for i in idxs:
                    v = str(s_rows[i].get(sc) or "").strip()
                    if v and v not in vals: vals.append(v)
                out.append("；".join(vals))
            return out
        summary_rows = [
            [raw_key(s_rows[idxs[0]], s_cols).split(KEY_SEP)[i] if i < len(s_cols) else "" for i in range(len(t_cols))]
            + [len(idxs), raw_key(s_rows[idxs[0]], s_cols)] + carry_vals(idxs)
            for _, idxs in ordered_missing
        ]
        detail_rows = [
            [s_rows[i].get(c) for c in orig_header]
            for _, idxs in ordered_missing
            for i in idxs
        ]
        write_xlsx_sheets(path, [
            ("明细", orig_header, detail_rows, None, None, real_key_cols),
            ("结论", t_cols + ["涉及源文件行数", "导入索引"] + carry_cols, summary_rows, None),
        ])
        print(f"待导入已写入 {path}（明细 sheet 对照源表行、主键列标黄；结论 sheet 为待新建清单，末列导入索引与关联表索引字段同名同值），核对后再导入关联表")

    if total_rows and m_rows / total_rows < 0.7:
        print("\n可匹配率低于七成，先回头查索引拼接规则和主键定义，不要直接把清单丢给客户人工对")


if __name__ == "__main__":
    main()
