"""源文件读取与文本归一化，供 get_source_profile.py 和 render_match_report.py 共用。"""
import csv
import re
import unicodedata

KEY_SEP = "｜"


def normalize(value):
    """归一化到可比对的形态。两侧必须用同一套规则，否则匹配率会假性偏低。"""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    # 全角转半角，顺带统一中文括号
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("（", "(").replace("）", ")")
    # 各类空白压成无
    s = re.sub(r"\s+", "", s)
    return s.lower()


# ---- 相似判定：与业务无关的口径，只看主键各段的文本形态 ----
_PAREN = re.compile(r"\([^()]*\)")
_NUM = re.compile(r"\d+\.?\d*")


def seg_core(value):
    """去掉括号里的修饰语（无蔗糖、一致性评价、别名）后的主体，用于名称段比对。"""
    s = normalize(value)
    s = s.replace("×", "*").replace("x", "*").replace("—", "-").replace("–", "-")
    return _PAREN.sub("", s).strip("-* ")


def seg_nums(value):
    return _NUM.findall(normalize(value))


def nums_equivalent(a, b):
    """数字集合相同，或首个数字相同且其余数字乘积相同（7片*4板 ≡ 28s）。"""
    na, nb = seg_nums(a), seg_nums(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    if na[0] != nb[0]:
        return False
    def prod(xs):
        r = 1.0
        for x in xs:
            r *= float(x)
        return r
    return len(na) > 1 and len(nb) > 1 and abs(prod(na[1:]) - prod(nb[1:])) < 1e-9


def seg_close(a, b, cutoff):
    """一段主键是否算相近。返回 (等级, 分数)：2=等价，1=相似，0=不像。"""
    from difflib import SequenceMatcher
    na, nb = normalize(a), normalize(b)
    if na == nb:
        return 2, 1.0
    ca, cb = seg_core(a), seg_core(b)
    if ca and ca == cb:
        return 2, 0.98
    if nums_equivalent(a, b):
        return 2, 0.95
    if ca and cb and (ca in cb or cb in ca) and min(len(ca), len(cb)) >= 4:
        return 1, 0.9
    r = SequenceMatcher(None, ca or na, cb or nb).ratio()
    return (1, r) if r >= cutoff else (0, r)


def keys_similar(src_parts, tgt_parts, cutoff=0.75):
    """两侧主键按段比对。疑似的条件：每一段都相近，多段时还要至少一段等价。返回 (是否疑似, 综合分)。"""
    if len(src_parts) != len(tgt_parts):
        return False, 0.0
    levels, scores = [], []
    for a, b in zip(src_parts, tgt_parts):
        lv, sc = seg_close(a, b, cutoff)
        if lv == 0:
            return False, sc
        levels.append(lv); scores.append(sc)
    # 多段主键要求至少一段等价（同名不同规格、同规格错字都算）；单段主键只有相似度可依
    if len(levels) > 1 and 2 not in levels:
        return False, sum(scores) / len(scores)
    return True, sum(scores) / len(scores)


def build_key(row, columns, allow_blank=True):
    """按列顺序拼接主键。

    allow_blank=True：空值用空串占位继续拼接。判重时用这个，因为两行同样空着就是真重复。
    allow_blank=False：任一列为空即返回 None，代表这行无法参与关联匹配。
    """
    parts = []
    for col in columns:
        v = normalize(row.get(col))
        if not v and not allow_blank:
            return None
        parts.append(v)
    return KEY_SEP.join(parts)


def has_blank(row, columns):
    return any(not normalize(row.get(c)) for c in columns)


def raw_key(row, columns):
    """未归一化的主键，用于展示给人看。"""
    return KEY_SEP.join("" if row.get(c) is None else str(row.get(c)).strip() for c in columns)


def read_rows(path, sheet=1, title_row=1, encoding="utf-8"):
    """读 xlsx/xls/csv，返回 (表头列表, 行字典列表)。"""
    lower = path.lower()
    if lower.endswith(".csv"):
        return _read_csv(path, title_row, encoding)
    if lower.endswith(".xls"):
        return _read_xls(path, sheet, title_row)
    return _read_excel(path, sheet, title_row)


def _read_xls(path, sheet, title_row):
    try:
        import xlrd
    except ImportError:
        raise SystemExit("读 .xls 需要 xlrd：pip3 install xlrd")
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_index(sheet - 1)

    def cell_value(c):
        # xls 数值统一是 float，整数值去掉 .0；日期转 ISO 格式
        if c.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate_as_tuple(c.value, wb.datemode)
            return f"{dt[0]:04d}-{dt[1]:02d}-{dt[2]:02d}" if dt[0] else c.value
        if c.ctype == xlrd.XL_CELL_NUMBER and c.value == int(c.value):
            return int(c.value)
        if c.ctype == xlrd.XL_CELL_EMPTY:
            return None
        return c.value

    if ws.nrows < title_row:
        raise SystemExit(f"第 {title_row} 行没有表头")
    header = [str(cell_value(c)).strip() if cell_value(c) is not None else f"列{i+1}"
              for i, c in enumerate(ws.row(title_row - 1))]
    data = []
    for ridx in range(title_row, ws.nrows):
        values = [cell_value(c) for c in ws.row(ridx)]
        if not any(v is not None and str(v).strip() for v in values):
            continue
        data.append({header[i]: (values[i] if i < len(values) else None) for i in range(len(header))})
    return header, data


def _read_csv(path, title_row, encoding):
    with open(path, newline="", encoding=encoding) as f:
        rows = list(csv.reader(f))
    if len(rows) < title_row:
        raise SystemExit(f"文件不足 {title_row} 行，无法取表头")
    header = [str(c).strip() for c in rows[title_row - 1]]
    data = []
    for r in rows[title_row:]:
        if not any(str(c).strip() for c in r):
            continue
        data.append({header[i]: (r[i] if i < len(r) else None) for i in range(len(header))})
    return header, data


def _read_excel(path, sheet, title_row):
    try:
        import openpyxl
    except ImportError:
        raise SystemExit("需要 openpyxl：pip3 install openpyxl")
    import warnings

    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[sheet - 1]
    header = None
    data = []
    for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if idx < title_row:
            continue
        if idx == title_row:
            header = [str(c).strip() if c is not None else f"列{i+1}" for i, c in enumerate(row)]
            continue
        if not any(c is not None and str(c).strip() for c in row):
            continue
        data.append({header[i]: (row[i] if i < len(row) else None) for i in range(len(header))})
    wb.close()
    if header is None:
        raise SystemExit(f"第 {title_row} 行没有表头")
    return header, data


def check_columns(header, columns, label):
    missing = [c for c in columns if c not in header]
    if missing:
        raise SystemExit(f"{label}里没有这些列：{', '.join(missing)}\n实际列名：{', '.join(header)}")


def split_cols(text):
    return [c.strip() for c in text.split(",") if c.strip()]


def write_xlsx(path, header, rows, highlight=None, sheet_name=None):
    """写单 sheet xlsx 交付文件。highlight 是 0-based 数据行号集合，整行标黄。"""
    write_xlsx_sheets(path, [(sheet_name or "Sheet1", header, rows, highlight)])


def write_xlsx_sheets(path, sheets):
    """写多 sheet xlsx 交付文件。

    sheets 每项是 (sheet名, header, rows, highlight)，可再加：
    第 5 元素 {列名: [选项]}：给该列挂数据验证下拉，选而不是填；
    第 6 元素 [列名]：整列标黄（含表头），用于突出要看的列。
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        from openpyxl.worksheet.datavalidation import DataValidation
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise SystemExit("需要 openpyxl：pip3 install openpyxl")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    yellow = PatternFill("solid", start_color="FFFF00")
    for entry in sheets:
        name, header, rows, highlight = entry[:4]
        dropdowns = entry[4] if len(entry) > 4 else None
        highlight_cols = entry[5] if len(entry) > 5 else None
        ws = wb.create_sheet(name)
        ws.append(list(header))
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for i, r in enumerate(rows):
            ws.append(list(r))
            if highlight and i in highlight:
                for cell in ws[ws.max_row]:
                    cell.fill = yellow
        # 列宽按表头和前 50 行内容粗调，交付文件打开即可读
        for col_idx in range(1, len(header) + 1):
            width = max(
                len(str(ws.cell(row=r, column=col_idx).value or ""))
                for r in range(1, min(ws.max_row, 51) + 1)
            )
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max(width * 1.8, 10), 50)
        if highlight_cols:
            for col_name in highlight_cols:
                if col_name not in header:
                    continue
                col_idx = header.index(col_name) + 1
                for r in range(1, len(rows) + 2):
                    ws.cell(row=r, column=col_idx).fill = yellow
        if dropdowns and len(rows):
            for col_name, options in dropdowns.items():
                if col_name not in header:
                    continue
                letter = get_column_letter(header.index(col_name) + 1)
                dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
                dv.error = "请从下拉选项中选择"
                dv.errorTitle = "无效输入"
                ws.add_data_validation(dv)
                dv.add(f"{letter}2:{letter}{len(rows) + 1}")
    wb.save(path)
