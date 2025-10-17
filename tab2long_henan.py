# -*- coding: utf-8 -*-
"""
将四类“河南省××表”的宽表转换为严格按模板字段的长表。

特点/修复：
- 自动选引擎（.xlsx -> openpyxl；.xls -> xlrd<2.0>）
- 处理合并单元格：维度列向下填充，只要求主维度非空
- 识别“数值列”的阈值放宽到 **>=0.55**（修复第一数量列被误判为维度列）
- 同时支持全角/半角括号：如 “第一数量（万） / 第一数量(万)”
- 商品量值表：当“第一数量/同比（当月）”在 8月块左侧且列头不写“8月”，
  结合 --dt 推断月份 + 列序位置（与首个“8月”列比较）来判定为“当月”
- 贸易方式：剥离名称前缀编号（写入 col_xuhao），仅名称写入 col_name，稳定排序
- move 映射统一：进出口=0，出口=1，进口=2

用法示例见文末。
"""
import argparse
import re
from pathlib import Path
from typing import List, Optional
from datetime import datetime

import pandas as pd

# ---------- Excel 读取 ----------
def _detect_engine(path: Path) -> str:
    p = str(path).lower()
    if p.endswith((".xlsx", ".xlsm")):
        return "openpyxl"
    if p.endswith(".xls"):
        return "xlrd"
    return "openpyxl"

def _read_excel_any(path: Path, *, sheet_name=None, header=None) -> pd.DataFrame:
    eng = _detect_engine(path)
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header, engine=eng)
    except ImportError as e:
        need = "openpyxl" if eng == "openpyxl" else 'xlrd<2.0'
        raise SystemExit(f"[依赖缺失] {need}\n请安装：pip install -U {need}\n{e}")
    except Exception as e:
        raise SystemExit(f"[读取失败] {path.name}: {e}")

def _excel_file_any(path: Path) -> pd.ExcelFile:
    eng = _detect_engine(path)
    try:
        return pd.ExcelFile(path, engine=eng)
    except ImportError as e:
        need = "openpyxl" if eng == "openpyxl" else 'xlrd<2.0'
        raise SystemExit(f"[依赖缺失] {need}\n请安装：pip install -U {need}\n{e}")
    except Exception as e:
        raise SystemExit(f"[打开失败] {path.name}: {e}")

# ---------- 规范化 / 期间判定 ----------
def _norm_token(s: str) -> str:
    s = (s or "")
    s = re.sub(r"\s+", "", s)
    return s.replace("—","-").replace("–","-").replace("－","-")

def _month_token_from_dt(dt: str) -> str:
    try:
        m = int(datetime.fromisoformat(dt).month)
    except Exception:
        m = int(re.search(r"[-/.](\d{1,2})[-/.]", dt).group(1))
    return f"{m}月"

def _is_cumulative(g: str) -> bool:
    g = _norm_token(g)
    return ("累计" in g) or bool(re.fullmatch(r"1-\d{1,2}月", g))

def _is_single_month(period: str) -> bool:
    p = _norm_token(period)
    if p in ("当月","本月"):
        return True
    if re.fullmatch(r"\d{1,2}月", p):  # 8月/08月
        return True
    return ("月" in p) and ("-" not in p)

def _row_name_from_period_metric(period: str, metric: str) -> str:
    is_single = _is_single_month(period)
    if "同比" in (metric or ""):
        return "当月同比" if is_single else "累计同比"
    return "当月值" if is_single else "累计值"

def _decide_period_general(g1: str, g2: str, month_token: str) -> str:
    """国别/贸易方式：优先累计；其次明确月份；否则回退 month_token。"""
    n1, n2 = _norm_token(g1), _norm_token(g2)
    if _is_cumulative(n1) or _is_cumulative(n2):
        return "1-" + month_token
    if month_token in (n1, n2) or re.fullmatch(rf"0?{re.escape(month_token[:-1])}月", n1) or re.fullmatch(rf"0?{re.escape(month_token[:-1])}月", n2):
        return month_token
    return g2 or g1 or month_token

# ---------- 宽表 -> 通用长表 ----------
def _infer_data_start(df: pd.DataFrame, max_check: int = 8) -> int:
    for i in range(min(max_check, len(df))):
        row = df.iloc[i]
        non_empty = row.notna().sum()
        if non_empty == 0:
            continue
        numeric_like = pd.to_numeric(row, errors="coerce").notna().sum()
        if numeric_like >= max(2, int(0.5 * non_empty)):
            return max(1, i)
    return 1

def _build_colnames_from_header(header_block: pd.DataFrame) -> List[str]:
    hb = header_block.copy()
    hb = hb.apply(lambda s: s.fillna(method="ffill"), axis=1)
    hb = hb.fillna(method="ffill")
    hb = hb.applymap(lambda x: None if (isinstance(x, str) and str(x).strip().startswith("Unnamed")) else x)
    hb = hb.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    names, seen = [], {}
    for col in hb.columns:
        parts = []
        for r in range(hb.shape[0]):
            v = hb.iloc[r, hb.columns.get_loc(col)]
            if v is not None and v != "":
                parts.append(str(v))
        if not parts:
            parts = [f"col_{col}"]
        name = " | ".join(parts)
        cnt = seen.get(name, 0)
        if cnt:
            name = f"{name} ({cnt+1})"
        seen[name] = cnt + 1
        names.append(name)
    return names

def _parse_wide_to_generic_long(path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if not sheet_name:
        xls = _excel_file_any(path)
        target_sheet = xls.sheet_names[0]
    else:
        target_sheet = sheet_name

    df_raw = _read_excel_any(path, sheet_name=target_sheet, header=None)
    if isinstance(df_raw, dict):
        df_raw = next(iter(df_raw.values()))
    df_raw = df_raw.dropna(axis=1, how="all")

    data_start = _infer_data_start(df_raw)
    header_end = max(1, min(3, data_start))
    header_block = df_raw.iloc[:header_end, :]
    data_block   = df_raw.iloc[header_end:, :].reset_index(drop=True)

    colnames = _build_colnames_from_header(header_block)
    data_block.columns = colnames

    # —— 识别维度列：数值占比阈值 >= 0.55（关键修复）——
    def is_numeric_series(s: pd.Series) -> bool:
        return pd.to_numeric(s, errors="coerce").notna().mean() >= 0.55

    id_cols = []
    for c in data_block.columns:
        if not is_numeric_series(data_block[c]):
            id_cols.append(c)
        else:
            break
    if not id_cols:
        id_cols = [data_block.columns[0]]

    # 合并单元格向下填充，仅要求主维度非空
    df = data_block.copy()
    for c in id_cols:
        df[c] = df[c].replace("", pd.NA).ffill()
    df = df[~df[id_cols[0]].isna()].copy()

    # 删除全空数据行
    num_part = df.drop(columns=id_cols)
    if not num_part.empty:
        df = df[~num_part.isna().all(axis=1)].copy()

    # 记录数值列在原表中的序号（用于位置判定）
    value_cols = [c for c in df.columns if c not in id_cols]
    pos_map = {c: i for i, c in enumerate(value_cols)}

    long = df.melt(id_vars=id_cols, value_vars=value_cols, var_name="metric_full", value_name="value")
    long["__col_pos"] = long["metric_full"].map(pos_map)

    parts = long["metric_full"].str.split(r"\s*\|\s*", expand=True)
    long["level1"] = parts[0].fillna("") if parts.shape[1] > 0 else ""
    long["level2"] = parts[1].fillna("") if parts.shape[1] > 1 else ""
    long["level3"] = parts[2].fillna("") if parts.shape[1] > 2 else ""

    # 提取单位：支持全角/半角括号
    def extract_unit(s: str):
        if not isinstance(s, str):
            return None
        m = re.search(r"（([^）]+)）", s)
        if m:
            return m.group(1)
        m2 = re.search(r"\(([^)]+)\)", s)
        return m2.group(1) if m2 else None

    def strip_unit(s: str) -> str:
        s = re.sub(r"（[^）]+）", "", s or "")
        s = re.sub(r"\([^)]*\)", "", s)
        return s.strip()

    deepest = long["level3"].where(long["level3"] != "", long["level2"])
    long["unit"] = deepest.apply(extract_unit)          # 万元/%/万 等
    long["metric"] = deepest.apply(lambda x: strip_unit(x) if isinstance(x, str) else "")

    rename_map = {id_cols[0]: "dimension"}
    if len(id_cols) > 1:
        rename_map[id_cols[1]] = "id1"                 # 第一计量单位
    long = long.rename(columns=rename_map)

    cols = ["dimension"] + (["id1"] if "id1" in long.columns else []) + \
           ["level1","level2","metric","unit","value","__col_pos"]
    long = long[cols].rename(columns={"level1":"group1","level2":"group2"})

    for c in long.columns:
        if long[c].dtype == object:
            long[c] = long[c].astype(str).str.strip()
    long["value"] = pd.to_numeric(long["value"], errors="coerce")
    return long.reset_index(drop=True)

# ---------- 通用长表 -> 模板 ----------
MOVE_ID_MAP = {"进出口": 0, "出口": 1, "进口": 2}
ROW_NAME_ORDER = {"当月值": 1, "当月同比": 2, "累计值": 3, "累计同比": 4}
COL_TYPE_ORDER = {"第一数量": 1, "金额": 2}

def _normalize_scope(s: str) -> str:
    s = (s or "").strip()
    if "进出口" in s: return "进出口"
    if "出口" in s and "进出口" not in s: return "出口"
    if "进口" in s and "进出口" not in s: return "进口"
    return s

def _split_trade_code_and_name(text: str):
    """把 '1 总值' / '03 加工贸易' / '1.1 海关特殊监管区' 等拆成：(主编号, 完整编号, 名称)"""
    if text is None:
        return (None, None, "")
    s = str(text).strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)*)\s*[:：、\.\s]*\s*(.+?)\s*$", s)
    if m:
        full_code = m.group(1)
        try:
            major = int(full_code.split('.')[0])
        except Exception:
            major = None
        return (major, full_code, m.group(2))
    m2 = re.match(r"^\s*(\d+)\s*(.+?)\s*$", s)
    if m2:
        try:
            major = int(m2.group(1))
        except Exception:
            major = None
        return (major, str(major) if major is not None else None, m2.group(2))
    return (None, None, s)

def generic_to_country_or_trade(df_long: pd.DataFrame, dt: str) -> pd.DataFrame:
    month_token = _month_token_from_dt(dt)
    rows = []
    unique = list(pd.unique(df_long["dimension"]))
    name2idx = {n: i + 1 for i, n in enumerate(unique)}

    for _, r in df_long.iterrows():
        scope = _normalize_scope(r["group1"])
        period = _decide_period_general(str(r.get("group1","")), str(r.get("group2","")), month_token)
        row_name = _row_name_from_period_metric(period, r["metric"])
        rows.append({
            "dt": dt,
            "move_id": MOVE_ID_MAP.get(scope, None),
            "move_nm": scope,
            "col_xuhao": name2idx.get(r["dimension"], None),
            "col_name": r["dimension"],
            "row_xuhao": ROW_NAME_ORDER.get(row_name, None),
            "row_name": row_name,
            "val": r["value"],
            "dw": ("万元" if "金额" in r["metric"] else "%"),
        })
    out = pd.DataFrame(rows)
    return out[["dt","move_id","move_nm","col_xuhao","col_name","row_xuhao","row_name","val","dw"]]

def generic_to_trade(df_long: pd.DataFrame, dt: str) -> pd.DataFrame:
    """贸易方式：剥离编号到 col_xuhao，名称写入 col_name；稳定排序。"""
    month_token = _month_token_from_dt(dt)
    parsed = df_long["dimension"].apply(_split_trade_code_and_name)
    major = parsed.apply(lambda t: t[0])
    clean = parsed.apply(lambda t: t[2]).astype(str).str.strip()

    df2 = df_long.copy()
    df2["__clean_name"] = clean
    df2["__major"] = major

    # 按出现顺序补号
    seen, order_idx = {}, []
    for nm in df2["__clean_name"]:
        if nm not in seen:
            seen[nm] = len(seen) + 1
        order_idx.append(seen[nm])
    df2["__order_name"] = order_idx
    df2["__col_xuhao"] = df2["__major"].fillna(df2["__order_name"])

    rows = []
    for _, r in df2.iterrows():
        scope = _normalize_scope(r["group1"])
        period = _decide_period_general(str(r.get("group1","")), str(r.get("group2","")), month_token)
        row_name = _row_name_from_period_metric(period, r["metric"])
        rows.append({
            "dt": dt,
            "move_id": MOVE_ID_MAP.get(scope, None),
            "move_nm": scope,
            "col_xuhao": int(r["__col_xuhao"]) if pd.notna(r["__col_xuhao"]) else None,
            "col_name": r["__clean_name"],
            "row_xuhao": ROW_NAME_ORDER.get(row_name, None),
            "row_name": row_name,
            "val": r["value"],
            "dw": ("万元" if "金额" in r["metric"] else "%"),
        })
    out = pd.DataFrame(rows)
    out = out[["dt","move_id","move_nm","col_xuhao","col_name","row_xuhao","row_name","val","dw"]]
    return out.sort_values(["dt","move_id","col_xuhao","row_xuhao"], kind="stable").reset_index(drop=True)

def generic_to_goods(df_long: pd.DataFrame, dt: str, move_nm: str) -> pd.DataFrame:
    month_token = _month_token_from_dt(dt)

    # 定位首个“8月”列的位置（用 group1）
    norm_g1 = df_long["group1"].map(_norm_token)
    month_pos_candidates = df_long.loc[norm_g1 == _norm_token(month_token), "__col_pos"]
    first_month_pos = int(month_pos_candidates.min()) if not month_pos_candidates.empty else None

    rows = []
    uniq_goods = list(pd.unique(df_long["dimension"]))
    goods2idx = {n: i + 1 for i, n in enumerate(uniq_goods)}

    for _, r in df_long.iterrows():
        col_name   = r["dimension"]
        first_unit = r["id1"] if "id1" in r.index else None
        metric     = r["metric"]
        g1, g2     = str(r.get("group1","")), str(r.get("group2",""))
        pos        = r.get("__col_pos")

        # 文本+位置双重判定 period
        n1, n2 = _norm_token(g1), _norm_token(g2)
        if _is_cumulative(n1) or _is_cumulative(n2) or n2.startswith("1-"):
            period = "1-" + month_token
        elif _is_single_month(g2):
            period = g2
        elif "第一数量" in (metric or ""):
            # 列头未写“8月”，但在“8月块”左侧 => 当月
            if first_month_pos is not None and pd.notna(pos) and int(pos) < first_month_pos:
                period = month_token
            else:
                period = month_token
        else:
            period = month_token

        row_name = _row_name_from_period_metric(period, metric)

        # 列类型 + 单位
        if "同比" in (metric or ""):
            col_type = "金额" if "金额" in metric else ("第一数量" if "第一数量" in metric else None)
            dw = "%"
        else:
            if "金额" in (metric or ""):
                col_type, dw = "金额", "万元"
            elif "第一数量" in (metric or ""):
                col_type = "第一数量"
                metric_unit = r.get("unit")
                if first_unit and str(first_unit).strip() not in ("","NaN","nan"):
                    dw = str(first_unit).strip()
                elif metric_unit and metric_unit not in ("万元","%"):
                    dw = str(metric_unit).strip()  # 如 “万”
                else:
                    dw = ""
            else:
                col_type, dw = None, ""

        rows.append({
            "dt": dt,
            "move_xuhao": MOVE_ID_MAP.get(move_nm, None),
            "move_nm": move_nm,
            "col_xuhao": goods2idx.get(col_name, None),
            "col_name": col_name,
            "col_type_xuhao": COL_TYPE_ORDER.get(col_type, None),
            "col_type": col_type,
            "row_xuhao": ROW_NAME_ORDER.get(row_name, None),
            "row_name": row_name,
            "val": r["value"],
            "dw": dw,
        })

    out = pd.DataFrame(rows)
    return out[["dt","move_xuhao","move_nm","col_xuhao","col_name","col_type_xuhao","col_type","row_xuhao","row_name","val","dw"]]

# ---------- 直接导出现有模板长表（可选） ----------
def export_existing_template_sheet(xlsx: Path, sheet: str, out_csv: Path):
    df = _read_excel_any(xlsx, sheet_name=sheet, header=0)
    df.to_csv(out_csv, index=False, encoding="utf-8-sig")

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="宽表 -> 模板长表（河南省四类统计表）")
    ap.add_argument("file", help="输入 Excel 路径（宽表或已有模板长表）")
    ap.add_argument("-o","--out", required=True, help="输出 CSV 路径")
    ap.add_argument("--dt", help="月份日期 YYYY-MM-DD（如 2025-08-01）")
    ap.add_argument("--type", choices=["country","trade","goods_export","goods_import"],
                    help="宽表转换时指定表类型")
    ap.add_argument("--sheet", help="宽表所在工作表名（不填默认第一表）")
    ap.add_argument("--from-long-sheet", help="若已有模板长表工作表，直接导出其为 CSV（如 long / merged_long / merged_goods）")
    args = ap.parse_args()

    xlsx = Path(args.file)
    out_csv = Path(args.out)

    if args.from_long_sheet:
        export_existing_template_sheet(xlsx, args.from_long_sheet, out_csv)
        print(f"[OK] Exported: {xlsx.name}::{args.from_long_sheet} -> {out_csv}")
        return

    if not args.type:
        raise SystemExit("从宽表转换时必须提供 --type（country|trade|goods_export|goods_import）")
    if not args.dt:
        raise SystemExit("请提供 --dt（如 2025-08-01），用于写入 dt 与识别“当月”。")

    generic = _parse_wide_to_generic_long(xlsx, sheet_name=args.sheet)

    if args.type == "country":
        final_df = generic_to_country_or_trade(generic, dt=args.dt)
    elif args.type == "trade":
        final_df = generic_to_trade(generic, dt=args.dt)
    elif args.type == "goods_export":
        final_df = generic_to_goods(generic, dt=args.dt, move_nm="出口")
    elif args.type == "goods_import":
        final_df = generic_to_goods(generic, dt=args.dt, move_nm="进口")
    else:
        raise SystemExit("不支持的 --type")

    final_df.to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"[OK] {xlsx.name} ({args.type}) -> {out_csv}  rows={len(final_df)}")

if __name__ == "__main__":
    main()
