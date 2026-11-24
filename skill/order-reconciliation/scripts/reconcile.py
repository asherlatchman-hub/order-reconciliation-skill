#!/usr/bin/env python3
"""
Order Reconciliation: compare a seller order export against a warehouse export
and produce an exception report for the operations team.

Usage:
    python reconcile.py SELLER_FILE WAREHOUSE_FILE -o report.xlsx
        [--order-col NAME] [--sku-col NAME] [--qty-col NAME]

Accepts .csv, .xlsx or .xls. Column names are detected automatically
(e.g. "Order ID", "order_no", "Ref"); use the --*-col flags to override.
Prints a JSON summary to stdout so the results can be explained in chat.
"""
import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ORDER_ALIASES = ["order id", "order_id", "orderid", "order no", "order number",
                 "order_no", "order #", "order", "order ref", "reference", "ref"]
SKU_ALIASES = ["sku", "product code", "product_code", "item code", "item_code",
               "item", "product id", "product_id", "article"]
QTY_ALIASES = ["quantity", "qty", "units", "qty ordered", "qty shipped",
               "quantity shipped", "quantity ordered", "count"]


def norm(s):
    return re.sub(r"[^a-z0-9#]+", " ", str(s).lower()).strip()


def find_col(df, aliases, override, label, required=True):
    if override:
        if override not in df.columns:
            sys.exit(f"Column '{override}' not found. Columns are: {list(df.columns)}")
        return override
    lookup = {norm(c): c for c in df.columns}
    for a in aliases:
        if norm(a) in lookup:
            return lookup[norm(a)]
    for a in aliases:  # fallback: partial match
        for n, c in lookup.items():
            if norm(a) in n:
                return c
    if required:
        sys.exit(f"Could not find a {label} column. Columns are: {list(df.columns)}. "
                 f"Pass it explicitly with --{label}-col.")
    return None


def load(path):
    p = Path(path)
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p, dtype=str)
    return pd.read_excel(p, dtype=str)


def prepare(df, args, source):
    oc = find_col(df, ORDER_ALIASES, args.order_col, "order")
    sc = find_col(df, SKU_ALIASES, args.sku_col, "sku", required=False)
    qc = find_col(df, QTY_ALIASES, args.qty_col, "qty")
    out = pd.DataFrame({
        "order_id": df[oc].fillna("").astype(str).str.strip().str.upper(),
        "sku": df[sc].fillna("").astype(str).str.strip().str.upper() if sc else "",
        "qty_raw": df[qc],
        "source_row": df.index + 2,  # +2 = header row + 1-based rows
    })
    out["qty"] = pd.to_numeric(out["qty_raw"], errors="coerce")
    blank = out[out["order_id"] == ""]
    bad_qty = out[out["qty"].isna() & (out["order_id"] != "")]
    out = out[(out["order_id"] != "") & out["qty"].notna()].copy()
    cols = {"order": oc, "sku": sc, "qty": qc}
    return out, blank, bad_qty, cols


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("seller")
    ap.add_argument("warehouse")
    ap.add_argument("-o", "--output", default="exception_report.xlsx")
    ap.add_argument("--order-col")
    ap.add_argument("--sku-col")
    ap.add_argument("--qty-col")
    args = ap.parse_args()

    s_raw, w_raw = load(args.seller), load(args.warehouse)
    s, s_blank, s_badq, s_cols = prepare(s_raw, args, "seller")
    w, w_blank, w_badq, w_cols = prepare(w_raw, args, "warehouse")

    use_sku = bool(s_cols["sku"] and w_cols["sku"])
    key = ["order_id", "sku"] if use_sku else ["order_id"]

    # 1. Duplicates: same key appearing more than once within one file
    def dupes(df, name):
        d = df[df.duplicated(key, keep=False)].copy()
        d.insert(0, "file", name)
        return d
    duplicates = pd.concat([dupes(s, "Seller"), dupes(w, "Warehouse")])
    duplicates = duplicates.sort_values(["file"] + key + ["source_row"]).drop(columns=["qty_raw"])

    # Aggregate per key so duplicates don't also show as quantity problems twice
    s_agg = s.groupby(key, as_index=False)["qty"].sum().rename(columns={"qty": "seller_qty"})
    w_agg = w.groupby(key, as_index=False)["qty"].sum().rename(columns={"qty": "warehouse_qty"})
    merged = s_agg.merge(w_agg, on=key, how="outer", indicator=True)

    # 2. Missing orders
    # (a line whose quantity was unreadable is reported under Data Quality, not as missing)
    keyset = lambda df: set(map(tuple, df[key].values.tolist()))
    w_bad_keys, s_bad_keys = keyset(w_badq), keyset(s_badq)
    left = merged[merged["_merge"] == "left_only"]
    right = merged[merged["_merge"] == "right_only"]
    missing_wh = left[[tuple(r) not in w_bad_keys for r in left[key].values.tolist()]][key + ["seller_qty"]]
    missing_seller = right[[tuple(r) not in s_bad_keys for r in right[key].values.tolist()]][key + ["warehouse_qty"]]

    # 3. Quantity discrepancies
    both = merged[merged["_merge"] == "both"].copy()
    qty = both[both["seller_qty"] != both["warehouse_qty"]].copy()
    qty["difference"] = qty["warehouse_qty"] - qty["seller_qty"]
    qty["issue"] = qty["difference"].apply(lambda d: "Short-shipped" if d < 0 else "Over-shipped")
    dup_keys = keyset(duplicates) if len(duplicates) else set()
    qty["note"] = ["Check Duplicates tab: a repeated line may explain this" if tuple(r) in dup_keys else ""
                   for r in qty[key].values.tolist()]
    qty = qty[key + ["seller_qty", "warehouse_qty", "difference", "issue", "note"]]

    # 4. Data quality
    dq = []
    for name, blank, badq in [("Seller", s_blank, s_badq), ("Warehouse", w_blank, w_badq)]:
        for _, r in blank.iterrows():
            dq.append({"file": name, "source_row": r["source_row"], "problem": "Blank order ID"})
        for _, r in badq.iterrows():
            dq.append({"file": name, "source_row": r["source_row"],
                       "problem": f"Quantity not a number: '{r['qty_raw']}'",
                       "order_id": r["order_id"]})
    data_quality = pd.DataFrame(dq, columns=["file", "source_row", "order_id", "problem"])

    matched = int((both["seller_qty"] == both["warehouse_qty"]).sum())
    summary = {
        "seller_lines": int(len(s_raw)),
        "warehouse_lines": int(len(w_raw)),
        "matched_by": " + ".join(key),
        "matched_cleanly": matched,
        "missing_from_warehouse": int(len(missing_wh)),
        "missing_from_seller": int(len(missing_seller)),
        "duplicate_lines": int(len(duplicates)),
        "quantity_discrepancies": int(len(qty)),
        "net_unit_difference": float(qty["difference"].sum()) if len(qty) else 0.0,
        "data_quality_issues": int(len(data_quality)),
        "columns_used": {"seller": s_cols, "warehouse": w_cols},
    }
    total_keys = len(merged)
    summary["match_rate_pct"] = round(100 * matched / total_keys, 1) if total_keys else 100.0

    write_report(args.output, summary, missing_wh, missing_seller, duplicates, qty, data_quality,
                 Path(args.seller).name, Path(args.warehouse).name)
    print(json.dumps(summary, indent=2))


HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14)


def style_sheet(ws, df_cols):
    for i, _ in enumerate(df_cols, 1):
        c = ws.cell(row=1, column=i)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col) + 3
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width, 12), 45)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def write_report(path, summary, missing_wh, missing_seller, duplicates, qty, dq, sname, wname):
    nice = lambda df: df.rename(columns=lambda c: c.replace("_", " ").title())
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame().to_excel(xw, sheet_name="Summary", index=False)
        ws = xw.sheets["Summary"]
        ws["A1"] = "Order Reconciliation: Exception Report"
        ws["A1"].font = TITLE_FONT
        ws["A2"] = f"Generated {datetime.now():%d %b %Y %H:%M}  |  Seller file: {sname}  |  Warehouse file: {wname}"
        rows = [
            ("Seller order lines", summary["seller_lines"]),
            ("Warehouse order lines", summary["warehouse_lines"]),
            ("Matched on", summary["matched_by"]),
            ("Orders matched cleanly", summary["matched_cleanly"]),
            ("Match rate", f'{summary["match_rate_pct"]}%'),
            ("", ""),
            ("EXCEPTIONS", "Count"),
            ("Missing from warehouse (sold, not shipped)", summary["missing_from_warehouse"]),
            ("Missing from seller (shipped, no sale)", summary["missing_from_seller"]),
            ("Duplicate lines", summary["duplicate_lines"]),
            ("Quantity discrepancies", summary["quantity_discrepancies"]),
            ("Net unit difference (warehouse − seller)", summary["net_unit_difference"]),
            ("Data quality issues", summary["data_quality_issues"]),
        ]
        for i, (k, v) in enumerate(rows, start=4):
            ws.cell(row=i, column=1, value=k)
            ws.cell(row=i, column=2, value=v)
        for c in (ws.cell(row=10, column=1), ws.cell(row=10, column=2)):
            c.fill, c.font = HEADER_FILL, HEADER_FONT
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["B"].width = 18

        sheets = [("Missing from Warehouse", missing_wh), ("Missing from Seller", missing_seller),
                  ("Quantity Discrepancies", qty), ("Duplicates", duplicates),
                  ("Data Quality", dq)]
        for name, df in sheets:
            out = nice(df) if len(df) else pd.DataFrame({"Result": ["No exceptions found"]})
            out.to_excel(xw, sheet_name=name, index=False)
            style_sheet(xw.sheets[name], out.columns)


if __name__ == "__main__":
    main()
