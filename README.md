# Order Reconciliation Skill

**Live case study:** https://asherlatchman-hub.github.io/order-reconciliation-skill/

A custom **Claude Skill** that reconciles a seller order export against a warehouse/fulfilment export and produces an Excel **exception report** for an operations team to review.

Upload two spreadsheets and Claude runs the same checks every time, without the process being explained again:

- ✅ Matches orders by order ID + SKU
- ✅ Finds orders that were **sold but never shipped**
- ✅ Finds items **shipped with no matching sale**
- ✅ Flags **quantity discrepancies** (short- or over-shipped)
- ✅ Detects **duplicate lines** within each file
- ✅ Reports **data-quality problems** (blank IDs, non-numeric quantities)
- ✅ Generates a formatted, filterable Excel report

## Test results

The skill was tested against synthetic data with errors planted on purpose. It caught every one.

| Problem planted | Planted | Found | Report tab |
|---|---:|---:|---|
| Orders sold but never shipped | 3 orders (4 lines) | 4 lines | Missing from Warehouse |
| Shipments with no matching sale | 2 | 2 | Missing from Seller |
| Quantities changed in the warehouse file | 4 | 4 | Quantity Discrepancies |
| Rows exported twice | 3 rows | 6 lines | Duplicates (+3 linked notes) |
| Quantity typed as text ("two") | 1 | 1 | Data Quality |
| Blank order ID | 1 | 1 | Data Quality |
| Order IDs typed in lowercase | 14 rows | all matched | Normalised, not flagged |

**80 order lines compared · 66 matched cleanly · 82.5% clean match rate.**
See [`samples/exception_report.xlsx`](samples/exception_report.xlsx) for the full output.

## Repository layout

```
skill/order-reconciliation/
├── SKILL.md              # Instructions Claude follows: when to use it, workflow, how to report
└── scripts/
    └── reconcile.py      # Deterministic matching + Excel report generation
samples/
├── seller_orders.xlsx    # Synthetic seller export (Order ID, SKU, Quantity…)
├── warehouse_orders.csv  # Synthetic warehouse export (order_no, product_code, qty_shipped…)
└── exception_report.xlsx # Report produced by the skill
tools/
└── make_samples.py       # Regenerates the test files with planted errors
```

## Design decisions

- **Logic in code, judgement in the AI.** The matching runs in a Python script, so the same inputs always give the same output. Claude handles the parts that need judgement: identifying which file is which, coping with unusual column names, sanity-checking low match rates and explaining the results.
- **Match on order ID + SKU, not ID alone.** One order can contain several products. Matching on ID alone would hide a missing item.
- **Sum duplicates before comparing quantities.** A repeated row isn't counted as both a duplicate and a separate quantity problem; the two are linked with a note.
- **Unreadable rows go to Data Quality.** A quantity typed as "two" is a data-entry fault, not a missing order.

## Using it

**As a Claude Skill:** add the `skill/order-reconciliation` folder as a skill in Claude, then upload two order files and ask Claude to reconcile them.

**From the command line:**

```bash
pip install pandas openpyxl
python skill/order-reconciliation/scripts/reconcile.py \
    samples/seller_orders.xlsx samples/warehouse_orders.csv -o exception_report.xlsx
```

Column names are detected automatically; override them with `--order-col`, `--sku-col` and `--qty-col` if needed.

## Tech

Python · pandas · openpyxl · Claude Skills

---
*All sample data is synthetic and was generated for testing. No real customer or company data is included.*
