---
name: order-reconciliation
description: Reconcile a seller order export against a warehouse/fulfilment export and produce an exception report of missing orders, duplicates, quantity mismatches and bad data.
---

# Order Reconciliation

Compares two order spreadsheets, what the seller **sold** and what the warehouse **shipped**, and produces an Excel exception report the operations team can work through.

## When to use
Use it when the user uploads two order files (CSV or Excel) and wants them matched, checked or reconciled. Typical phrases: "reconcile these orders", "what's missing from the warehouse", "check seller vs fulfilment", "find order discrepancies".

## Workflow

1. **Identify the two files.** Work out which is the seller/sales export and which is the warehouse/fulfilment/dispatch export. Use the file names first, then the column names (e.g. `qty_shipped` or `carrier` means warehouse). If it is still unclear, ask the user once.
2. **Inspect the columns.** Read the headers of both files. The script auto-detects the order ID, SKU and quantity columns (for example `Order ID` / `order_no` / `Ref`, `SKU` / `product_code`, `Quantity` / `qty_shipped`). If a file uses unusual names, pass them with `--order-col`, `--sku-col`, `--qty-col`.
3. **Run the bundled script:**
   ```bash
   python3 scripts/reconcile.py SELLER_FILE WAREHOUSE_FILE -o exception_report.xlsx
   ```
   The script prints a JSON summary. If it exits with "Could not find a … column", rerun with the matching `--*-col` flag.
4. **Check the result.** Confirm `columns_used` in the JSON names the right columns. If the match rate is very low (under 50%), suspect a wrong column or an ID format difference (prefixes, leading zeros) and look at a few IDs from each file before reporting.
5. **Deliver.** Send `exception_report.xlsx` to the user. In chat, give a short summary: the match rate, then each exception count, starting with the issue that costs the most money (usually sold but not shipped). Name 2 or 3 specific order IDs that need attention first.

## What the report contains
| Tab | Meaning | Likely action |
|---|---|---|
| Summary | Line counts, match rate, exception counts | Overview for a manager |
| Missing from Warehouse | Sold, never shipped | Chase fulfilment. The customer is waiting |
| Missing from Seller | Shipped with no matching sale | Check for stock loss, a test order or a wrong ID |
| Quantity Discrepancies | Both sides have the order but the units differ (short- or over-shipped) | Correct the shipment or issue a refund |
| Duplicates | Same order + SKU repeated inside one file | Remove the duplicate or check for a double-ship |
| Data Quality | Blank order IDs, quantities that aren't numbers | Fix at source |

## Matching rules
- Order IDs and SKUs are trimmed and upper-cased before matching, so `gn-10401` matches `GN-10401`.
- Rows match on order ID + SKU when both files have a SKU column, otherwise on order ID alone.
- Duplicate lines are summed before quantities are compared. A quantity mismatch that a duplicate may explain is marked in the Note column, so it isn't treated as two separate problems.
- A row whose quantity can't be read is reported under Data Quality, not as a missing order.
