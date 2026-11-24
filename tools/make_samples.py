"""Regenerate the synthetic test files in samples/ with deliberately planted errors.
Run from the repository root:  python tools/make_samples.py
"""
import random, pandas as pd
random.seed(7)
skus = ["GN-TEE-BLK-M","GN-TEE-WHT-L","GN-HOOD-GRY-M","GN-CAP-NVY","GN-MUG-01","GN-SOCK-3PK","GN-BAG-TOTE","GN-BTL-750"]
dates = pd.date_range("2026-11-01", periods=20, freq="D")
seller=[]
for i in range(1, 61):
    oid=f"GN-{10400+i}"
    for sku in random.sample(skus, random.choice([1,1,2])):
        seller.append({"Order ID":oid,"Order Date":random.choice(dates).date(),"SKU":sku,"Quantity":random.choice([1,1,2,3]),"Channel":random.choice(["Shopify","Amazon","eBay"])})
wh=[{"order_no":r["Order ID"].lower() if random.random()<0.15 else r["Order ID"],"product_code":r["SKU"],"qty_shipped":r["Quantity"],"dispatch_date":r["Order Date"],"carrier":random.choice(["Royal Mail","DPD","Evri"])} for r in seller]
# Planted issues
missing = {"GN-10412","GN-10437","GN-10455"}            # sold but never shipped
wh=[r for r in wh if r["order_no"].upper() not in missing]
wh.append({"order_no":"GN-10399","product_code":"GN-MUG-01","qty_shipped":2,"dispatch_date":dates[2].date(),"carrier":"DPD"})  # shipped, no sale
wh.append({"order_no":"GN-10501","product_code":"GN-CAP-NVY","qty_shipped":1,"dispatch_date":dates[9].date(),"carrier":"Evri"})
wh.append(dict(wh[5])); wh.append(dict(wh[20]))      # warehouse duplicates
seller.append(dict(seller[30]))                      # seller duplicate
for idx, newq in [(8,0),(15,5),(27,1),(41,4)]:       # quantity mismatches
    wh[idx]["qty_shipped"]=newq if newq!=wh[idx]["qty_shipped"] else newq+1
wh[33]["qty_shipped"]="two"                          # bad data
seller.append({"Order ID":"","Order Date":dates[0].date(),"SKU":"GN-SOCK-3PK","Quantity":1,"Channel":"eBay"})
pd.DataFrame(seller).to_excel("samples/seller_orders.xlsx", index=False)
pd.DataFrame(wh).to_csv("samples/warehouse_orders.csv", index=False)
print(len(seller), len(wh))
