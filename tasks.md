# tasks.md

1. **Familiarize yourself with `part_history_checker.py`**  
   - Key functions:  
     - `query_part_manufacturing_history()`  
     - `query_part_sales_history()`  
     - `query_part_average_cost()`  
     - `generate_part_summary()`  

2. **Define the output schema**  
   Extend `generate_part_summary()` (or add a new function) so it returns a dict like:
   ```python
   {
     "PartNumber": str,
     "CurrentRevision": str,
     "TotalBuilds": int,
     "BuildsByRevision": {"002":18, "03":8, "04":3, "NS":6},
     "AvgCostAllRevs": float,
     "AvgCostCurrentRev": float or None,
     "RecentUnitCost": float,
     "RecentSalesOrders": [
       {"OrderDate":"2022-06-16","Qty":8,"SONumber":"048659","TotalValue":326.88},
       … up to 5 records …
     ],
     "AnnualRevenue": {
       2020:1168000.00,
       2021:12043189.60,
       2022:146769.12,
       2023:0.00,
       2024:0.00,
       2025:0.00,
     },
     "AvgAnnualRevenue": float,
     "RFQQty": int,                # comes from the CSV row
     "RecentTotalValue": float,    # from the latest non-zero sale
     "RecentSOQty": int,           
     "RecentSODate": "YYYY-MM-DD",
     "RecentUnitPrice": float,     # computed = RecentTotalValue/RecentSOQty
     "PotentialRevenue": float,    # = RFQQty*RecentUnitPrice
     "UnitMargin": float,          # = RecentUnitPrice−RecentUnitCost
     "TotalMargin": float,         # = PotentialRevenue−(RecentUnitCost*RFQQty)
     "RiskByPotential": str,       # Low/Med/High/Very High
     "RiskByAvgAnnual": str,       # Low/Med/High/Very High
   }
