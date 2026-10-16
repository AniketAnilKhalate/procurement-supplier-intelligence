# Procurement & Supplier Intelligence Dashboard

An end-to-end **procurement analytics portfolio project** built with Microsoft Excel, Power BI Desktop, and DAX. The project uses **synthetic practice data only** to demonstrate how purchasing, supplier, delivery, and inventory information can be turned into metrics and visual reports. It does not contain real company or supplier records.

## Table of contents

- [Project overview](#project-overview)
- [Tools and files](#tools-and-files)
- [Dataset at a glance](#dataset-at-a-glance)
- [Business questions](#business-questions)
- [Excel analysis](#excel-analysis)
- [Power BI report](#power-bi-report)
- [Featured DAX measure: Current Stock Gap Units](#featured-dax-measure-current-stock-gap-units)
- [Key observations](#key-observations)
- [Assumptions and limitations](#assumptions-and-limitations)
- [Skills demonstrated](#skills-demonstrated)
- [Suggested CV description](#suggested-cv-description)

## Project overview

The aim was to practise an end-to-end procurement analysis workflow:

1. Prepare and validate synthetic purchasing and inventory data.
2. Create supplier, purchase-order, spend, delivery, and inventory metrics in Excel.
3. Identify purchase orders that may need follow-up and products that warrant inventory review.
4. Build a three-page Power BI report to communicate the analysis.
5. Document the calculations, assumptions, and limitations clearly.

The analysis is designed to support review and prioritisation. It is not an automated purchasing decision or a substitute for checking open orders, business rules, and supplier constraints.

## Tools and files

### Tools

- **Microsoft Excel** — formulas, lookups, sorting, summary tables, reconciliation checks, and charts.
- **Microsoft Power BI Desktop** — KPI cards, bar and column charts, line charts, visual-level Top N filters, and report pages.
- **DAX** — calculated measures, including inventory stock-gap analysis using the latest snapshot.

### Suggested repository files

```text
procurement-supplier-intelligence/
├── README.md
├── Procurement_Supplier_Intelligence_Project.xlsx
└── Procurement_Supplier_Intelligence_Dashboard.pbix
```

Include the current, edited workbook and saved `.pbix` report in the repository. The README can be viewed directly on GitHub. If the Power BI file exceeds GitHub's browser-upload limit, use an appropriate Git/Git LFS workflow or link to a suitable file host rather than omitting the report without explanation.

## Dataset at a glance

The synthetic workbook contains multiple related tables for products, suppliers, supplier quotes, purchase orders and their lines, purchase-order tracking, inventory history, and supplier performance.

| Item | Size / description |
|---|---:|
| Suppliers | 20 |
| Purchase orders | 320 |
| Products | 80 |
| Inventory history | 12 monthly snapshots per product |
| Inventory records | 960 product-month records |

The inventory snapshot history spans 12 months in 2025. The reported counts below reflect the workbook and dashboard as analysed during this project; they are not real-world operational results.

## Business questions

- Which purchase orders are marked as overdue and require urgent follow-up?
- Which suppliers have the largest number of overdue orders?
- How is purchase spend distributed across suppliers?
- How do suppliers compare on delivery and quality indicators?
- What are the recorded delivery lead times and late-delivery counts?
- Which inventory records are at or below their calculated reorder points?
- Which products have the largest stock gap against the reorder point in the latest snapshot?

## Excel analysis

### 1. Inventory reorder review

Added a `ReorderPointUnits` calculation in `Products`:

```excel
=ROUNDUP((E2/365)*F2,0)+G2
```

In the workbook, the referenced columns are `AnnualDemand` (E), `TargetLeadTimeDays` (F), and `SafetyStockUnits` (G). The calculation estimates demand during supplier lead time, rounds up to whole units, and adds safety stock.

Added `ReorderReviewFlag` to `InventoryMonthly`:

```excel
=IF(C2<=XLOOKUP(B2,Products!A:A,Products!I:I),"REORDER REVIEW","Stock OK")
```

Here, `InventoryMonthly` column B contains `ProductID`, column C contains `OnHandUnits`, and `Products` column I contains `ReorderPointUnits`. A monthly record is flagged for review when its on-hand stock is at or below the product's reorder point.

### 2. Supplier performance

Added `SupplierScore` in `SupplierPerformance` as the simple average of on-time delivery and quality acceptance:

```excel
=AVERAGE(C2,D2)
```

The workbook uses column C for `OnTimeDeliveryPct` and column D for `QualityAcceptancePct`. This is an equal-weight practice score, not a formal procurement standard.

Created a `SupplierSummary` with one row per supplier and an average score across its quarterly records:

```excel
=AVERAGEIF(SupplierPerformance!$A:$A,A2,SupplierPerformance!$H:$H)
```

Added illustrative performance categories using these exercise-specific thresholds:

```excel
=IF(B2>=95,"High",IF(B2>=90,"Medium","Low"))
```

- **High:** score of 95 or above
- **Medium:** score from 90 to below 95
- **Low:** score below 90

A count summary and chart were also created for the three categories.

### 3. Purchase-order follow-up

Added `FollowUpFlag` to `PurchaseOrders` based on `POStatus` in column E:

```excel
=IF(E2="Overdue","Urgent Follow-Up",IF(E2="Received","Complete",IF(E2="Open","Monitor",IF(E2="Partially Received","Monitor","Check Status"))))
```

The rules are:

| Recorded PO status | Follow-up flag |
|---|---|
| `Overdue` | `Urgent Follow-Up` |
| `Received` | `Complete` |
| `Open` | `Monitor` |
| `Partially Received` | `Monitor` |
| Any other value | `Check Status` |

Built `POSummary` to count orders by category and `SupplierFollowUp` to count urgent follow-up orders by supplier, sort the suppliers, assign practice priority categories, and chart the suppliers with the highest overdue-order counts.

Observed purchase-order follow-up counts:

| Follow-up category | Orders |
|---|---:|
| Urgent Follow-Up | 35 |
| Monitor | 108 |
| Complete | 177 |
| Check Status | 0 |
| **Total** | **320** |

These classifications reflect the synthetic status values and the rules above.

### 4. Supplier spend

Used a lookup to add `SupplierID` to purchase order lines. In `POLines`, column H contains `SupplierID`, and column F contains `LineValueGBP`.

Created `SupplierSpend` to calculate total spend by supplier:

```excel
=SUMIF(POLines!$H:$H,A2,POLines!$F:$F)
```

Calculated each supplier's share of total spend:

```excel
=IFERROR(B2/SUM($B$2:$B$21),0)
```

Created a top-five supplier spend chart and checked the summary against the total of the purchase-order line values. The grand-total check row was excluded from supplier-level visuals so it would not be treated as a supplier.

### 5. Delivery performance

Retrieved actual receipt dates from `POTracking` into `PurchaseOrders`, then created these fields:

- `ActualReceiptDate` — the recorded receipt date, if available.
- `DeliveryVarianceDays` — actual receipt date minus required date.
- `DeliveryPerformance` — `Late`, `On Time/Early`, or `Not Yet Received`.
- `ActualLeadTimeDays` — actual receipt date minus order date.

The key calculations were:

```excel
=IF(J2="","",J2-D2)
```

```excel
=IF(J2="","Not Yet Received",IF(K2<=0,"On Time/Early","Late"))
```

```excel
=IF(J2="","",J2-C2)
```

Built `DeliverySummary` with late-order count, received-order count, average lead time, and on-time/early rate by supplier. Created supplier delivery charts and a detail table.

Blank actual receipt dates are retained for orders without a recorded receipt date. These orders should not automatically be treated as confirmed late deliveries. Delivery metrics depend on receipt dates matching the correct purchase orders in `POTracking`.

## Power BI report

The report contains three pages.

### Page 1 — Procurement Overview

- KPI cards for supplier count, urgent follow-up orders, total purchase orders, and total purchase spend.
- Top-five supplier purchase-spend chart.
- Top-five suppliers by overdue orders.
- Late deliveries by supplier.
- Monthly inventory records by reorder-review status.

### Page 2 — Supplier Performance

- KPI cards for late orders, received orders, overall on-time/early rate, and average lead time.
- Top-ten suppliers by late deliveries.
- Top-ten suppliers by average lead time.
- Supplier comparison table with late orders, received orders, average lead time, and on-time/early rate.

### Page 3 — Inventory Monitoring

- KPI cards for unique products tracked, monthly inventory records, reorder-review records, and latest snapshot date.
- Monthly trend of inventory records requiring reorder review.
- Inventory records by reorder status.
- Top-ten products by current stock gap in units.
- Inventory detail table filtered to records marked `REORDER REVIEW`.

The inventory product ranking was improved from counting the number of historical months with a reorder flag to calculating each product's stock gap in the latest snapshot. This resolves the problem where many products tied at 12 because they were flagged across all 12 monthly snapshots.

## Featured DAX measure: Current Stock Gap Units

This measure ranks products by how far their on-hand inventory is below their reorder point in the latest available snapshot.

```dax
Current Stock Gap Units =
VAR LatestSnapshotDate =
    CALCULATE(
        MAX(InventoryMonthly[SnapshotDate]),
        REMOVEFILTERS(InventoryMonthly)
    )
RETURN
    CALCULATE(
        SUMX(
            InventoryMonthly,
            MAX(
                0,
                InventoryMonthly[ReorderPointUnits]
                    - InventoryMonthly[OnHandUnits]
            )
        ),
        InventoryMonthly[SnapshotDate] = LatestSnapshotDate
    )
```

### How the measure works

1. **`VAR LatestSnapshotDate`** stores the most recent inventory snapshot date.
2. **`MAX(InventoryMonthly[SnapshotDate])`** finds the latest date in the table.
3. **`REMOVEFILTERS(InventoryMonthly)`** removes filters from the inventory table while finding the overall latest date, so the result is not limited to a particular product's visible history.
4. **`RETURN`** begins the calculation returned by the measure.
5. **`CALCULATE(..., InventoryMonthly[SnapshotDate] = LatestSnapshotDate)`** evaluates the stock-gap calculation using the latest snapshot. In a chart grouped by `ProductID`, the product context allows the measure to return a separate gap for each product.
6. **`SUMX(InventoryMonthly, ...)`** evaluates the expression row by row and adds the results.
7. **`ReorderPointUnits - OnHandUnits`** calculates the number of units by which on-hand stock is below the reorder point.
8. **`MAX(0, difference)`** prevents negative gaps. If stock is at or above the reorder point, the gap is zero; if stock is below it, the positive shortfall is returned.

### Example

If a product has a reorder point of 150 units and the latest snapshot shows 110 units on hand, the gap is **40 units**. If another product has 180 units on hand against a reorder point of 150, its gap is **0 units**.

This measure is a review indicator, **not a purchase recommendation or final order quantity**. It does not deduct outstanding purchase orders, account for supplier constraints, order multiples, minimum order quantities, or future demand changes.

## Key observations

- The workbook contains 20 suppliers and 320 purchase orders.
- The purchase-order follow-up summary flagged 35 orders for urgent follow-up, 108 for monitoring, and 177 as complete.
- The inventory history contains 80 products across 12 monthly snapshots (960 product-month records).
- At the time of analysis, 766 historical inventory records were marked `REORDER REVIEW`. This is a count of monthly records, not 766 unique products or a count of products that need an order today.
- The current stock-gap measure makes the inventory ranking more useful for prioritising the latest stock position than ranking products solely by historical flag counts.

## Assumptions and limitations

- All source records are synthetic practice data; findings are for learning and demonstration, not actual supplier or business decisions.
- The supplier score assigns equal weight to delivery and quality. Different organisations may use different metrics and weights.
- Supplier performance categories and supplier follow-up priority thresholds are illustrative exercise thresholds, not universal standards.
- A status-based `Urgent Follow-Up` flag reflects the recorded `Overdue` status; it is not an independent validation of the real-world order situation.
- Missing receipt dates mean delivery performance cannot be assessed from the recorded dates alone.
- Historical reorder-review counts and latest-snapshot stock gaps answer different questions and should be labelled distinctly.
- The stock-gap measure does not account for open purchase orders or other replenishment constraints.
- Before applying this workflow to real data, validate data types, dates, supplier and product keys, units of measure, and business rules with relevant stakeholders.

## Skills demonstrated

- Excel formulas: `IF`, `IFERROR`, `AVERAGE`, `AVERAGEIF`, `COUNTIF`, `COUNTIFS`, `SUM`, `SUMIF`, and `XLOOKUP`.
- Supplier KPI creation and performance categorisation.
- Purchase-order follow-up and exception flagging.
- Spend aggregation, spend-share calculation, and reconciliation checks.
- Delivery-date variance and actual lead-time calculations.
- Inventory reorder-point and stock-gap logic.
- DAX measures, KPI cards, visual-level filters, and Top N charts in Power BI.
- Dashboard design, data validation, and communication of analytical limitations.

## Suggested CV description

**Procurement & Supplier Intelligence Analysis — Excel and Power BI**

- Analysed synthetic purchase-order, supplier, spend, delivery, and inventory data using Excel formulas, lookups, and summary tables.
- Calculated supplier performance scores, supplier spend and spend share, purchase-order follow-up flags, delivery lead times, and on-time/early rates.
- Created supplier rankings and charts to highlight overdue-order follow-up and spend concentration.
- Built a three-page Power BI report covering procurement overview, supplier performance, and inventory monitoring.
- Developed a DAX measure to rank product stock gaps against reorder points in the latest inventory snapshot and prioritise further review.

---

*This repository is a learning portfolio project. All data is synthetic and should not be represented as real business results.*
<img width="1210" height="681" alt="image" src="https://github.com/user-attachments/assets/9c7357b5-9a8b-4572-b393-2bf428285f77" />
<img width="1213" height="679" alt="image" src="https://github.com/user-attachments/assets/0309e883-159a-4e87-8a1d-5eca6e247a40" />
<img width="1216" height="683" alt="image" src="https://github.com/user-attachments/assets/ef482d78-de58-491a-bd23-718471ec5b03" />
