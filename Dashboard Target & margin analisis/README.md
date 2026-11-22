# Sales Target and Margin Analysis Dashboard

A Data Analyst portfolio project. An interactive Excel dashboard (Power Query, Power Pivot, DAX) that tracks target attainment and margin for a distributor of everyday goods with 12 branches across Java Island.

> **Note:** all data is **synthetic** (created for practice). Company, branch, store, salesperson, and product brand names are fictional. The findings below apply to this data only, not to a real company.

![Target Analysis Page](screenshots/halaman-1-target-analisis.png)
![Margin Analysis Page](screenshots/halaman-2-margin-analisis.png)

---

## Summary

| Item | Details |
|---|---|
| Data period | 1 January to 18 November 2026 (November is still in progress; analysis covers the complete months January-October) |
| Scope | 12 branches in Java, 48 customers, 25 products, 554 transaction rows |
| Currency | Rupiah (IDR) |
| Tools | Excel (Power Query, Power Pivot / Data Model), DAX, Google Sheets (data source), VBA (navigation and full screen) |
| Output | A 2-page dashboard: **Target Analysis** and **Margin Analysis** |

### Key findings

1. Sales for January-October reached **96.2% of target** (Rp3.18 B against Rp3.30 B), a gap of **-Rp124.1 million**.
2. The shortfall is concentrated in **Central Java-DIY**: its four branches account for -Rp115.2 million, equal to 93% of the net gap.
3. In October, DI Yogyakarta reached only **27% of its target** (Rp9.79 M against Rp36.00 M) because its two largest customers placed no orders at all.
4. **Discounts erode margin.** October gross margin was only **4.61%**, below the January-October average (5.28%). Supermarket and Wholesale (Grosir) make up about 83% of the discount value, and the Supermarket margin dropped to **2%**.

---

## Business background

PT Sinar Nusantara Distribusi is a distributor of staple foods, instant food, beverages, and household and personal care products. Its customers are small grocery stores, minimarkets, supermarkets, and wholesalers.

**Problems to solve**

1. Each branch sends its report in its own Excel file, merging them takes 3-4 days, and the figures are often inconsistent.
2. Target attainment is only reviewed at the end of each quarter, so lagging branches are noticed too late.
3. Discounts to large segments rose in the second half of the year, but their effect on margin cannot be seen.

### Core question

> Which branches, segments, and products cause sales to fall behind target and margin to thin out, and what action should be taken this month?

### Audience

- **Primary: Sales Director.** Decides on discount policy, target allocation, and interventions at branches. Uses the dashboard in the monthly review meeting.
- **Secondary: Area Managers.** Drill into branches, products, and customers in their area.

---

## Data and model

The data is organized as a **star schema** in Google Sheets, then loaded into the Excel Data Model.

| Table | Type | Contents |
|---|---|---|
| `Fact_Penjualan` | Fact | Transactions by product: date, customer, product, qty, discount, payment method, delivery status |
| `Fact_Target` | Fact | Monthly sales target per branch |
| `Fact_Target_Segmen` | Fact | Monthly target per branch and customer segment |
| `Dim_Tanggal` | Dimension | Daily calendar (year, quarter, month, week, day) |
| `Dim_Wilayah` | Dimension | 12 branches, provinces, areas |
| `Dim_Sales` | Dimension | 15 salespeople |
| `Dim_Pelanggan` | Dimension | 48 customers, segment, account-holding salesperson, credit limit |
| `Dim_Produk` | Dimension | 25 products, category, brand, cost price, selling price, target weight |
| `Dim_Segmen` | Dimension | 4 customer segments, used as a bridge to the target table |

The derived columns in `Fact_Penjualan` (region, salesperson, unit price, total sales, COGS, gross profit) are calculated with formulas from the dimension tables.

### Main relationships

```
Dim_Tanggal[Tanggal]      -> Fact_Penjualan[Tanggal]
Dim_Pelanggan[ID]         -> Fact_Penjualan[ID_Pelanggan]
Dim_Produk[ID_Produk]     -> Fact_Penjualan[ID_Produk]
Dim_Wilayah[ID_Wilayah]   -> Fact_Penjualan, Fact_Target, Fact_Target_Segmen
Dim_Segmen[Segmen]        -> Dim_Pelanggan[Segmen], Fact_Target_Segmen[Segmen]
```

`Fact_Target` cannot be linked directly to `Dim_Tanggal`, since one month has many date rows. Targets are matched through `Periode_Bulan` inside the DAX measure (see below).

---

## Metric definitions

| Dashboard metric | Definition |
|---|---|
| Gross Sale | Qty x Unit Price, before discount |
| Nominal discount | Gross Sale x Discount %, calculated per row and then summed |
| Net Revenue | Gross Sale - Nominal discount |
| COGS | Qty x Cost Price |
| Gross Profit | Net Revenue - COGS |
| Margin | Gross Profit / Net Revenue |
| COGS ratio | COGS / Net Revenue (= 1 - Margin) |
| Gap | Actual - Target |
| Attainment % | Actual / Target |

**Why is the discount weighted?** A simple average of the `Diskon_Persen` column treats every transaction as equally important. Large transactions with higher discounts then count for less, so the discount looks smaller than the money actually lost. The weighted discount (nominal discount divided by Gross Sale) always matches the difference in profit before and after discount.

### DAX measure examples

```dax
Gross Sale =
SUMX ( Fact_Penjualan, Fact_Penjualan[Qty] * Fact_Penjualan[Harga_Satuan] )

Diskon Nominal =
SUMX (
    Fact_Penjualan,
    Fact_Penjualan[Qty] * Fact_Penjualan[Harga_Satuan] * Fact_Penjualan[Diskon_Persen]
)

Net Revenue = [Gross Sale] - [Diskon Nominal]

Gross Profit = [Net Revenue] - SUM ( Fact_Penjualan[Total_HPP] )

Margin % = DIVIDE ( [Gross Profit], [Net Revenue] )
```

Monthly target, matched through `Periode_Bulan`:

```dax
Target =
SUMX (
    VALUES ( Dim_Tanggal[Periode_Bulan] ),
    CALCULATE (
        SUM ( Fact_Target_Segmen[Target_Penjualan] ),
        FILTER (
            Fact_Target_Segmen,
            Fact_Target_Segmen[Periode_Bulan] = Dim_Tanggal[Periode_Bulan]
        )
    )
)
```

Target per category or product (split according to the January-June sales share):

```dax
Target Produk = [Target] * SUM ( Dim_Produk[Bobot_Target] )
```

Year-to-date accumulation (for the growth chart):

```dax
Revenue Kumulatif =
CALCULATE (
    [Net Revenue],
    FILTER ( ALL ( Dim_Tanggal[Tanggal] ), Dim_Tanggal[Tanggal] <= MAX ( Dim_Tanggal[Tanggal] ) )
)
```

Month names are fixed (they do not follow the Excel language setting) so the file opens safely on other computers:

```dax
Month Short =
SWITCH (
    Dim_Tanggal[Nama_Bulan],
    "Januari", "Jan", "Februari", "Feb", "Maret", "Mar", "April", "Apr",
    "Mei", "May", "Juni", "Jun", "Juli", "Jul", "Agustus", "Aug",
    "September", "Sep", "Oktober", "Oct", "November", "Nov", "Desember", "Dec"
)
```

---

## Dashboard

### Page 1: Target Analysis

| Visual | Purpose |
|---|---|
| Gap, Actual, Target cards | Attainment position under the selected filters |
| Biggest Driver card | Province with the largest gap |
| Gap Actual vs Target by Branch | Who is above and below target |
| Actual vs Target by month | Cumulative growth of actual against target |
| Actual vs Target by Product Category | Product categories that are behind or ahead |
| Attainment Percentage by Province | Map of attainment per province |

Filters: **Month** and **Province**.

### Page 2: Margin Analysis

| Visual | Purpose |
|---|---|
| Net Revenue (with Gross Sale), COGS (with COGS ratio), Gross Profit (with Margin) cards | Profit structure for one month |
| Margin Pressure card | Segment with the thinnest margin |
| Nominal discount vs Net Revenue by Segment | Size of the discount relative to sales for each segment |
| Margin vs COGS Ratio by Segment | Margin for each segment |
| Gross Profit by Category | Profit contribution per product category |
| Net Revenue vs Margin by Month | Monthly trend of revenue and margin |

Filters: **Month** and **Segment**.

Additional features: page navigation buttons (Target Analysis, Margin Analysis, Back), a full-screen button (VBA macro), custom number formats (Rp...jt, Rp...M), and green/red colors depending on the sign of the number.

---

## Findings

### 1. Total sales below target (January-October)

| | Value |
|---|---|
| Actual | Rp3.176 billion |
| Target | Rp3.300 billion |
| Gap | -Rp124.1 million (-3.8%) |
| Attainment | **96.2%** |

Seven of the 12 branches are below target:

| Branch | Gap |
|---|---|
| Yogyakarta | -Rp43.0 M |
| Semarang | -Rp34.9 M |
| Tegal | -Rp26.5 M |
| Bandung | -Rp16.9 M |
| Surakarta | -Rp10.8 M |
| Cirebon | -Rp6.9 M |
| Tangerang | -Rp5.7 M |

The other five branches exceeded their targets, but only by Rp20.6 million in total. **The four Central Java-DIY branches (Yogyakarta, Semarang, Tegal, Surakarta) account for a shortfall of Rp115.2 million, equal to 93% of the net gap.** This area's attainment is only 85.6%, while East Java is at 102.1% and DKI Jakarta at 100.5%.

Adding November (data through 18 November), cumulative actual is Rp3.34 B against a target of Rp3.64 B. The November target covers a full month while the realization is only partial, so this month's gap will narrow as the month progresses.

### 2. October: total above target, but uneven

October reached Rp359.1 M against a target of Rp329 M (109%). Gap by province:

| Province | October Gap |
|---|---|
| Central Java | +Rp19.31 M |
| East Java | +Rp15.92 M |
| DKI Jakarta | +Rp14.52 M |
| Banten | +Rp11.51 M |
| West Java | -Rp4.95 M |
| **DI Yogyakarta** | **-Rp26.21 M** |

DI Yogyakarta is the biggest driver of the shortfall, with attainment of only 27% (Rp9.79 M against Rp36.00 M).

### 3. Why Yogyakarta fell behind in October

Yogyakarta's target is concentrated in two segments, and neither generated any sales:

| Segment | Actual | Target |
|---|---|---|
| Wholesale (Grosir) | Rp0 | Rp19.2 M |
| Minimarket | Rp0 | Rp7.2 M |
| Supermarket | Rp3.8 M | Rp7.8 M |
| Small Grocery Store (Toko Kelontong) | Rp6.0 M | Rp1.8 M |

Wholesale and Minimarket carry about 73% of the target, yet customers in both segments placed no orders throughout October. Small grocery stores, on the other hand, exceeded their target. This means the problem lies with **two large customers who did not order**, not with the whole branch.

### 4. Thin margin in October, and large discounts to two segments

| Measure (October) | Value |
|---|---|
| Gross Sale | Rp381 M |
| Net Revenue | Rp359 M |
| COGS | Rp343 M (ratio 95.39%) |
| Gross Profit | Rp17 M |
| Margin | **4.61%** |

October's margin is lower than the January-October average (5.28%), because October's discount was higher (about 5.8% of Gross Sale, compared with 4.97% across January-October).

By segment:

| Segment | Nominal discount | Net Revenue | Margin | COGS Ratio |
|---|---|---|---|---|
| Wholesale (Grosir) | Rp9.33 M | Rp139 M | 4% | 96% |
| Supermarket | Rp8.82 M | Rp111 M | **2%** | 98% |
| Minimarket | Rp2.86 M | Rp71 M | 7% | 93% |
| Small Grocery Store (Toko Kelontong) | Rp0.82 M | Rp38 M | 8% | 92% |

**Supermarket and Wholesale make up about 83% of all discount value**, even though their share of Net Revenue is about 70%. The segments with the largest discounts have the thinnest margins, and this is what the Margin Pressure card flags (Supermarket).

### 5. Supermarket margin has declined since mid-year

The Supermarket margin ranged around 3-6% in January-June (averaging about 4%), then fell to around 2-3% in July-October, with the lowest point of 2% in October. This decline lines up with the rise in discounts in the second half.

Supermarket Gross Profit in October was only Rp2.47 M, made up of Instant Food Rp853 K, Beverages Rp796 K, Staples Rp343 K, Personal Care Rp322 K, and Household Care Rp155 K.

---

## Recommendations

1. **Contact the two large customers in Yogyakarta** (Wholesale and Minimarket) who did not order in October, and find the cause before the quarter ends.
2. **Review the discount limits for Supermarket and Wholesale.** Ask for a justification on every discount above the set limit, since these two segments erode margin the most.
3. **Monitor the four Central Java-DIY branches monthly**, not quarterly, because their shortfall has been accumulating since mid-year.
4. **Track Gross Profit by segment alongside the sales target.** Sales that hit target are not necessarily healthy if they were achieved through heavy discounting.

All the points above are hypotheses drawn from sales data patterns. The exact causes need to be confirmed with the branch and sales teams.

---

## Limitations

- **Synthetic data.** The patterns were built in deliberately (discounts rising in the second half, Central Java-DIY sales weakening, certain customers stopping orders) so that they could be discovered through analysis. These findings do not represent a real company.
- **Synthetic targets.** Some targets were created to follow the realization pattern, which makes them less realistic than targets set before the month begins.
- **Product targets are allocations.** Category and product targets are split from branch targets according to the January-June sales share, rather than being set independently.
- **Gross margin only.** The data contains cost price only, with no operating expenses, so net margin cannot be calculated.
- **Uniform price per product.** All segments buy at the same base price, so margin differences between segments come only from discounts.
- **Small volume per branch.** With a limited number of transactions, sales at some branches swing sharply from month to month, and can even be zero in certain months.
- **Chart figures are rounded.** Margins by segment and by month are shown in the charts as whole percentages.
- **November is incomplete.** Realization through 18 November is compared against a full-month target.

---

## Repository structure

```
.
├── README.md
├── Project 4.xlsm                  # Excel dashboard
├── data/
│   └── Penjualan.xlsx              # export of fact and dimension tables
├── screenshots/
│   ├── halaman-1-target-analisis.png
│   └── halaman-2-margin-analisis.png
└── assets/
    └── demo.gif                    # recording of the dashboard in use
```

Sheets inside the Excel file: `Target` (page 1), `Margin` (page 2), and `Detail` (the pivot that feeds the charts).

## How to use

1. Open `Project 4.xlsm` in Excel for Windows (2013 or later, with Power Pivot).
2. Click **Enable Content** so the navigation and full-screen buttons work.
3. Select a month, province, or segment in the slicers to filter the view. Use the filter icon with the cross to clear the selection.
4. To update the data, replace the contents of the source tables, then click **Data > Refresh All**.

The province map requires an internet connection (Bing Maps). Macros do not run in Excel for the web.
