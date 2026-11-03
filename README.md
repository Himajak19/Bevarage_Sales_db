
# Beverage Sales Performance Dashboard | Power BI

## 1. Project Overview

This project analyses a beverage sales dataset for **2022** and presents the results in an interactive Microsoft Power BI dashboard. The report brings together four headline KPIs, retailer revenue comparison, beverage-brand revenue contribution, a revenue trend over time, and slicers for brand, retailer, and region.

The project demonstrates an end-to-end business intelligence workflow:

1. Inspect and validate the source dataset.
2. Clean and prepare data in Power Query Editor.
3. Build a model with a calendar table and date relationship.
4. Create DAX measures for revenue, profit, margin, units sold, brand contribution, and retailer ranking.
5. Build interactive visuals and validate their filter behaviour.
6. Document the dataset, calculations, and business findings.

**Files**
- `Power-BI-DAX-Data-Set.xlsx` — source dataset (worksheet: `Data`)
- `save.pbix` — Power BI report file
- `README.md` — project documentation

## 2. Tools and Technologies

- **Microsoft Power BI Desktop** — data modelling, DAX, report visuals, and slicers
- **Power Query Editor** — data cleaning and transformation
- **DAX (Data Analysis Expressions)** — measures and analytical calculations
- **Excel** — source dataset inspection
- **GitHub Markdown** — project documentation

## 3. Dataset Description

The workbook contains one worksheet named `Data`, with **3,725 rows and 10 columns**. The records cover dates from **2 January 2022 to 25 December 2022**.

| Column | Data type in source | Meaning |
|---|---|---|
| `Retailer` | Text | Retailer name |
| `Retailer ID` | Whole number | Identifier for the retailer |
| `Date` | Date | Date of the sales record |
| `Region` | Text | Sales region |
| `State` | Text | State associated with the record |
| `Beverage Brand` | Text | Beverage brand |
| `Price per Unit` | Decimal number | Selling price per unit |
| `Units Sold` | Whole number | Quantity sold |
| `Revenue` | Decimal number | Revenue for the record |
| `Expenses` | Decimal number | Expenses associated with the record |

### Dataset profile

- **Rows:** 3,725
- **Columns:** 10
- **Retailers:** 4
- **Beverage brands:** 6
- **Regions:** 5
- **States:** 48
- **Distinct dates represented:** 270
- **Exact duplicate rows detected in the raw workbook:** 0
- **Blank/null cells detected in the raw workbook:** 0

The file includes `Revenue` and `Expenses`, but it does not include a separate `Profit` column. Profit is therefore derived as revenue minus expenses in the analytical model.

## 4. Data Inspection and Cleaning in Power Query

Before creating visuals, the source data should be inspected and prepared in **Home → Transform data**. Power Query records each transformation as an Applied Step, which can be repeated during refresh.

The following checks were performed on the supplied workbook as part of this analysis. In Power Query, use the matching transformations if they are not already present in your report.

### 4.1 Confirm headers and schema

- Confirm that the first row is used as column headers.
- Check that all 10 expected columns are present and named consistently.
- Keep descriptive fields such as `Retailer`, `Region`, `State`, and `Beverage Brand` as text.
- Rename fields only when a clearer, consistent name is needed; if you rename a field, update any formulas or visuals that reference it.

### 4.2 Set appropriate data types

Use **Transform → Data Type** to confirm:

- `Date` → Date
- `Retailer ID` → Whole Number
- `Units Sold` → Whole Number
- `Price per Unit`, `Revenue`, and `Expenses` → Decimal Number or Fixed Decimal Number, as appropriate
- Retailer, region, state, and brand → Text

Correct data types matter because they determine how fields aggregate, sort, filter, and participate in relationships.

### 4.3 Check blanks, duplicates, and invalid values

The supplied workbook was checked programmatically:

- No blank/null cells were found in the 10 columns.
- No exact duplicate rows were found.
- No negative values were found in `Price per Unit`, `Units Sold`, `Revenue`, or `Expenses`.
- Two rows have zero `Units Sold`, zero `Revenue`, and zero `Expenses`. These may be legitimate zero-activity records or records to investigate; they should not be deleted automatically without understanding their meaning.
- `Retailer ID` maps to one distinct ID per retailer in this dataset.

These checks do not prove that every business value is correct. For example, a row can be structurally valid but still be an incorrect transaction.

### 4.4 Validate revenue calculation

The data satisfies this relationship for every row:

`Revenue = Price per Unit × Units Sold`

This is a useful quality check. In Power Query, a temporary custom column can be created to calculate price multiplied by units, then compared with `Revenue`. Remove the temporary column after validation if it is not needed in the final model.

### 4.5 Standardize text and review categories

If needed, use **Transform → Format → Trim/Clean** on text columns and inspect category values for spelling variations or unwanted spaces. Do not change valid names just to make them look different.

Check the distinct values of retailer, brand, region, and state fields. This dataset contains 4 retailers, 6 brands, 5 regions, and 48 states.

### 4.6 Applied Steps and loading

A typical Power Query sequence is:

1. Source
2. Navigation to the `Data` worksheet
3. Promoted Headers (if required)
4. Changed Type
5. Any required text standardization or validation transformations
6. Close & Apply

The exact Applied Steps saved in the PBIX cannot be confirmed from the report layout alone. The list above is a recommended sequence, not a claim that every listed step was used. Retain only steps that match the actual query.

## 5. Data Modelling

The report references the main sales table, `Data`, and a date table named `Calendar`.

### 5.1 Calendar table

A calendar table supports chronological analysis and consistent date filtering. A basic DAX example for this dataset is:

```DAX
Calendar =
CALENDAR (
    DATE ( 2022, 1, 1 ),
    DATE ( 2022, 12, 31 )
)
```

The source data ranges from 2 January to 25 December 2022, so a full-year calendar includes the year boundaries even where the source has no record on a particular day. If the model needs to support future years or additional source data, extend the calendar dynamically.

Recommended steps:

1. Ensure `Calendar[Date]` is a Date data type.
2. Mark `Calendar` as a date table using `Calendar[Date]`.
3. Create an active one-to-many relationship from `Calendar[Date]` (one side) to the sales table's `Date` column (many side).
4. Use the calendar date for the line chart axis and date-based filtering.
5. Check that the sales date column is also a Date data type.

Conceptual model:

```text
Calendar[Date]  1 ───────── *  Data[Date]
                               ├── Retailer
                               ├── Region / State
                               ├── Beverage Brand
                               └── Price / Units / Revenue / Expenses
```

Confirm the exact relationship column and settings in Model view. The diagram is conceptual.

### 5.2 Why modelling matters

- A calendar table gives the report a consistent date dimension.
- Explicit measures centralize calculation logic and make KPI behaviour more consistent.
- Relationships determine how filters travel through the model.
- Slicers should filter the intended visuals without introducing ambiguous relationships or unexpected totals.

After modelling, test each slicer and verify that the cards and charts update as expected.

## 6. DAX Measures

The formulas below are appropriate examples for the supplied workbook's column names. They explain the business logic and can be adapted to the exact measure names in the PBIX. Create measures using **Modeling → New measure**.

### 6.1 Total Revenue

```DAX
Total Revenue =
SUM ( Data[Revenue] )
```

**Purpose:** Adds revenue across the current filter context. When a user selects a region or brand, the result should reflect that selection.

**Dataset result:** Total revenue is **8,166,072.50** in the source data.

### 6.2 Total Expenses

```DAX
Total Expenses =
SUM ( Data[Expenses] )
```

**Purpose:** Aggregates expenses for the selected data. This measure is useful for reconciliation and can also be used in a report visual if desired.

**Dataset result:** Total expenses are **1,752,700.95**.

### 6.3 Total Profit

```DAX
Total Profit =
[Total Revenue] - [Total Expenses]
```

**Purpose:** Calculates profit as revenue minus expenses. Using measures makes the calculation responsive to the same slicers as revenue and expenses.

**Dataset result:** Total profit is **6,413,371.55**.

### 6.4 Profit Margin %

```DAX
Profit Margin % =
DIVIDE ( [Total Profit], [Total Revenue], 0 )
```

**Purpose:** Measures profit as a proportion of revenue. `DIVIDE` handles a zero denominator safely. Format the measure as a percentage in the model.

**Dataset result:** The overall margin is approximately **78.54%**. This is based on the dataset's definition of profit as revenue minus expenses; confirm that this definition matches the business context.

### 6.5 Total Units Sold

```DAX
Total Units Sold =
SUM ( Data[Units Sold] )
```

**Purpose:** Adds units sold across the selected records.

**Dataset result:** **16,401,800 units**.

### 6.6 Revenue % of Total by Brand

```DAX
Revenue % of Total by Brand =
DIVIDE (
    [Total Revenue],
    CALCULATE (
        [Total Revenue],
        REMOVEFILTERS ( Data[Beverage Brand] )
    ),
    0
)
```

**Purpose:** Shows the selected brand's revenue as a share of revenue across all brands, while preserving other filters such as retailer and region. Format as a percentage.

**Filter-context note:** This denominator removes only the brand filter. If the desired business question is a share of all revenue regardless of other slicers, a different DAX pattern is required. Test the measure by selecting a brand, region, and retailer.

### 6.7 Retailer Revenue Rank

```DAX
Retailer Revenue Rank =
RANKX (
    ALL ( Data[Retailer] ),
    [Total Revenue],
    ,
    DESC,
    DENSE
)
```

**Purpose:** Ranks retailers by revenue, with rank 1 assigned to the highest revenue. `DENSE` ranking does not skip a rank after ties.

**Filter-context note:** `ALL(Data[Retailer])` removes the retailer filter for the ranking comparison but retains other filters. If the rank should respond only to visible/selected retailers, consider `ALLSELECTED` and test the intended behaviour.

## 7. Dashboard Design and Deep Visual Explanation

The uploaded report has one page, titled **Beverage Sales Performance Dashboard**, with the subtitle:

*Revenue • Profitability • Retailer & Brand Analysis • 2022*

The report includes four KPI cards, a retailer revenue bar chart, a beverage-brand revenue-share doughnut chart, a revenue trend line chart, and slicers for beverage brand, retailer, and region.

### 7.1 KPI cards

#### Total Revenue
Shows the total sales value for the active selection. Revenue provides a high-level view of business scale, but it does not by itself show how much money remains after expenses.

#### Total Profit
Shows revenue minus expenses. Comparing profit with revenue helps distinguish sales volume from financial contribution.

#### Profit Margin %
Shows profit relative to revenue. For example, a margin of 20% means 20 units of profit for every 100 units of revenue. A high margin can be valuable even when revenue is lower, so this KPI should be interpreted together with revenue and profit.

#### Total Units Sold
Shows sales volume. Units can help distinguish a revenue change caused by selling more items from a change driven by price or product mix.

**How to read the cards together:** Review revenue, profit, margin, and units as a group. Use slicers to compare the same four KPIs for different brands, retailers, and regions. Confirm that all cards respond correctly to the intended filters.

### 7.2 Retailer revenue bar chart

**Question answered:** Which retailers contribute the most revenue?

The length of each bar represents revenue for a retailer in the current filter context. Sorting the bars from highest to lowest makes performance comparisons easier. The report metadata references `Total Revenue` and `Retailer Revenue Rank` for this visual.

**Analysis approach:**
- Identify the leading retailer and compare it with the other retailers.
- Select a region to see whether retailer performance changes geographically.
- Compare revenue with profit and margin; a revenue leader is not automatically the most profitable retailer.
- Confirm that ranking and sort order respond as intended to the slicers.

### 7.3 Beverage-brand revenue-share doughnut chart

**Question answered:** What proportion of revenue comes from each beverage brand?

Each slice represents a brand's share of revenue under the active filters. Larger slices mean a greater revenue contribution. The report references `Beverage Brand` and `Revenue % of Total by Brand`.

**Analysis approach:**
- Identify the brands with the largest and smallest revenue contributions.
- Select a retailer or region to explore changes in brand mix.
- Check whether the percentages total approximately 100% for the intended comparison group.
- Validate the DAX denominator: percentages can behave differently depending on which filters the measure removes.

A doughnut chart is useful for a quick composition view. If the slices are very similar or there are many categories, a sorted bar chart may make exact comparisons easier.

### 7.4 Revenue trend line chart

**Question answered:** How does revenue change over time?

The line chart uses `Calendar[Date]` for the date axis and `Total Revenue` for the value. A line chart makes peaks, dips, and broad movement easier to identify.

**Analysis approach:**
- Check that the date axis is chronological.
- Identify periods with higher and lower revenue.
- Use brand, retailer, and region slicers to compare trends within a subset.
- Consider monthly aggregation for a cleaner overview if a daily line is too noisy.

A single year can show variation within that year, but it is not enough by itself to prove a recurring seasonal pattern.

### 7.5 Slicers

- **Beverage Brand:** filters to one or more brands.
- **Retailer:** filters to one or more retailers.
- **Region:** filters to one or more regions.

Slicers allow users to explore the same KPIs and visuals without creating separate pages for every category. For example, select `West` to examine its revenue, profit, margin, units sold, retailer performance, brand contribution, and time trend.

Test single selections, multiple selections, and clearing selections. Use **Format → Edit interactions** to control which visuals respond to each slicer.

## 8. Findings from the Supplied Dataset

The following are calculated from the supplied Excel dataset, not estimates. Monetary values use the units represented in the source file; the dataset does not explicitly state a currency symbol in the column names, so the README does not assume one.

### 8.1 Overall performance

| Metric | Calculated result |
|---|---:|
| Records | 3,725 |
| Total units sold | 16,401,800 |
| Total revenue | 8,166,072.50 |
| Total expenses | 1,752,700.95 |
| Derived total profit | 6,413,371.55 |
| Derived profit margin | 78.54% |

Profit is calculated as `Revenue − Expenses`, and margin is `Profit ÷ Revenue`. These figures should reconcile with the Power BI cards if the measures and filter context are configured consistently.

### 8.2 Revenue by retailer

| Retailer | Records | Revenue | Derived profit | Profit margin |
|---|---:|---:|---:|---:|
| Sodapop | 2,285 | 4,354,537.50 | 3,420,668.30 | 78.60% |
| FizzySip | 720 | 2,115,587.50 | 1,663,961.50 | 78.66% |
| BevCo | 576 | 1,291,535.00 | 1,012,827.25 | 78.43% |
| DreamCo | 144 | 404,412.50 | 315,914.50 | 78.12% |

**Observation:** Sodapop has the highest total revenue in the dataset, followed by FizzySip. Sodapop also has the most records, so its larger total reflects its larger representation in the data as well as its revenue. For a fairer comparison, consider average revenue per record or compare retailers across the same time and regional selection.

### 8.3 Revenue by beverage brand

| Beverage brand | Revenue | Derived profit | Profit margin |
|---|---:|---:|---:|
| Coca-Cola | 1,911,540.00 | 1,493,113.85 | 78.12% |
| Dasani Water | 1,630,462.50 | 1,281,858.50 | 78.58% |
| Diet Coke | 1,383,612.50 | 1,090,466.25 | 78.79% |
| Sprite | 1,151,607.50 | 905,667.08 | 78.64% |
| Powerade | 1,125,025.00 | 883,887.75 | 78.57% |
| Fanta | 963,825.00 | 758,378.13 | 78.68% |

**Observation:** Coca-Cola has the highest total revenue among the six brands, while Fanta has the lowest. The margins are relatively close, so revenue contribution and units sold are useful additional measures when comparing brands. These are full-dataset totals; the ranking may change when slicers are applied.

### 8.4 Revenue by region

| Region | Records | Revenue | Derived profit |
|---|---:|---:|---:|
| West | 792 | 2,366,237.50 | 1,859,537.00 |
| Northeast | 922 | 1,753,305.00 | 1,376,252.45 |
| Southeast | 499 | 1,604,885.00 | 1,259,824.65 |
| South | 576 | 1,291,535.00 | 1,012,827.25 |
| Midwest | 936 | 1,150,110.00 | 904,930.20 |

**Observation:** West has the highest total revenue and profit, while Midwest has the lowest total revenue. However, Midwest has the largest number of records, demonstrating why record count alone should not be interpreted as sales performance. Compare revenue per record and use consistent periods or retailer selections for deeper analysis.

### 8.5 Monthly revenue trend

| Month (2022) | Revenue |
|---|---:|
| January | 477,275.00 |
| February | 465,775.00 |
| March | 459,805.00 |
| April | 468,825.00 |
| May | 643,872.50 |
| June | 843,100.00 |
| July | 969,637.50 |
| August | 888,450.00 |
| September | 651,025.00 |
| October | 592,037.50 |
| November | 739,057.50 |
| December | 967,212.50 |

**Observation:** March has the lowest monthly revenue in this dataset, and July has the highest. Revenue rises strongly from April through July, declines from August through October, and rises again in November and December. This is a description of the observed 2022 data, not proof of a recurring seasonal effect.

### 8.6 Data-quality findings

- `Revenue` equals `Price per Unit × Units Sold` for every record.
- No exact duplicate rows were found.
- No nulls/blanks or negative numeric values were detected in the inspected columns.
- Two records have zero units, revenue, and expenses. These should be checked against business rules before deciding whether to keep them.
- The source has records on 270 distinct dates across the 2022 reporting period.

## 9. Suggested Business Questions

The dashboard can be used to investigate:

1. Which retailer contributes the most revenue, and does it also lead in profit?
2. Which brand contributes the most revenue overall and within each region?
3. How does the brand revenue mix change by retailer?
4. Which months have the highest and lowest revenue?
5. Are higher revenue periods also associated with higher units sold?
6. Which regions have high revenue relative to their record counts?
7. Which retailer or brand has the highest profit margin under the selected filters?
8. Do retailer rankings change when a region or brand is selected?

Use the actual report selections and values to answer these questions. Do not assume that full-dataset rankings remain the same for every filtered view.

## 10. Validation Checklist

Before publishing the report:

- Reconcile the total revenue with the source dataset.
- Reconcile expenses and confirm the profit formula.
- Format `Profit Margin %` as a percentage.
- Confirm that `Calendar[Date]` has an active relationship to the sales date.
- Check that the date axis is sorted chronologically.
- Validate the brand-share denominator and the retailer-rank behaviour.
- Test every slicer with one selection, multiple selections, and cleared selections.
- Confirm that KPI cards and all intended charts respond correctly.
- Check that zero-activity records are handled according to business rules.
- Refresh the report and verify that Power Query steps still work.

## 11. Limitations and Assumptions

- The monetary unit/currency is not explicitly stated in the source column names.
- Profit is derived as `Revenue − Expenses`; it is not supplied as a source column.
- The dataset is for 2022 and has 270 distinct dates, not a record for every calendar day.
- Two zero-activity records exist and are retained in the analysis until their business meaning is confirmed.
- Retailer totals are affected by different record counts; totals alone do not establish performance per transaction.
- The dashboard's visual layout and measure names were inspected, but the complete saved Power Query Applied Steps and all stored DAX expressions were not independently extracted. The DAX shown in this README documents a recommended implementation using the real source column names; compare it with the exact formulas in the PBIX.

## 12. How to Open the Project

1. Download or clone the project repository.
2. Open `save.pbix` in Microsoft Power BI Desktop.
3. If prompted, reconnect the source to `Power-BI-DAX-Data-Set.xlsx`.
4. Select **Refresh**.
5. Use the brand, retailer, and region slicers to explore the dashboard.

Before publishing to Power BI Service, verify data-source credentials, refresh settings, privacy settings, and report access permissions.

## 13. Skills Demonstrated

- Excel dataset inspection and data quality checks
- Power Query data cleaning and transformation
- Data types and category validation
- Calendar table and relationship modelling
- DAX measures for revenue, expenses, profit, margin, contribution, and ranking
- KPI cards, bar charts, doughnut charts, line charts, and slicers
- Filter-context testing and interactive report validation
- Business analysis and data storytelling
- GitHub project documentation

## 14. Project Summary

This project transforms a 3,725-row beverage sales dataset into an interactive Power BI report. The analysis identifies Sodapop as the largest retailer by total revenue, Coca-Cola as the leading brand by revenue, West as the leading region by total revenue, and July as the highest-revenue month in 2022. The report combines those business views with revenue, profit, margin, and units-sold KPIs, while slicers enable users to investigate the results by brand, retailer, and region.
=======
# Bevarage-sales-dashboard
>>>>>>> a062e1fac52e3554f5bd2cfc98b67bc9c257fb56
