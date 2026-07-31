# Online Retail Sales Analysis and Excel Dashboard

## Project Overview

This project presents an end-to-end retail data analysis completed in Microsoft Excel using the **UCI Online Retail dataset**. The objective was to transform transaction-level data into a structured analytical workbook that helps management understand sales performance, product demand, customer behaviour, country-level revenue contribution, cancellations, and revenue risk.

The final workbook includes cleaned data, PivotTable-based analysis, validation checks, scenario modelling, and a management-focused dashboard with key performance indicators and interactive filters.

## Business Objectives

This project investigates:

1. How sales revenue changes across months and quarters.
2. Which products generate the highest revenue and sales volume.
3. Which products have the highest cancellation activity.
4. Which customers contribute the most revenue.
5. Which countries generate the highest sales revenue.
6. How lower prices or lower sales quantities affect revenue.
7. Which month is most exposed under the modelled scenarios.

## Dataset

**Dataset:** Online Retail  
**Source:** UCI Machine Learning Repository  
**Reference:** Chen, D. (2015), *Online Retail*  
**DOI:** `10.24432/C5BW33`

The dataset contains transactions from a UK-based non-store online retailer between December 2010 and December 2011.

### Original Variables

| Variable | Description |
|---|---|
| `InvoiceNo` | Unique invoice identifier |
| `StockCode` | Product code |
| `Description` | Product description |
| `Quantity` | Quantity purchased or cancelled |
| `InvoiceDate` | Transaction date and time |
| `UnitPrice` | Price per unit in pounds sterling |
| `CustomerID` | Customer identifier |
| `Country` | Customer country |

## Data Preparation

The raw dataset was preserved before cleaning. The preparation process included:

- checking and removing exact duplicate records;
- identifying cancelled invoices beginning with `C`;
- identifying negative quantities;
- checking zero or negative unit prices;
- identifying missing customer IDs and descriptions;
- converting invoice dates to valid Excel date/time values;
- creating calculated time, revenue, transaction, customer, and product fields.

### Derived Fields

- `Year`
- `Month`
- `Month_Year`
- `Quarter`
- `Transaction_Type`
- `Sales_Revenue`
- `Cancellation_Value`
- `Net_Revenue`
- `Customer_Status`
- `Description_Status`
- `Product_Key`
- `Cancellation_Quantity`

### Example Excel Formulas

```excel
=DATE(YEAR([@InvoiceDate]),MONTH([@InvoiceDate]),1)
```

```excel
="Q"&(INT((MONTH([@InvoiceDate])-1)/3)+1)
```

```excel
=IF([@UnitPrice]<=0,"Invalid Price",
 IF(OR(LEFT([@InvoiceNo]&"",1)="C",[@Quantity]<0),
 "Cancellation","Sale"))
```

```excel
=IF([@Transaction_Type]="Sale",[@Quantity]*[@UnitPrice],0)
```

```excel
=IF([@Transaction_Type]="Cancellation",
 ABS([@Quantity]*[@UnitPrice]),0)
```

```excel
=[@Sales_Revenue]-[@Cancellation_Value]
```

## Workbook Structure

| Worksheet | Purpose |
|---|---|
| `Raw_Data` | Preserves the original imported dataset |
| `Clean_Data` | Contains cleaned records and calculated fields |
| `Monthly_Sales` | Analyses monthly revenue, quantity, invoices, customers, and average order value |
| `Product_Analysis` | Identifies top products, cancellations, unit prices, and product profiles |
| `Customer_Analysis` | Calculates revenue, quantity, invoice count, purchase dates, and recency |
| `Scenario_Testing` | Models price and quantity reduction scenarios |
| `Dashboard` | Presents management KPIs and visual summaries |
| `Checks` | Contains reconciliation and validation tests |
| `Assumptions` | Documents analytical decisions and limitations |

## Analysis Performed

### Monthly Sales Analysis

- Total monthly sales revenue
- Total quantity sold
- Distinct invoice count
- Distinct customer count
- Average order value
- Monthly and quarterly sales trends

### Product Analysis

- Top 10 products by revenue
- Top 10 products by quantity sold
- Products with the highest cancellation quantity
- Cancellation value by product
- Average unit price by product
- High-volume but low-revenue products
- High-revenue but low-quantity products

### Customer Analysis

- Total revenue per customer
- Distinct number of invoices
- Total quantity purchased
- Average order value
- First purchase date
- Most recent purchase date
- Recency in days

Transactions with missing customer IDs were retained for overall sales analysis but excluded from individual customer analysis.

### Country Analysis

Countries were ranked using total sales revenue to identify the markets contributing most strongly to company performance.

## Scenario Testing

| Scenario | Description |
|---|---|
| Base Case | Uses cleaned sales and cancellation values |
| Scenario 1 | Unit prices decrease by 5% |
| Scenario 2 | Sales quantities decrease by 10% |

For each scenario, the workbook calculates revised sales revenue, revised cancellation value, revised net revenue, change from the base case, percentage change, and highest-risk month.

The highest-risk month was defined as the month with the lowest revised net revenue.

## Dashboard

The Excel dashboard was designed for a non-technical management audience and includes:

- total sales revenue;
- total net revenue;
- number of sale invoices;
- number of unique customers;
- average order value;
- monthly sales trend;
- top products by revenue;
- top countries by revenue;
- quarterly revenue distribution;
- scenario comparison;
- interactive slicers.

Add the final dashboard image to the repository:


![Online Retail Dashboard](images/dashboard.png)


## Validation Results

| Validation Measure | Result |
|---|---:|
| Raw records | 541,909 |
| Clean records | 536,641 |
| Exact duplicates removed | 5,268 |
| Missing customer IDs | 135,037 |
| Missing descriptions | 1,454 |
| Cancelled invoices | 9,251 |
| Negative quantities | 10,587 |
| Zero or negative prices | 2,512 |
| Sales revenue | £10,642,110.80 |
| Cancellation value | £893,979.73 |
| Net revenue | £9,748,131.07 |

### Example Validation Formulas

```excel
=IF(Raw_Count-Duplicates_Removed=Clean_Count,"PASS","CHECK")
```

```excel
=IF(ABS(Sales_Revenue-Cancellation_Value-Net_Revenue)<0.01,
"PASS","CHECK")
```

```excel
=IF(ABS(Pivot_Total-Source_Total)<0.01,"PASS","CHECK")
```

All final workbook validation checks should display `PASS`.

## Tools and Techniques

- Microsoft Excel
- Excel Tables
- Structured references
- PivotTables and PivotCharts
- Distinct Count using the Data Model
- Dynamic array functions
- `SUMIFS`
- `COUNTIF`
- `FILTER`
- `UNIQUE`
- `SORT`
- `MAXIFS`
- Conditional formatting
- Slicers
- Scenario modelling
- Reconciliation checks

## Repository Structure

```text
online-retail-sales-analysis-excel/
│
├── README.md
├── workbook/
│   └── ID_MA7444_CW1_Resit.xlsx
├── audit-trail/
│   └── ID_MA7444_CW1_Audit_Trail.pdf
├── images/
│   └── dashboard.png
└── data/
    └── README.md
```

The original dataset should not be uploaded if redistribution is restricted. Provide the official UCI dataset reference instead.

## How to Use the Workbook

1. Download or clone the repository.
2. Open the Excel workbook in Microsoft Excel.
3. Select **Data → Refresh All**.
4. Review the `Checks` worksheet and confirm all checks show `PASS`.
5. Use the dashboard slicers to explore results.
6. Review the `Assumptions` worksheet before interpreting the output.

## Limitations

- Missing customer IDs limit customer-level analysis.
- Product categories were not included in the source dataset.
- December 2011 contains only a partial month.
- Cancellation records do not provide detailed return reasons.
- The scenario model assumes proportional changes and does not model demand elasticity.
- Results are based on one retailer and may not generalise to other businesses.

## Future Improvements

- RFM customer segmentation
- Automated product categorisation
- Revenue forecasting
- Cancellation-rate modelling
- Power Query automation
- Power BI or Tableau dashboard development
- SQL integration for larger datasets

## Key Skills Demonstrated

- Data cleaning and quality assessment
- Excel formula development
- Business-focused analysis
- PivotTable and PivotChart design
- Customer and product analytics
- Scenario and risk modelling
- Dashboard development
- Data validation and documentation

## Author

**Venkata Santosh Kumar Kornu**  
MSc Data Science  
University of Leicester

## Academic Context

This project was completed as part of **MA7444 Excel for Data Science**. The repository demonstrates Excel modelling, analytical methodology, validation, and dashboard design.
