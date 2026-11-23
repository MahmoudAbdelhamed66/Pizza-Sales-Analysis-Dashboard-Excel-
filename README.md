 # 🍕 Pizza Sales Analysis Dashboard

An Excel project that takes a year of pizza orders (2015) through data preparation, pivot-table modeling, and an interactive dashboard with slicers and a timeline filter.


  ## 🎯 Project Overview 

The goal is to answer practical sales questions for a pizza business:

- How much revenue did we make, from how many orders and customers?
- Which pizza categories, sizes and individual pizzas drive revenue?
- Which months and weekdays are busiest?

### Headline results

| Metric | Value |
|---|---|
| Total sales | $8,436.05 |
| Total orders | 499 |
| Total customers (counted per order) | 499 |
| Total quantity (pizzas sold) | 511 |
| Average sales per order | $16.91 |
| Period covered | 1 Jan 2015 – 28 Dec 2015 |

## 📊 Dataset
The dataset contains pizza order information, including:
- Order ID
- Pizza ID
- Customer Name
- Quantity
- Order Time
- Order Date
- Pizza Size
- Pizza Category
- Pizza Ingredients
- Pizza Name

## 1. Data preparation & cleaning

**Microsoft Excel**
- Data preparation and calculated columns.
- Excel formulas (Sales calculation and WEEKDAY/TEXT functions).
- PivotTables for data summarization and analysis.
- Charts and visualizations.
- Interactive dashboard with slicers and filters.
```excel
month    =TEXT([@[order_date (تاريخ الطلب)]], "mmmm")
week day =TEXT([@[order_date (تاريخ الطلب)]], "dddd")
```

Before modeling, the data was **validated** to confirm it was clean and consistent:

## 2. Data Modeling with Pivot Tables

`Table1` is loaded into the Excel Data Model, so every pivot, slicer and the timeline share one source and support distinct counts. All pivots live on the `Pivot Table` sheet.

<img width="1053" height="387" alt="Pivot Table" src="https://github.com/user-attachments/assets/cda1cc18-d1b0-4e4d-be37-6a4813449682" />


| Pivot | Location | Rows | Values | Used by |
|---|---|---|---|---|
| Total sales | B3 | – | Sum of Sales | KPI card |
| Total customers | B6 | – | Count of Customer Name | KPI card |
| Total orders | B9 | – | Count of order_id | KPI card |
| Total quantity | B12 | – | Sum of quantity | KPI card |
| Sales by category | E3:F8 | pizza_category | Sum of Sales | 3D pie chart |
| Sales by size | H3:I8 | pizza_size | Sum of Sales | Doughnut chart |
| Orders by month | E11:F24 | month | Count of order_id | Line chart |
| Orders by weekday | H11:I19 | week day | Count of order_id | Column chart |
| Top 5 pizzas | M3:N9 | pizza_name (Top 5 by Sales) | Sum of Sales | Bar chart |
| Bottom 5 pizzas | M11:N17 | pizza_name (Bottom 5 by Sales) | Sum of Sales | Bar chart |

### Key findings

**Category**

| Category | Sales | Share |
|---|---|---|
| Veggie | $2,365.25 | 28.0% |
| Classic | $2,274.75 | 27.0% |
| Supreme | $2,052.05 | 24.3% |
| Chicken | $1,744.00 | 20.7% |

**Size**

| Size | Sales | Share | Order lines |
|---|---|---|---|
| L | $4,004.50 | 47.5% | 197 |
| M | $2,262.25 | 26.8% | 138 |
| S | $1,888.80 | 22.4% | 153 |
| XL | $280.50 | 3.3% | 11 |

**Top 5 pizzas by sales**: The Classic Deluxe Pizza ($459.00), The Four Cheese Pizza ($432.75), The Thai Chicken Pizza ($428.50), The Greek Pizza ($402.00), The Barbecue Chicken Pizza ($383.75).

**Bottom 5 pizzas by sales**: The Brie Carre Pizza ($47.30), The Spinach Supreme Pizza ($74.75), The Chicken Alfredo Pizza ($87.75), The Pepperoni, Mushroom, and Peppers Pizza ($122.50), The Calabrese Pizza ($166.50).

**Orders by month**: Jun 54 · Aug 51 · May / Jul / Sep 47 · Dec 46 · Oct 44 · Apr / Nov 34 · Jan 33 · Feb 32 · Mar 30. Summer is the peak; late winter is the slowest.

**Orders by weekday**: Fri 91 · Sat 77 · Thu 73 · Tue 72 · Sun 72 · Mon 66 · Wed 48. Friday is the busiest day and Wednesday the quietest.

## 3. Dashboard (Source)

The dashboard sheet is built entirely from the pivots above. Every KPI is a `GETPIVOTDATA` formula and every chart is a PivotChart.

<img width="1312" height="422" alt="Pizza Sales Dashboard" src="https://github.com/user-attachments/assets/f082eba8-85c5-469f-bd90-7b077447959a" />


### Layout

```text
PIZZAS SALES DASHBOARD
│
├── KPI row (formulas on the 'Pivot Table' sheet)
│   ├── Total Sales      =GETPIVOTDATA("[Measures].[Sum of Sales]",'Pivot Table'!$B$3)
│   ├── Total customers  =GETPIVOTDATA("[Measures].[Count of Customer Name (اسم العميل)]",'Pivot Table'!$B$6)
│   ├── Total Orders     =GETPIVOTDATA("[Measures].[Count of order_id]",'Pivot Table'!$B$9)
│   └── Total Quantity   =GETPIVOTDATA("[Measures].[Sum of quantity]",'Pivot Table'!$B$12)
│
├── Total sales by category        [3D Pie]            'Pivot Table'!E3:F8
├── Monthly order volume peaks     [Line + markers]    'Pivot Table'!E11:F24
├── Top 5 best-selling pizzas      [Horizontal bar]    'Pivot Table'!M3:N9
├── Sales distribution by size     [Doughnut]          'Pivot Table'!H3:I8
├── Weekly order distribution      [Column]            'Pivot Table'!H11:I19
├── Bottom 5 selling pizzas        [Horizontal bar]    'Pivot Table'!M11:N17
│
└── Filters (connected to all pivots)
    ├── Slicer:   pizza_category
    ├── Slicer:   pizza_size
    └── Timeline: order_date
```

### Build steps

1. Convert the data to a Table (`Ctrl+T`) and name it `Table1`.
2. Add the `month` and `week day` formula columns.
3. Insert → PivotTable → tick **Add this data to the Data Model**.
4. Create the 10 pivots listed in Section 2 on the `Pivot Table` sheet. For Top / Bottom 5, use *Value Filters → Top 10* on `pizza_name` by Sum of Sales.
5. Insert one PivotChart per pivot on the `Pizzas Sales Dashboard` sheet and style them with the brown / gold theme.
6. Add the four KPI cards using the `GETPIVOTDATA` formulas above.
7. Insert two slicers (`pizza_category`, `pizza_size`) and one timeline (`order_date`), then use **Report Connections** to connect each of them to all pivots.

## Repository Structure

```text
pizza-sales-analysis-dashboard/
├── README.md
├── excel/
│   └── Data_of_Pizza_Sales.xlsx
└── screenshots/
    ├── dashboard.png
    └── pivot_tables.png
```

## Tools

- Microsoft Excel: Tables, formulas, Data Model, PivotTables, PivotCharts, Slicers, Timeline

## Author

"Mahmoud Abdelhamed"
Linkedin : https://www.linkedin.com/in/mahmoud-abd-elhamed-8bb015238/
Gmail : melgamasy49@gmail.com
