# Excel Sales Dashboard | Jan - Mar 2026

An interactive sales dashboard built in Microsoft Excel using Pivot Tables, Pivot Charts, a Slicer, a Timeline, a KPI card and a VBA macro.

![Dashboard](images/Excel_Sales_Dashboard_Jan-Mar_2026.png)

## Demo

Clicking the Slicer updates all three pivot tables and charts together.

![Interactive Demo](demo/Excel_Interactive_Dashboard_Pivot_Slicer_Timeline.gif)

## Features

- **3 Pivot Tables**: Sales by City, Sales by Month, Quantity by Product
- **3 Charts**: Bar chart (sales by city), line chart (monthly sales trend), pie chart (quantity by product)
- **Slicer** on Category, connected to all 3 pivot tables via Report Connections
- **Timeline** on Order Date for filtering by month
- **KPI card** showing Total Sales, linked to the pivot Grand Total
- **VBA macro** that formats a header row (bold, fill colour, borders) in one click
- Clean dashboard layout: custom title, gridlines removed, Indian currency format (₹)

## Key Insights

| Metric | Value |
|---|---|
| Total Sales | ₹33,43,700 |
| Total Quantity Sold | 309 units |
| Top City | Delhi (₹8,66,800) |
| Best Month | March (₹15,55,700) |
| Top Product by Quantity | Mouse (132 units) |

## Project Structure

```
excel-sales-dashboard/
├── README.md
├── Excel_Practice_Workbook.xlsm
├── images/
│   └── Excel_Sales_Dashboard_Jan-Mar_2026.png
└── demo/
    └── Excel_Interactive_Dashboard_Pivot_Slicer_Timeline.gif
```

## Tools & Skills Used

- Microsoft Excel: Pivot Tables, Pivot Charts, Slicers, Timelines
- Dashboard design: layout, KPI cards, number formatting
- VBA: macro recording and editing

## How to Use

1. Download `Excel_Practice_Workbook.xlsm`.
2. Open it in Excel and click **Enable Content** (required for the macro).
3. Go to the **Dashboard** sheet.
4. Use the **Slicer** (Accessories / Devices) and the **Timeline** to filter the data.
5. To run the macro: **View > Macros > View Macros > FORMAT_HEADER > Run**.

## Macro

The `FORMAT_HEADER` macro formats the header row (`A1:H1`) of the active sheet with bold text, a fill colour and borders.

```vba
Sub FORMAT_HEADER()
    Range("A1:H1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Selection.Borders.LineStyle = xlContinuous
End Sub
```

## Author

**Aryan**
[LinkedIn](https://www.linkedin.com/in/faustus-aryan-b71978224) | [GitHub](https://github.com/faustusaryan)
