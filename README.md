# Coffee Sales Dashboard in Excel

![Coffee Sales Dashboard](https://github.com/kartik981/Coffee-Sales-Excel-Dashboard/blob/1fc8523603aecbd68749b496693e21482d254904/Coffee%20Sales%20Dashboard.png)
## Introduction
This project focuses on creating an **interactive Coffee Sales Dashboard** in Excel using **Pivot Tables, Pivot Charts, and Slicers**. The dashboard provides insights into sales trends, customer behavior, and product performance.

## Methodology
### Data Preparation
1. **Dataset Overview:**
   - Contains **sales transactions** with details such as order date, customer ID, product ID, quantity, and prices.
   - Additional information retrieved from **Customers** and **Products** tables.

2. **Data Enrichment:**
   - **XLOOKUP Formula**: Retrieved customer name, email, and country.
   - **INDEX MATCH Formula**: Used to fetch coffee type, roast type, size, and price.
   - **Calculated Sales Column**: Computed as `Quantity × Unit Price`.
   - **IF Formulas**: Used for categorical mapping (e.g., abbreviations to full coffee names).

### Data Formatting
1. **Date Format:** Converted order dates to `dd-mmm-yyyy`.
2. **Number Formatting:**
   - **Size Column**: Displayed in kilograms (`0.0 kilo`).
   - **Currency Columns**: Unit Price and Sales formatted in **US Dollars**.
3. **Duplicate Check:** Removed duplicate records to ensure data integrity.
4. **Table Conversion:** Converted the dataset into an **Excel Table** for dynamic updates.

## Dashboard Components
### Pivot Tables and Charts
1. **Total Sales Over Time**
   - A **line chart** displaying total sales by month.
   - Segmented by **coffee type** (Arabica, Excelsa, Liberica, Robusta).

2. **Sales by Country**
   - A **bar chart** showing total sales per country.
   - Sorted in **descending order** to highlight the highest-grossing region.

3. **Top 5 Customers**
   - A **bar chart** listing the top five customers based on total sales.
   - Applied a **Top 5 Value Filter** in the pivot table.

### Interactive Elements
1. **Timeline Filter**
   - Enables dynamic filtering of all visuals based on selected time ranges.
   - Custom formatting applied for better readability.
2. **Slicers**
   - **Roast Type:** Filters by Dark, Light, or Medium roast.
   - **Size:** Filters by package sizes (0.2kg, 0.5kg, 1kg, 2.5kg).
   - **Loyalty Card:** Filters based on customer loyalty program status.
   - Connected all slicers to pivot tables for **synchronized filtering**.

## Dashboard Design
1. **Header Design:**
   - Added a **title bar** with a custom shape and background color.
   - Used **Calibri Bold** font for a professional look.
2. **Grid Removal:**
   - Disabled **grid lines and formula bar** for a cleaner presentation.
   - Removed **scrollbars** to create a dashboard-like appearance.
3. **Element Placement:**
   - Charts and slicers arranged neatly for **easy interpretation**.
   - Used **ALT key snapping** for precise alignment.

## Conclusion
The **Coffee Sales Dashboard** provides a structured and interactive way to analyze sales data. Key benefits include:
- **Automated updates** through Excel tables and pivot refresh.
- **Dynamic filtering** using slicers and timelines.
- **Clear visualization** of sales trends and customer insights.

### Future Improvements
- **Integrating Power BI** for enhanced reporting.
- **Connecting to an external database** for **real-time updates**.

## Tools Used
- **Microsoft Excel** (Pivot Tables, Pivot Charts, Slicers, Timelines)
- **Excel Functions** (XLOOKUP, INDEX MATCH, IF, SUM, Custom Formatting)
- **Data Visualization Techniques** (Line Charts, Bar Charts, Conditional Formatting)

## References
- Microsoft Office Documentation on Pivot Tables and Slicers
- Best Practices for Dashboard Design in Excel
