# ☕ Coffee Sales Dashboard (Excel Project)

This project outlines the end-to-end creation of a **dynamic, interactive Coffee Sales Dashboard** built entirely in **Microsoft Excel**. The process covers data gathering, transformation, and visualization using pivot tables, pivot charts, timelines, and slicers.

The goal of this dashboard is to provide analysts and stakeholders with clear, actionable insights into coffee sales performance over time, geographical distribution, and key customer segments.

---

 ## 📈 Dashboard Preview
 <img src="coffeeSalesDashboard.png">

---

## 📊 Dashboard Features & Visualizations

The final dashboard includes the following key components, all connected and interactive:

* **Total Sales Over Time:** A line chart visualizing monthly total sales, segmented by the four different coffee types (Arabica, Excelsa, Liberica, Robusta).
* **Sales By Country:** A horizontal bar chart showing sales distribution across the three primary markets (U.S., Ireland, UK).
* **Top 5 Customers:** A bar chart highlighting the top customers based on total sales volume.
* **Interactive Filters:**
    * **Timeline:** Allows filtering all visuals by specific time periods (Year/Month).
    * **Slicers:** Enables dynamic filtering by:
        * Roast Type Name (Dark, Light, Medium)
        * Size (0.2kg, 0.5kg, 1.0kg, 2.5kg)
        * Loyalty Card (Yes/No)

---

## 🛠️ Technologies Used

* **Microsoft Excel:** Used for all data manipulation, analysis, and visualization.
* **Formulas:** `XLOOKUP`, `INDEX MATCH`, `IF` (nested) for data transformation.
* **Tools:** Pivot Tables, Pivot Charts, Timelines, and Slicers.

---

## ⚙️ Project Walkthrough (Step-by-Step)

The project follows a standard data pipeline process:

### 1. Data Gathering & Transformation

* **Data Structure:** The project begins with three separate data sheets: `Orders`, `Customers`, and `Products`.
* **Data Lookup:**
    * Used `XLOOKUP` to bring customer information (Name, Email, Country, Loyalty Card status) from the `Customers` sheet into the `Orders` sheet.
    * Used the powerful `INDEX MATCH` combination to dynamically retrieve all product details (Coffee Type, Roast Type, Size, Unit Price) from the `Products` sheet using a single formula that could be dragged across multiple columns.
* **Calculated Fields:** Created the **Sales** column by multiplying `Unit Price` by `Quantity Sold`.
* **Data Cleaning:** Used nested `IF` functions to convert abbreviations (e.g., 'rob' to 'Robusta', 'M' to 'Medium') into full, readable names for the Coffee Type and Roast Type columns.
* **Formatting:** Applied custom date (`dd-mmm-yyyy`) and number formatting (USD currency, size in `0.0 kilo`).
* **Table Conversion:** Converted the raw data range into an official Excel **Table** (named `Orders Table`) to ensure pivot tables automatically capture new data additions (like the new Loyalty Card column).

### 2. Analysis and Visualization Setup

* **Pivot Table Creation:** Started the analysis by creating the first pivot table (Total Sales), using the `Alt + N + V + T` shortcut.
* **Grouping:** Grouped the `Order Date` field by **Years and Months** for time-series analysis.
* **Pivot Chart Creation:**
    * **Total Sales:** Created the main **Line Chart** using Order Date (Rows), Coffee Type Name (Columns), and Sum of Sales (Values).
    * **Sales by Country:** Created a **Bar Chart** using Country and Sum of Sales, then sorted by sales amount (ascending) to present the top performer at the top of the chart.
    * **Top 5 Customers:** Applied a **Value Filter (Top 10... -> Top 5)** on the Customer Name field to isolate the key segment, then created the corresponding Bar Chart.
* **Styling:** Applied a custom, theme-consistent color scheme (based on deep purples and warm browns in the video) and hid field buttons for a cleaner look.

### 3. Interactivity and Dashboard Finalization

* **Inserting Controls:** Inserted a **Timeline** (linked to Order Date) and three **Slicers** (Roast Type Name, Size, Loyalty Card). The new Loyalty Card column was automatically included after refreshing the pivot table data source.
* **Connecting Controls:** Used the **Report Connections** feature for the Timeline and all Slicers to ensure they filter **all three pivot tables/charts** simultaneously, making the dashboard fully dynamic.
* **Layout:** Created the final **Dashboard** sheet with a clean header and background, then cut (`Ctrl+X`) and pasted (`Ctrl+V`) all charts, the timeline, and the slicers onto the canvas, adjusting their size and position for optimal viewing.

---

## 🚀 Getting Started

If you wish to recreate this project:

1.  Download the [coffeeOrdersRawData.xlsx](https://github.com/KavyaEkanayake/excel-project-coffee-sales/blob/main/coffeeOrdersRawData.xlsx) sample data file.
2.  Follow the steps outlined in the Project Walkthrough.
