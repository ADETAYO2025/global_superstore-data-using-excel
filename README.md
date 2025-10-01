# global_superstore-data-using-excel
## 📝 Project Summary

This project focused on analyzing a comprehensive sales dataset using Microsoft Excel. The goal was to clean, transform, and extract insights from the data using advanced Excel features including formulas, pivot tables, slicers, and charts.

---

## 🧹 Data Preparation & Cleaning

- Removed duplicate entries and blank rows
- Formatted date fields consistently (`Order Date`, `Ship Date`)
- Validated numerical columns such as `Sales`, `Profit`, `Quantity`, and `Shipping Cost`
- Ensured consistent text formatting for `Category`, `Sub-Category`, `Region`, and `Order Priority`

---

## 🧾 Dataset Overview

The dataset consists of detailed sales transaction records with the following key variables:

- **Order Info:** `Row ID`, `Order ID`, `Order Date`, `Ship Date`, `Ship Mode`, `Order Priority`
- **Customer Info:** `Customer ID`, `Customer Name`, `Segment`, `Postal Code`, `City`, `State`, `Country`, `Region`, `Market`
- **Product Info:** `Product ID`, `Category`, `Sub-Category`, `Product Name`
- **Sales Info:** `Sales`, `Quantity`, `Discount`, `Profit`, `Shipping Cost`

---

## 🛠 Excel Skills Demonstrated

### 1. 🔗 Cell Referencing

- Used **relative referencing** for dynamic formulas across rows
- Used **absolute referencing** (e.g., `$A$2`) to lock ranges in functions like `SUMIF`, `SUMIFS`
  
  **Example:**
  ```excel
  =SUMIFS($S$2:$S$1000, $L$2:$L$1000, "West", $O$2:$O$1000, "Technology")

  2. 🧮 Basic Functions & Formulas
SUM(): Total sales, profit, and shipping cost
SUMIF() / SUMIFS(): Conditional sales and profit summaries
AVERAGEIF(): Average sales by category or region
Created a User-Defined Function (UDF) in VBA to convert currency (e.g., to USD)
UDF Example:
Function ConvertToUSD(amount As Double, rate As Double) As Double
    ConvertToUSD = amount * rate
End Function
3. 📊 Pivot Table & Visualization
Built Pivot Tables to analyze:
Total Sales and Profit per Category and Sub-Category
Sales trends by Region and Market
Applied Slicers for interactive filtering:
Order Date
Region
Category
Inserted a Bar Chart to visually represent Sales by Region
📈 Insights Uncovered
Identified top-performing Sub-Categories by Profit
Detected regions with highest Shipping Cost-to-Profit ratios
Analyzed seasonal trends using Order Date slicer
Compared Sales distribution across Markets and Customer Segments
📂 Tools & Technologies Used
Microsoft Excel
Excel Formulas & Functions
Pivot Tables
Slicers
Charts (Bar Chart)
VBA (User-Defined Function)
🚀 Outcome
By leveraging Excel's analytical tools, I created a dynamic and interactive dashboard that provides insights into sales performance, customer behavior, and product profitability. This showcases my ability to work with large datasets, apply analytical thinking, and present data visually and interactively.



<img width="1190" height="478" alt="image" src="https://github.com/user-attachments/assets/e7744afa-19a6-40fe-a2b0-18d23eda9bc0" />

<img width="1296" height="670" alt="image" src="https://github.com/user-attachments/assets/d5e2b299-a7e3-420b-9290-1aba158f3b26" />

<img width="1278" height="422" alt="image" src="https://github.com/user-attachments/assets/96fc8510-2473-4fde-b28b-7d7d3cd4aeb2" />



