# 🛒 Retail Store Performance Dashboard

An interactive Excel dashboard for analyzing retail store sales, profit, and business performance — built as a Module Final Project.

---

## 📌 Project Overview

| Detail | Info |
|--------|------|
| **Project Title** | Retail Store Performance Dashboard |
| **Tool Used** | Microsoft Excel |
| **Module** | Computer Basic and Component |
| **Submitted By** | Soni |
| **Date** | December 20, 2025 |

---

## 🎯 Objective

To analyze retail store transactional data and build a professional, interactive Excel dashboard that helps in understanding business performance and supports faster decision-making.

---

## 📁 File Structure

```
Retail_Store_Performance_Dashboard.xlsx
│
├── Raw_Data          → Original transactional dataset
├── KPI_Calculation   → Calculated Key Performance Indicators
├── Pivot_Analysis    → Pivot tables for all analyses
└── Dashboard         → Final interactive dashboard
```

---

## 📊 Dataset Details

**Sheet:** `Raw_Data`

The dataset contains retail store transactions with the following columns:

`Order_ID` | `Order_Date` | `Store_Name` | `Region` | `Product_Category` | `Product_Name` | `Quantity` | `Sales` | `Cost` | `Profit` | `Payment_Mode` | `Customer_Gender`

> **Profit Formula:** `Profit = Sales - Cost`

The raw data was converted into an Excel Table (`Ctrl + T`) named **Retail_Table**.

---

## 🧹 Data Cleaning Steps

- Removed blank rows
- Applied correct data types (Date, Currency)
- Removed duplicate records
- Validated all numeric columns

**Excel Tools Used:** Filter, Remove Duplicates, Format Cells

---

## 📈 KPI Metrics

| KPI | Formula |
|-----|---------|
| Total Sales | `SUM(Sales)` |
| Total Profit | `SUM(Profit)` |
| Total Orders | `COUNTA(Order_ID)` |
| Total Quantity | `SUM(Quantity)` |
| Average Order Value | `Total Sales / Total Orders` |

---

## 🔄 Pivot Table Analysis

| Analysis | Rows | Values |
|----------|------|--------|
| Sales by Region | Region | Sum of Sales |
| Sales by Category | Product_Category | Sum of Sales |
| Profit by Store | Store_Name | Sum of Profit |
| Payment Mode | Payment_Mode | Sum of Sales |

---

## 📉 Charts Created

| Analysis | Chart Type |
|----------|------------|
| Sales by Region | Column Chart |
| Sales by Category | Pie Chart |
| Profit by Store | Bar Chart |
| Payment Mode | Doughnut Chart |

---

## 🖥️ Dashboard Features

- **KPI Cards** — Dynamic cards showing Total Sales, Profit, Orders, Quantity, and Avg Order Value
- **2×2 Chart Grid** — All 4 charts arranged neatly
- **Slicers** for interactive filtering:
  - Region
  - Store Name
  - Product Category
  - Payment Mode
- All slicers connected to all pivot tables via **PivotTable Connections**
- Clean design: no gridlines, no chart field buttons, consistent font & color theme

---

## ✅ How to Use

1. Open `Retail_Store_Performance_Dashboard.xlsx` in Microsoft Excel
2. Go to the **Dashboard** sheet
3. Use the **slicers** on the right/top to filter data by Region, Store, Category, or Payment Mode
4. KPI cards and charts will update automatically

---

## 💡 Key Learnings

- Excel Table creation and structured references
- Pivot Tables and Pivot Charts
- KPI card design using Shapes + Formulas
- Slicer connections for full interactivity
- Professional dashboard layout and formatting

---

## 🏁 Conclusion

This project demonstrates practical skills in Excel-based data analysis and dashboard design. It is suitable for academic submission as well as interview presentations.

---

> **Note:** Open in Microsoft Excel (2016 or later) for best experience with slicers and interactive features.
