# My-Excel-Projects
# 📊 Data Technician Excel Bootcamp

## 🚀 Project Overview

Welcome to my **Data Technician Excel Portfolio**.

This repository contains Excel tasks completed as part of my Data Technician training. The projects show practical spreadsheet skills including filtering, sorting, formulas, pivot tables, data categorisation, and basic data visualisation.

The work uses retail sales and bike sales datasets to explore sales performance, commission calculations, product categories, customer demographics, and revenue trends.

---

## 📁 Files Included

| File Name                                                 | Description                                                                                        |
| --------------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| `Data_Technician_Workbook_Week_1_Excel 2026 v2-2.docx`    | Completed workbook with task evidence, screenshots, formulas, pivot table outputs, and reflections |
| `Zakariya_retail_sales_dataset_Master 2026 - Copy 1.xlsx` | Retail sales dataset used for filtering, sorting, commission calculations, and formula practice    |
| `Bike_Sales_Pivot_Lab.xlsx`                               | Bike sales dataset used to create pivot tables and analyse sales by country, age group, and gender |
| `Bike_Sales_Visualizations_Lab.xlsx`                      | Dataset used to create visual summaries and analyse revenue by country and product category        |
| `Biggest Atlantic Hurricanes -1 - Yusuf.xlsx`             | Additional Excel dataset included in the project files                                             |

---

## 🛠️ Skills Demonstrated

* 🔎 Sorting and filtering data in Excel
* 🧮 Using formulas such as `SUM`, `AVERAGE`, and multiplication formulas
* ➕ Creating calculated columns
* 🔗 Using `CONCATENATE`
* 📊 Creating Pivot Tables
* 📌 Summarising data by categories
* 🧠 Using the `SWITCH` function to classify sales volume
* 📈 Analysing sales and revenue trends
* 📉 Creating basic Excel visualisations
* 📝 Interpreting data and writing findings

---

## 🛒 Retail Sales Analysis

The retail sales task focused on working with transaction data and commission calculations.

### ✅ Tasks Completed

* Filtered the `Age` column from largest to smallest
* Restored the original order by sorting by `Transaction ID`
* Calculated total commission using the `SUM` function
* Calculated average commission using the `AVERAGE` function
* Created additional analysis using formulas and concatenation

### 🧮 Example Formulas Used

```excel
=M2*O$2
=SUM(L1:L1001)
=AVERAGE(L2:L1001)
```

### 🔍 Key Findings

| Metric             |  Result |
| ------------------ | ------: |
| Commission Total   | 2708.95 |
| Average Commission |    2.70 |

---

## 🚲 Bike Sales Pivot Table Analysis

The Bike Sales Pivot Lab was used to practise creating pivot tables and analysing sales data across countries, age groups, and gender.

### ❓ Questions Answered

**In which age group does Germany have customers?**

Germany has customers in:

```text
Adults (35–64)
```

**What country has sales in all markets?**

The countries identified as having sales in all age-group markets were:

```text
Australia and the United Kingdom
```

**What is the most profitable country by age group and gender?**

The United States had the highest total sales quantity overall:

```text
66
```

---

## 🏷️ Product Sales Categorisation

A separate task used a small dataset containing counties, products, and sales volume. The objective was to create a pivot table and categorise sales volume using the `SWITCH` function.

### 📌 Dataset Fields

* County
* Product
* Sales Volume
* Category

### 🧠 SWITCH Formula Used

```excel
=SWITCH(TRUE, C2 > 600, "High", C2 >= 300, "Medium", "Low")
```

### 📊 Sales Volume Categories

| Sales Volume     | Category |
| ---------------- | -------- |
| Greater than 600 | High     |
| 300 to 600       | Medium   |
| Less than 300    | Low      |

### 🗂️ Example Categorised Data

| County             | Product     | Sales Volume | Category |
| ------------------ | ----------- | -----------: | -------- |
| Yorkshire          | Laptops     |          500 | Medium   |
| Yorkshire          | Smartphones |          200 | Low      |
| Cornwall           | Laptops     |          700 | High     |
| Cornwall           | Printers    |          400 | Medium   |
| Lancashire         | Smartphones |          150 | Low      |
| Lancashire         | Laptops     |          600 | Medium   |
| Essex              | Printers    |          800 | High     |
| Essex              | Smartphones |          300 | Medium   |
| Durham             | Laptops     |          250 | Low      |
| Durham             | Printers    |          300 | Medium   |
| Greater Manchester | Smartphones |          600 | Medium   |
| Greater Manchester | Laptops     |          400 | Medium   |

---

## 📈 Bike Sales Revenue Visualisation

The Bike Sales Visualizations Lab analysed revenue by country and product category.

### 🌍 Revenue by Country and Product Category

| Country         |    Accessories |          Bikes |      Clothing |    Grand Total |
| --------------- | -------------: | -------------: | ------------: | -------------: |
| Australia       |      3,284,787 |     20,231,486 |     1,911,313 |     25,427,586 |
| Canada          |      2,305,298 |      4,317,696 |     1,391,542 |      8,014,536 |
| France          |      1,627,689 |      7,378,349 |       841,175 |      9,847,213 |
| Germany         |      1,724,549 |      7,544,500 |       713,154 |      9,982,203 |
| United Kingdom  |      1,951,000 |      8,184,668 |       954,338 |     11,090,006 |
| United States   |      5,819,323 |     21,551,497 |     3,443,954 |     30,814,774 |
| **Grand Total** | **16,712,646** | **69,208,196** | **9,255,476** | **95,176,318** |

### 🔍 Key Insights

* 🇺🇸 The **United States** generated the highest total revenue.
* 🚲 **Bikes** were the highest revenue-generating product category.
* 🇦🇺 **Australia** also performed strongly, especially in bike sales.
* 👕 **Clothing** generated the lowest revenue compared with Bikes and Accessories.
* 💰 Overall revenue across all countries and product categories was **95,176,318**.

---

## 🧰 Tools Used

* Microsoft Excel
* Pivot Tables
* Excel formulas and functions
* Data filtering and sorting
* Basic data visualisation tools

---

## 🎯 What I Learned

Through this project, I developed practical Excel skills for cleaning, analysing, and summarising data.

I learned how to:

* Use formulas to calculate totals and averages
* Create pivot tables to summarise datasets
* Categorise data using logical functions such as `SWITCH`
* Analyse sales and revenue trends
* Interpret spreadsheet results and identify business insights

---

## 🌱 Future Improvements

In future versions of this project, I could improve the analysis by:

* 📊 Adding more charts and dashboards
* 🧹 Cleaning and formatting datasets more consistently
* 🎛️ Creating interactive slicers for pivot tables
* 🖼️ Adding screenshots of completed Excel work
* 📌 Including a final dashboard summary
* 📆 Comparing sales trends across different time periods

---

## 👤 Author

**Zakariya Omar**

📊 Data Technician Excel Bootcamp

