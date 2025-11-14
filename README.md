# 📊 Multi-City Sales Performance Dashboard Using Excel with Automation

## Overview
This project presents a **well-structured automated Excel dashboard** to analyze multi-city sales performance. By utilizing Excel formulas, Pivot Tables, Slicers, Charts, and VBA automation, it provides **real-time KPIs and easy-to-understand reporting**.

> **Dataset:** Contains EmpCode, Sales Executive, Region, Day1–Day5 sales, and Sales Target (sourced from a public GitHub repository).

---

## Reporting Process

1. **Understanding of Data:** Analyze the dataset to identify KPIs and reporting requirements.  
2. **Building Logic as per End Results:** Define metrics and logic to achieve the desired reporting outcomes: Total Sales, Target Hit %, Away from Target %.  
3. **Applying Formula:** Implement Excel formulas to compute metrics accurately.  
4. **Creating Report:** Build Pivot Tables, charts, and slicers; automate region filtering and button-driven views using VBA.

---

## Dashboard & Button Views

The dashboard contains four interactive views, each connected to a specific Pivot Table:

- **Dashboard 1 (Top 5 Sales Executives)** → Linked to **Pivot Table 1**  
- **Dashboard 2 (Bottom 5 Sales Executives)** → Linked to **Pivot Table 2**  
- **Dashboard 3 (Target Hit % Wise Executives)** → Linked to **Pivot Table 3**  
- **Dashboard 4 (Away from Target % Wise Executives)** → Linked to **Pivot Table 4**  

**Highlights:**

- Top-performing sales executives for each region  
- Underperforming sales executives for each region  
- KPI tracking: Total Sales, Target Hit %, Away from Target %  

**Region Filtering Behavior:**

- **Checkbox ticked:** Pivot table and chart filtered for the selected city  
- **Checkbox unticked:** Pivot table and chart display data for all cities

**Dashboard Preview:**  

![Dashboard Preview](images/sales-dashboard.jpg)

---

## Key Features

- Multi-city sales performance analysis across Chennai, Delhi, Mumbai, Pune, Nagpur, Patna, Ranchi, and Surat  
- Interactive button-driven views: Top/Bottom performers, Target Hit %, Away from Target %  
- Automated region-wise highlighting for top-performing and underperforming executives  
- Dynamic region filtering using VBA  
- Executive-ready visualizations:
  - **Bar Chart:** → linked to Pivot Table 1 (Top 5 Executives)  
  - **Pie Chart:** → linked to Pivot Table 3 (Target Hit % Wise)  
  - **Line Chart:** → linked to Pivot Table 4 (Away from Target % Wise)

---

## VBA Automation

**Purpose:** Automatically connects/disconnects region slicers to Pivot Tables and toggles button-driven views.  

**Impact:**  
- Eliminates manual filtering and repetitive updates  
- Dynamically updates charts and Pivot Tables based on selected region  
- Enhances efficiency and usability for executive reporting  

> **Full VBA Code:** [Click to view VBA code](VBA_code.txt)

---

## Business Impact

- Provides **real-time insights** to support strategic decision-making  
- Highlights top and underperforming sales executives per region for optimized sales strategy  
- Reduces manual effort through automation  
- Delivers **professional, executive-ready dashboards** for business stakeholders  

---

## Getting Started

1. Clone or download this repository.  
2. Open `Multi-city-sales-dashboard.xlsx` in Excel.  
3. Use buttons to switch views; slicers automatically update Pivot Tables and charts.  

---

## Connect with Me

- **LinkedIn:** [https://www.linkedin.com/in/swatimirashi](https://www.linkedin.com/in/swatimirashi)  
- **Email:** swatimirashi298@gmail.com  

---

## Tags

#Excel #VBA #Automation #PivotTables #Dashboard #SalesAnalytics #BusinessIntelligence #DataVisualization #KPITracking #CorporateReporting
