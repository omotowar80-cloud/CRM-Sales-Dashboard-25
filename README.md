# CRM Sales Dashboard '25

This project contains a sales CRM dataset with multiple sheets designed for performance tracking, sales analysis, and dashboarding.

## 📂 File Overview
**File:** `Copy of CRM Sales Dashboard '25 (1).xlsx`

### Sheets
1. **CRM Sales Dashboard(in).csv**
   - Main dataset of opportunities
   - Columns: `opportunity_id`, `sales_agent`, `manager`, `regional_office`, `product`, `account`, `deal_stage`, `engage_date`, `close_date`, `close_value`

2. **Pivot Table 1**
   - Currently empty, likely a placeholder for future pivot tables or dashboards.

3. **Detail2-Won-Q4**
   - Subset of opportunities that were **Won** in Q4.
   - Useful for analyzing successful deals.

4. **Detail1-Q4**
   - All Q4 opportunities, including **Won** and **Lost**.
   - Useful for calculating **Win Rate** and pipeline performance.

5. **sales_teams**
   - Mapping of sales agents → managers → regional offices.

---

## 📊 Analyses & Visualizations

### Sales by Region
![Sales by Region](sales_by_region.png)

### Deal Stage Distribution
![Deal Stage Distribution](deal_stage_distribution.png)

### Top Products by Sales
![Top Products](top_products.png)

---

## 🚀 Usage
You can use this dataset to:
- Build interactive dashboards (Excel, Power BI, Tableau, Python).
- Run time-series analysis of sales performance.
- Conduct team-level and product-level performance reviews.

---

## 🔧 Tools Recommended
- **Python (Pandas, Matplotlib, Plotly)** for analytics & visualization.
- **Power BI / Tableau** for interactive dashboards.
- **Excel Pivot Tables** for quick summaries.

---

## 📅 Notes
- Dates are in the format `YYYY-MM-DD`.
- Monetary values are stored in `close_value` (assumed in USD).
- Deal stages include at least: `Won`, `Lost`.
- 
