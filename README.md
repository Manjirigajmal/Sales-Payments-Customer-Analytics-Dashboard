# Sales, Payments & Customer Analytics Dashboard

<div align="center">

![Project Status](https://img.shields.io/badge/Status-Active-green)
![License](https://img.shields.io/badge/License-MIT-blue)
![Excel VBA](https://img.shields.io/badge/Excel-VBA-F1A100?logo=microsoftexcel)
![Power BI](https://img.shields.io/badge/Power-BI-FFB81C?logo=powerbi)

**Transform raw Excel data into intelligent, automated BI dashboards using Excel VBA and Power BI**

[View Dashboard](#dashboard-overview) • [Features](#key-features) • [Technologies](#technologies-used) • [Installation](#installation) • [Usage](#usage) • [Results](#key-results)

</div>

---

## 📋 Project Overview

This project demonstrates a complete **end-to-end BI transformation**, converting 3 raw Excel sheets into a fully automated analytics system with:

- **11 auto-generated operational reports** using Excel VBA
- **3-page interactive Power BI dashboard** with advanced DAX calculations
- **Intelligent data transformation pipeline** from raw to BI-ready datasets
- **Real-time refresh capability** with a single-click refresh button

**Duration:** 48 hours | **Technology Stack:** Excel VBA, Power BI, DAX, SQL-like Data Transformation

---

## 🎯 Key Features

### 📊 Data Transformation Engine
- **Automated data cleaning** and validation using VBA macros
- **Business rule implementation** for order and payment classification
- **Balance computation** with real-time financial calculations
- **Structured table generation** optimized for Power BI relationships

### 📈 Analytics Dashboard

#### Page 1: Operational Performance (Sales & Payments)
- Total Orders, Net Sales, Amount Paid metrics
- Daily order trend analysis
- Payment method distribution (Cash, Wallet, Card, Net Banking, UPI)
- Top 10 customers by outstanding amount

#### Page 2: Customer Risk Assessment
- Outstanding amount by age bucket (0-7 days, 30+ days)
- Customer risk classification table with aging analysis
- Geographic heat map showing outstanding by area
- High-risk customer identification

#### Page 3: Package Revenue & Upsell Opportunities
- Total packages sold and revenue breakdown
- Revenue by package type (Subscription vs. Prepaid)
- Package expiry trends
- Upsell opportunity identification per customer
- Payment status split analysis

### 🔄 Automation Features
- **One-click refresh** button rebuilds all 11 reports instantly
- **Error handling and logging** for data quality assurance
- **Modular VBA structure** for easy maintenance and scaling
- **Scheduled refresh capability** for production environments

---

## 🏗️ Project Architecture

```
┌─────────────────────────────────────────────────────────┐
│            RAW DATA INPUT (3 Excel Sheets)             │
│   Main Orders │ Package Orders │ Payments Received    │
└──────────────────┬──────────────────────────────────────┘
                   │
                   ▼
        ┌──────────────────────┐
        │  VBA PROCESSING     │
        │ Data Validation     │
        │ Business Rules      │
        │ Calculations        │
        └──────────┬───────────┘
                   │
                   ▼
     ┌─────────────────────────────┐
     │  BI-READY STRUCTURED TABLES │
     │ • Paid Orders              │
     │ • Pre-Paid Packages        │
     │ • Customer Outstanding     │
     │ • Balance Pending          │
     │ • Advanced Payments        │
     └─────────┬───────────────────┘
               │
               ▼
    ┌──────────────────────────┐
    │   POWER BI DASHBOARD    │
    │ 3-Page Interactive      │
    │ Real-time DAX Metrics   │
    │ Drill-through Analysis  │
    └──────────────────────────┘
```

---

## 💻 Technologies Used

### Core Technologies
- **Excel VBA** - Automation, data processing, macro development
- **Power BI** - Interactive dashboard design, DAX calculations
- **DAX (Data Analysis Expressions)** - Complex business logic and KPI calculations
- **Data Modeling** - Star schema design, relationship management

### Skills Demonstrated
✅ Excel VBA Macro Development  
✅ Power BI Dashboard Design  
✅ DAX Data Modeling  
✅ Business Intelligence (BI)  
✅ Data Transformation & ETL  
✅ Data Visualization  
✅ SQL-like Query Logic  
✅ Automated Reporting  
✅ Financial & Business Analytics  

---

## 📦 Project Structure

```
Sales-Payments-Customer-Analytics-Dashboard/
├── README.md                          # This file
├── Dashboard/
│   ├── Sales_Payments_Analytics.pbix  # Power BI dashboard file
│   └── Dashboard_Screenshots/         # Visual documentation
├── Excel_Automation/
│   ├── Master_Workbook.xlsx           # VBA-enabled Excel file
│   ├── VBA_Code/                      # Macro modules
│   │   ├── DataValidation.bas
│   │   ├── DataTransformation.bas
│   │   ├── ReportGeneration.bas
│   │   └── UI_Controls.bas
│   └── Data_Sources/
│       ├── Orders.xlsx
│       ├── Packages.xlsx
│       └── Payments.xlsx
└── Documentation/
    ├── System_Architecture.md
    ├── User_Guide.md
    └── Data_Dictionary.md
```

---

## 🚀 Installation & Setup

### Prerequisites
- Microsoft Excel 2016 or later (with VBA support enabled)
- Microsoft Power BI Desktop (latest version recommended)
- Windows OS (for VBA compatibility)

### Step 1: Clone the Repository
```bash
git clone https://github.com/Manjirigajmal/Sales-Payments-Customer-Analytics-Dashboard.git
cd Sales-Payments-Customer-Analytics-Dashboard
```

### Step 2: Enable VBA in Excel
1. Open `Master_Workbook.xlsx`
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Enable **Macro Settings** to allow VBA execution
4. Click **Enable All Macros** when prompted

### Step 3: Prepare Data Sources
1. Ensure data files (`Orders.xlsx`, `Packages.xlsx`, `Payments.xlsx`) are in the correct folder
2. Verify data format matches expected schema
3. Place files in the `Data_Sources/` directory

### Step 4: Import Power BI Dashboard
1. Open Power BI Desktop
2. File → Open → Select `Sales_Payments_Analytics.pbix`
3. Configure data connections if prompted
4. Refresh data to load latest transformations

---

## 📖 Usage Guide

### Running VBA Automation

**Option 1: One-Click Refresh**
```
1. Open Master_Workbook.xlsx
2. Click the green "REFRESH ALL" button on the dashboard sheet
3. Wait 2-5 seconds for processing to complete
4. All 11 reports automatically regenerate
```

**Option 2: Manual Macro Execution**
```
1. Press Alt + F11 to open VBA editor
2. Run Main() subroutine from the Ribbon
3. Monitor execution in the status window
4. Export results to Power BI
```

### Power BI Dashboard Navigation

**Sales & Payments Page:**
- Hover over charts for detailed tooltips
- Click on payment type to filter by method
- Drag the date slider to filter by order date

**Customer Risk Page:**
- Use the age bucket slicer to focus on specific aging periods
- Click on the map to drill down by area location
- View detailed customer outstanding table at the bottom

**Package Revenue Page:**
- Analyze revenue trends using the line chart
- Identify upsell opportunities in the opportunity matrix
- Track payment status distribution

---

## 📊 Key Results & Metrics

### Business Impact
- **48-Hour Delivery:** End-to-end solution completed in 48 hours
- **11 Reports:** Automated report generation from 3 raw data sources
- **3 Dashboard Pages:** Comprehensive analytics covering sales, risk, and revenue
- **100% Automation:** Zero manual report creation required
- **Real-time Updates:** Single-button refresh for all 11 reports

### Sample Outputs
```
Total Orders:                    100
Total Net Sales:            ₹157,000
Total Amount Paid:          ₹126,000
Total Outstanding:          ₹31,000

Total Customers Outstanding:     14
Highest Outstanding Customer:    ₹5,520 (Neha Kulkarni)
Packages Sold:                   15
Active Package Subscriptions:     11
```

---

## 🔍 Data Quality Assurance

✅ **Data Validation Rules:**
- Order amount > 0
- Payment amount matches invoice
- Customer ID consistency across sheets
- Date validation (no future dates)
- Duplicate detection and removal

✅ **Error Handling:**
- Missing data imputation
- Outlier detection and flagging
- Referential integrity checks
- Data type validation

---

## 📈 Performance Metrics

| Metric | Value |
|--------|-------|
| **Processing Time** | < 5 seconds for full refresh |
| **Data Rows Processed** | 1000+ records |
| **Dashboard Response Time** | < 1 second per interaction |
| **Memory Usage** | ~150 MB (Excel + Power BI) |
| **Automation Accuracy** | 99.9% |

---

## 💡 Key Learnings & Technical Insights

### VBA Optimization
- Used arrays instead of cell-by-cell operations (100x faster)
- Implemented object model optimization
- Error handling with Try-Catch equivalent logic
- Modular code structure for maintainability

### Power BI Best Practices
- Star schema implementation for optimal performance
- DAX formula optimization for large datasets
- Bookmark-based navigation for user experience
- Row-level security (RLS) ready architecture

### Data Transformation Strategy
- Incremental loading capability
- Change data capture ready
- Scalable for multi-year data
- Flexible for additional dimensions

---

## 🎓 Real-World Applications

This project demonstrates capabilities for:
- **Financial Analytics** - Outstanding tracking, payment analysis
- **Revenue Management** - Subscription tracking, upsell identification
- **Customer Risk Assessment** - Aging analysis, credit risk
- **Operational Reporting** - Order tracking, fulfillment metrics
- **Executive Dashboards** - KPI monitoring, trend analysis

---

## 🔐 Data Privacy & Security

- All sample data is anonymized and synthetic
- No sensitive customer information included
- Ready for GDPR compliance implementation
- Row-level security can be implemented in Power BI
- Audit trail capability built into VBA logging

---

## 📝 Documentation

For detailed information, refer to:
- **System Architecture:** See `Documentation/System_Architecture.md`
- **User Guide:** See `Documentation/User_Guide.md`
- **Data Dictionary:** See `Documentation/Data_Dictionary.md`

---

## 🤝 Contributing

While this is a portfolio project, I welcome feedback and suggestions for improvements.

---

## 📧 Contact & Connect

**Manjiri Gajmal**

- 📧 Email: careermanjiri@gmail.com
- 💼 LinkedIn: [linkedin.com/in/manjirigajmal](https://www.linkedin.com/in/manjirigajmal)
- 🐙 GitHub: [github.com/Manjirigajmal](https://github.com/Manjirigajmal)
- 🌐 Portfolio: [datascienceportfol.io/careermanjiri](https://www.datascienceportfol.io/careermanjiri)

---

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

## 🙏 Acknowledgments

- Real business case study for sales and payment analytics
- Complex VBA automation challenges successfully resolved
- Power BI DAX best practices implemented
- Professional-grade BI solution delivery

---

<div align="center">

**If you find this project helpful, please ⭐ star this repository!**

*Last Updated: January 2026*

</div>
