# Retail Customer Analytics Dashboard | Excel

> **Interactive Excel dashboard for customer demographics, data quality, geographic distribution, and customer segmentation insights.**

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-Dashboard-217346?logo=microsoftexcel)](https://www.microsoft.com/microsoft-365/excel)
[![Analytics](https://img.shields.io/badge/Focus-Customer%20Analytics-blue)](#business-problem)
[![Dataset](https://img.shields.io/badge/Dataset-336%20customers-orange)](#dataset)

## Business Problem

Retail teams need a simple way to understand **who their customers are, where they are located, and how demographic patterns can support targeting decisions**.

This project builds an Excel-based analytical dashboard that converts a raw customer dataset into an interactive management view covering:

- Customer demographics
- Age-group distribution
- Gender mix
- City-level customer concentration
- Data-quality and missing-value treatment
- Demographic intersections for customer targeting

## What I Built

```text
Raw Customer Data
       |
       v
Data Quality Checks
       |
       +--> Missing-age identification
       +--> Gender-level average calculation
       +--> Age imputation
       +--> Age-group classification
       |
       v
Pivot Tables + Excel Formulas
       |
       v
Interactive Dashboard
       |
       v
Customer & Marketing Insights
```

## Dataset

The project uses a synthetic retail customer dataset containing **336 customer records** across major Indian cities.

Key fields include:

| Field | Purpose |
|---|---|
| `Customer_ID` | Unique customer identifier |
| `Customer_Name` | Anonymous customer label |
| `Age_Old` | Original age value, including missing records |
| `Age` | Cleaned/imputed age |
| `AgeGroup` | `<25`, `25-40`, `40-60`, `60+` |
| `Gender` | Female, Male, Other |
| `City` | Customer location |

### Data-quality treatment

Missing ages are handled using **gender-level average age**, with a fallback to the overall average where required. The resulting age field is then categorized into business-friendly age groups.

This makes the dashboard useful not only for visualization but also as an example of **data cleaning and feature engineering in Excel**.

## Dashboard Components

### 1. Executive Overview

Provides a quick view of:

- Total customers
- Average customer age
- Gender distribution
- Demographic mix

### 2. Demographic Analysis

Explores customer composition through:

- Age-group distribution
- Gender proportions
- Age × gender analysis
- Average age comparisons

### 3. Geographic Analysis

Highlights:

- Customer distribution by city
- High-density customer markets
- Regional concentration patterns

### 4. Data Quality

Makes the transformation process transparent by showing how missing demographic values were handled.

## Excel Techniques Used

This project demonstrates practical Excel analytics skills:

- `IF()` / `ISBLANK()` for missing-value handling
- `VLOOKUP()` for gender-based imputation
- `AVERAGE()` and `ROUND()` for statistical calculations
- Nested `IF()` logic for age-group classification
- Excel Tables and structured references
- Pivot Tables
- Pivot Charts
- Conditional Formatting
- Dashboard layout and KPI cards

### Example transformation logic

```excel
=IF(ISBLANK([@Age_Old]),
   IFERROR(VLOOKUP([@Gender],gender_age_table,2,FALSE),
   ROUND(AVERAGE([@Age_Old]),0)),
   [@Age_Old])
```

## Business Insights

The dashboard is designed to help retail stakeholders answer questions such as:

- Which age groups form the largest customer base?
- Is the customer base balanced across genders?
- Which cities represent the strongest customer concentration?
- Which demographic segments could be prioritized for targeted campaigns?
- How much of the dataset required data-quality treatment?

### Example business actions

1. Prioritize marketing campaigns toward the largest customer segments.
2. Compare city-level customer concentration when planning regional campaigns.
3. Use age × gender combinations to create more focused customer segments.
4. Monitor data completeness before using demographic fields for decision-making.
5. Extend the analysis with purchase and revenue data before making value-based targeting decisions.

> **Important:** The current dataset is synthetic and demographic-focused. It should not be interpreted as evidence of actual customer behavior or revenue performance.

## Project Structure

```text
Excel-Sales-Performance-and-Revenue-Insights-Dashboard/
│
├── retail_synthetic_dataset.xlsx   # Source synthetic customer data
├── Retail_Dashboard.xlsx           # Main Excel dashboard
├── README.md                       # Project documentation
└── .gitignore                      # Local/temporary file exclusions
```

## How to Explore the Project

1. Download or clone the repository.
2. Open `Retail_Dashboard.xlsx` in Microsoft Excel.
3. Start from the dashboard/summary sheet.
4. Use the available filters and charts to explore customer segments.
5. Review the underlying customer and data-quality sheets to understand the transformations.

## Recommended Future Enhancements

To evolve this project from demographic reporting into a stronger business-analytics portfolio project:

- Add transaction-level sales and revenue data.
- Introduce **RFM analysis** (Recency, Frequency, Monetary).
- Calculate customer lifetime value.
- Add cohort and retention analysis.
- Build customer acquisition trends over time.
- Add Power Query for repeatable ETL.
- Rebuild the final model in Power BI with DAX measures and drill-through pages.
- Add automated refresh and a KPI-driven executive summary.

## Skills Demonstrated

**Excel | Data Cleaning | Data Validation | Pivot Tables | Dashboard Design | KPI Reporting | Customer Analytics | Data Visualization | Business Insights**

## Author

**Anamika Yadav**  
Data Analyst | Power BI | SQL | Python | Excel | Data Analytics

[GitHub](https://github.com/anamikavyada)
