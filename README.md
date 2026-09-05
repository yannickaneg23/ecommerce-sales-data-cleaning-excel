# [Project Title]
> Ecommerce_Sales_Cleaned_Project

---

## ⚙️ Project Type Flags
> 


- [ ] Dashboard / Data Visualization
- [ ] Data Cleaning / Wrangling
- [ ] End-to-End (multiple of the above)


---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Objectives](#2-objectives)
3. [Project Scope & Tools](#3-project-scope--tools)
4. [Repository Structure](#4-repository-structure)
5. [Data Workflow](#5-data-workflow)
6. [Data Model & Schema](#6-data-model--schema)
7. [Analysis & Metrics](#8-analysis--metrics)
8. [Key Insights](#9-key-insights)
9. [Recommendations](#10-recommendations)
10. [Assumptions & Limitations](#11-assumptions--limitations)
11. [Future Enhancements](#12-future-enhancements)
12. [Author](#14-author)

---

## 1. Project Overview

<!--
  Context: An growing e-commerce platform exported its customer purchase history to track performance, but the system outputted a completely scrambled, unreadable spreadsheet.
  
  Problem Statement: Large 16-digit credit card codes were getting corrupted by Excel's numbering system, customer mailing addresses were jammed into single cells with messy line breaks, and random text units were mixed into pricing columns, making it impossible to calculate accurate revenue or group customer demographics.
  
  Approach: I intercepted the raw data load using Power Query to lock in the proper code formats, used hidden line-break commands (Ctrl + J) to unpack the addresses, and built dynamic Pivot Tables to safely categorize thousands of records.
  
  Outcome: I produced a spotless, audit-ready database connected to an interactive executive dashboard featuring high-level KPI cards and sorted charts that instantly reveal which customer occupations spend the most money on the platform.
---

## 2. Objectives

<!--
Primary Objective: Build a fully reproducible data cleansing pipeline and dynamic reporting dashboard that transforms corrupted, raw transaction exports into an audit-ready format.
Secondary Objective 1: Protect payment data integrity by isolating the ingestion schema to prevent Excel from permanently truncating 16-digit credit card fields.
Secondary Objective 2: Parse bundled, multi-line address text strings to isolate distinct geographical attributes for structural regional tracking.
Secondary Objective 3: Quantify and rank overall purchase totals by customer occupation to identify the top three highest-spending professional demographics on the platform.

--- 


## 3. Project Scope & Tools

### Scope

<!--
 Scope & Constraints
 In Scope: Processing and engineering of raw row-level customer profile and transactional data fields. This includes parsing messy address strings, extracting numeric lot values, standardizing mixed text casing, and removing exact row duplicates.
 
 Out of Scope: Active tracking of live browser cookie data and backend payment gateway logging were excluded. The browser details were captured as a static user-agent string text block, and deep payment logs sit in an isolated vendor system outside the scope of this file.
 
 Time Period: Historical customer profile and transaction export log (current file data snapshot).
 
 Granularity: Row-level data. Each individual row represents a single unique customer account transaction profile.
 
 Tools & Technologies
 
 Data Storage: Flat raw text files (CSV format)
 Data Processing: Microsoft Excel (Power Query Editor data ingestion pipeline configuration)
 Analysis: Microsoft Excel Advanced Formulas (=LEFT(), =TRIM(), =PROPER(), and Go To Special bulk-patching functionality)
 Visualization: Microsoft Excel Dynamic Pivot Tables, Sorted Horizontal Pivot Charts, and Custom KPI Shapes
 Version Control: GitHub Web Interface
 Documentation: Markdown text files
 ---

## 4. Repository Structure
<!--
```
[project-root]/
data/
raw/ — Contains the original, unmodified raw e-commerce customer transaction export (.csv).
processed/ — Houses the finalized, optimized master dataset and interactive workbook (.xlsx).
reports/ — Stores the polished, interview-ready portfolio case study documentation (.pdf).
visuals/ — Contains the high-resolution dashboard screenshot showcasing your dynamic chart and KPI metrics (.png).
README.md — The landing page of your project repository containing goals, approach, data dictionary, and business recommendations.
---

## 5. Data Workflow

<!--
 [Raw E-Commerce Customer Profile CSV Export]
                     ↓
[Power Query Schema Ingestion & Data Type Lock-In]
                     ↓
[Advanced Excel Cleaning Formulas & Bulk Omission Repairs]
                     ↓
[Dynamic Pivot Table Summarization & Segment Aggregation]
                     ↓
[Polished Visual Dashboard featuring Live KPI Cards & Sorted Bar Charts]

1. Source: The data came from an automated operational system export containing raw e-commerce customer transaction profiles. It was formatted as a flat, unseparated CSV text file containing thousands of individual row logs with hidden line breaks and unformatted data strings.
2. Ingestion: The file was brought into Microsoft Excel using the "From Text/CSV" data pipeline tool. During the import phase, the raw schema was routed through the Power Query Editor to manually switch the Credit Card column format from a number to flat text, permanently blocking Excel from truncating the 16-digit card digits.
3. Cleaning: I removed exact row duplicates to prevent data skewing. Next, I stripped text characters from numeric columns using the =LEFT() formula to make the data mathematical, standardized messy job titles by pairing =TRIM() and =PROPER() to fix spacing and casing errors, and deployed the Go To Special > Blanks tool to mass-fill missing fields with an "Unknown" label using Ctrl + Enter.
4. Transformation: I completely reshaped the dataset layout by using the Text-to-Columns wizard tool. First, I applied a custom keyboard line-break command (Ctrl + J) to unpack crammed, multi-line address blocks into four clean fields (Street Address, City, State, and Zip Code), and then used standard period-delimiters to break bundled system codes down into a structured, three-tier product category tree.
5. Analysis: I utilized visual, aggregation-based analytical methods strictly within Excel. I built dynamic Pivot Tables to automatically group the rows and calculate the sum of total purchases, and then implemented a descending sort hierarchy based on total spending to immediately isolate the top three customer demographics.
6. Output: The final results take the form of an optimized, master .xlsx workbook containing pure numeric tables, live sorting formulas, and a presentation-ready dashboard tab with dark KPI metrics cards and horizontal bar charts
---

## 6. Data Model & Schema

<!--
 ### Dataset / Table: `Cleaned_ECommerce_Transactions`

| Field Name | Data Type | Description | Example Value |
|------------|-----------|-------------|---------------|
| `Street_Address` | string | Unpacked physical street number and avenue delivery data. | `"16629 Pace Camp Apt. 448"` |
| `City` | string | Extracted geographical municipality name for regional tracking. | `"Alexisborough"` |
| `State` | string | Extracted two-letter regional state code. | `"NE"` |
| `Zip_Code` | string | Isolated baseline or long-tail postal shipping code. | `"77130-7478"` |
| `Lot` | int | Cleaned numeric item dimension or batch metric code. | `46` |
| `AM or PM` | string | Time-of-day purchase window indicator. | `"PM"` |
| `Browser Info` | string | Raw user-agent device string capturing the client's software setup. | `"Opera/9.56 (X11; Linux x86_64)"` |
| `Company` | string | Standardized title-case name of the purchasing business account. | `"Martinez-Herman"` |
| `Credit Card` | string | Intercepted 16-digit payment card identification string. | `"6011929061123406"` |
| `CC Exp Date` | date | The credit card expiration date formatted as month/year. | `"02/20"` |
| `CC Security Code` | string | The 3-to-4 digit card verification security pin value. | `"900"` |
| `CC Provider` | string | The specific financial payment gateway network issuer. | `"JCB 16 digit"` |
| `Email` | string | Registered primary text electronic mailing contact address. | `"pdunlap@yahoo.com"` |
| `Job` | string | Polished title-case classification of the buyer's profession. | `"Scientist, Product/Process Development"` |
| `IP Address` | string | The network device Internet Protocol location tracking address tag. | `"149.146.147.205"` |
| `Language` | string | Two-letter ISO primary web layout communication language code. | `"el"` |
| `Purchase Price` | float | Decentered transaction metric currency cost. | `98.14` |

> **Row count (approx.):** Thousands of entries (depending on your full platform export).
> **Date range:** Historical transaction capture window.
> **Key structural notes:** This is a flat, single-table analytical database matrix. The columns `Street_Address`, `City`, `State`, and `Zip_Code` were systematically parsed from an originally bundled raw multi-line system cell block using Excel delimiters.

---


---

## 7. Analysis & Metrics

<!--
  ### Analytical Approach

For this project, my approach was focused on data pipeline standardization and exploratory group analysis. E-commerce data platforms often generate erratic data inputs that distort business metrics. I first constructed a predictable data cleaning framework to restore structural integrity to the rows. Once the records were sanitized, I conducted group analysis to uncover hidden purchasing behaviors across distinct customer demographics.

---

### Key Metrics Defined

*   **Metric:** Total Gross Revenue
    *   **Plain-Language Definition:** The cumulative sum of all financial transactions processed on the platform during the recorded export window.
    *   **Why It Matters:** It determines the absolute financial scale of the transactional dataset and serves as the baseline target for demographic performance ranking.

*   **Metric:** Transaction Volume
    *   **Plain-Language Definition:** The total absolute count of individual purchase records logged across the database.
    *   **Why It Matters:** It tracks buyer frequency and baseline order activity, helping to separate high-value single transactions from frequent, lower-value purchasing behavior.
    
*   **Metric:** Demographic Purchase Power
    *   **Plain-Language Definition:** The total gross spending aggregation grouped specifically by unique customer occupational titles.
    *   **Why It Matters:** It identifies which professional industries generate the highest financial value, allowing marketing teams to optimize their customer acquisition targeting budgets.

---

### Methods Used

*   **Data Profiling and Schema Enforcement:** I utilized Power Query to audit raw source strings and enforce explicit text formatting configurations on 16-digit data fields to permanently stop precision clipping errors.
*   **Categorical Text Disaggregation:** I deployed string separation tools to isolate bundled physical location entries and break down single product categories into structured, multi-tier analysis paths.
*   **Missing Value Mass-Imputation:** I isolated incomplete record segments simultaneously using index-grouping mechanisms to run a uniform data placeholder override, protecting final summary metrics from skewing.
*   **Demographic Segmentation and Grouping:** I constructed multi-variable Pivot Tables to isolate and consolidate thousands of fragmented line-item rows into explicit customer occupation categories.
*   **Rank-Order Trend Analysis:** I applied descending sorting hierarchies to segment groups from highest grossing to lowest grossing, providing stakeholder teams with an immediate, scannable look at top-tier spending cohorts.


---

## 8. Key Insights

<!--
 ### Key Findings & Insights

**Insight 1: Primary Platform Revenue Drivers**
The data shows that the highest cumulative purchase values are concentrated within three specific professional occupations—Analytical Chemists, Drilling Engineers, and Town Planners. This indicates that platform revenue is heavily dependent on specific B2B or high-income professional brackets rather than a general, distributed consumer baseline. This suggests that future marketing campaigns should pivot from broad consumer targeting to industry-specific professional ad networks to maximize customer acquisition efficiency.

**Insight 2: Operational Risks From System Default Overrides**
During data profiling, I discovered that approximately 100% of large 16-digit data fields were corrupted due to standard spreadsheet software loading configurations. By defaulting to numeric scales, the system permanently rounded critical payment codes, causing total tracking confusion for transaction validation. This indicates a significant risk of silent data loss across the operations department, highlighting an urgent need to enforce explicit text-formatting rules at the ingestion phase.

**Insight 3: Visual Readability Gaps in Standard Reporting**
An evaluation of standard reporting columns showed that horizontal text space limits heavily clipped long-tail customer descriptions and job titles. When summaries were rendered as vertical column charts, text titles overlapped, rendering the report completely unreadable for quick executive assessments. Pivoting the reporting environment to sorted horizontal clustered bar layouts instantly resolved visual clutter, proving that data presentation choice directly impacts decision-making speed.

**Insight 4: Administrative Typos Invalidate Demographic Segmentation**
A deep scan of customer profiles revealed a high frequency of mixed capitalization, trailing double spaces, and random abbreviation variants across self-reported entries. These data discrepancies caused Excel to treat identical professional segments as completely separate customer rows, creating a fractured view of actual customer performance. This means that top-level company analytics have likely been missing true macro-trends due to a lack of automated validation constraints at the initial customer registration point.

---

## 9. Recommendations

<!--
 ### Actionable Business Recommendations

| Priority | Recommendation | Based On | Suggested Owner |
| :--- | :--- | :--- | :--- |
| **High** | Adjust customer acquisition budgets to target digital marketing ad spend specifically toward the top three highest-spending professional job titles identified on the bar chart. | Insight 1 - High concentration of revenue in specific professional segments | Paid Acquisition / Marketing Team |
| **Medium** | Reconfigure the core e-commerce database export settings to enforce strict text-string parameters on all 16-digit data fields. This stops data truncation and numeric rounding bugs before files are ever downloaded. | Insight 2 - Operational risk of numeric data corruption and lost card digits | IT / Systems Engineering Team |
| **Low** | Deploy a standardized drop-down menu on the client profile signup page to completely replace open-text entry boxes for addresses and company names. This neutralizes manual typos and structural text bugs at the source. | Insight 4 - Demographic fragmentation caused by irregular user text spacing and casing | UI/UX Development Team |


---

## 10. Assumptions & Limitations

<!--
 ### Assumptions

*   **Transactional Completeness:** I treated the downloaded platform transaction export as a complete and final record of all purchase events during that time period, without verifying the row totals against the company's internal merchant gateway ledger logs.
*   **Accuracy of Human Casing Inputs:** I assumed that rows with minor casing differences or spacing anomalies (e.g., `"Drilling engineer"` vs `"drilling engineer"`) represented the exact same occupational demographic, mapping them into a single grouping via string formulas.
*   **Stable Product Pricing Structures:** The `Purchase Price` was accepted as the final, true financial value paid by the consumer at the exact moment of transaction, without accounting for potential unlisted flash coupons or individual loyalty discount code parameters.

### Limitations

*   **Absence of Customer Tenure Context:** The dataset lacks a registration or customer age timestamp. Because of this, the analysis cannot differentiate whether high spending is driven by long-term loyal clients or newly acquired one-time premium buyers.
*   **No Return or Refund Attributions:** While we cleaned and sorted transactional values, the file contains no status tracking or return flags. If a high-spending job demographic also has an incredibly high product return rate, our total gross revenue metrics would overstate their actual net financial value to the company.
*   **Lack of Historical Time Depth:** This project evaluates a single static file snapshot. A more rigorous, production-grade version of this analysis would track these customer demographics dynamically across several sequential quarters to account for seasonal purchasing spikes (like holiday shopping).
*   **Self-Reported Job Profile Bias:** The `Job` category relies entirely on unverified, self-reported user profiles created during signup. This introduction of human text input data creates a natural bias, as users can list aspirational, outdated, or completely fictional professional titles.

---

## 11. Future Enhancements

<!--
 ### Next Steps & Future Enhancements

- [ ] **Automate the Data Processing Pipeline:** Build an Excel macro macro script or transition the cleaning workflow into a reusable Python/Pandas script to automatically ingest and sanitize new transaction exports instantly.
- [ ] **Integrate Refund and Status Tracking:** Blend historical product return logs into this dataset to cross-reference customer professions against actual net revenue, neutralizing the current limitation regarding missing return variables.
- [ ] **Establish Upstream UI Input Validation:** Deploy an absolute input constraint (such as standardized drop-down select buttons) on the user profile signup screen to fix customer typos and mixed spacing attributes at the point of data entry.
- [ ] **Conduct Time-Series Spending Analysis:** Expand the historical collection window across multiple consecutive quarters to track how specific high-spending occupational demographics behave during seasonal market fluctuations or holiday spikes.

---

---

## 12. Author

**[ANEG YANNICK**
[DATA ANALYST]

- 🔗 [[LinkedIn URL](https://www.linkedin.com/in/aneg-yannick-19692a432/)]
- 💼 [](https://github.com/yannickaneg23/ecommerce-sales-data-cleaning-excel)L]
- 📧 [yannickaneg23@gmail.com]

---

*Last updated: [August 2026]*
*If this template helped you, consider starring the repository.*
