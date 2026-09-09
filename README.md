# 📊 E-Commerce Customer Profile & Data Cleansing Project

A complete, end-to-end data preparation, optimization, and reporting dashboard architecture completed strictly within **Microsoft Excel**.

---

## 📌 Project Overview

*   **Context:** A rapidly growing e-commerce platform exported its customer purchase history to track performance, but the automated system outputted a completely scrambled, unreadable dataset.
*   **Problem Statement:** Large 16-digit credit card codes were getting corrupted by Excel's default numbering system, customer mailing addresses were jammed into single cells with hidden system line breaks (`\n`), and random text units were mixed into pricing columns, making it impossible to calculate accurate revenue metrics or group consumer segments.
*   **Approach:** I intercepted the raw data load using Power Query to lock in the proper code formats, handled missing records globally, deployed custom keyboard macro commands (`Ctrl + J`) to unpack messy addresses, and structured multi-variable summaries using dynamic data manipulation formulas.
*   **Outcome:** I transformed a corrupted raw text file into a spotless, audit-ready database infrastructure connected to an interactive executive dashboard featuring high-level KPI cards and horizontal bar charts that instantly identify top-spending consumer demographics.

---

## 🎯 Objectives

*   **Primary Objective:** Build a fully reproducible data cleansing pipeline and dynamic reporting dashboard that transforms corrupted, raw transaction exports into a structured, executive-ready database format.
*   **Secondary Objective 1:** Protect payment data integrity by isolating the ingestion schema to prevent Excel from permanently truncating 16-digit credit card code fields.
*   **Secondary Objective 2:** Parse bundled, multi-line address text strings to isolate distinct geographical attributes for localized regional shipping tracking.
*   **Secondary Objective 3:** Quantify and rank overall purchase totals by customer occupation to identify the top three highest-spending professional demographics on the platform.

---

## 📐 Project Scope & Tools

### Scope & Constraints
*   **In Scope:** Processing and engineering of raw row-level customer profile and transactional data fields. This includes parsing messy address strings, extracting numeric lot values, standardizing mixed text casing, and removing exact row duplicates.
*   **Out of Scope:** Active tracking of live browser cookie data and backend payment gateway logging were excluded. Browser details were captured as a static user-agent string text block, and deep payment logs sit in an isolated vendor system outside the scope of this project.
*   **Time Period:** Historical customer profile and transaction export log (current file data snapshot).
*   **Granularity:** Row-level data. Each individual row represents a single unique customer account transaction profile.

### Tools & Technologies
*   **Data Storage:** Flat raw text files (CSV format)
*   **Data Processing:** Microsoft Excel (Power Query Editor data ingestion pipeline configuration)
*   **Analysis:** Microsoft Excel Advanced Formulas (`=LEFT()`, `=TRIM()`, `=PROPER()`, and `Go To Special` bulk-patching functionality)
*   **Visualization:** Microsoft Excel Dynamic Pivot Tables, Sorted Horizontal Pivot Charts, and Custom KPI Shapes
*   **Version Control:** GitHub Web Interface
*   **Documentation:** Markdown text files

---

## 🗂️ Repository Structure

*   **`[project-root]/`**
    *   **`data/`**
        *   **`raw/`** — Contains the original, unmodified raw e-commerce customer transaction export (`.csv`).
        *   **`processed/`** — Houses the finalized, optimized master dataset and interactive workbook (`.xlsx`).
    *   **`reports/`** — Stores the polished, interview-ready portfolio case study documentation (`.pdf`).
    *   **`visuals/`** — Contains the high-resolution dashboard screenshot showcasing your dynamic chart and KPI metrics (`.png`).
    *   **`README.md`** — The landing page of your project repository containing goals, approach, data dictionary, and business recommendations.

---

## 🔄 Data Workflow

*   **1. Source:** The data came from an automated operational system export containing raw e-commerce customer transaction profiles. It was formatted as a flat, unseparated CSV text file containing thousands of individual row logs with hidden line breaks and unformatted data strings.
*   **2. Ingestion:** The file was brought into Microsoft Excel using the "From Text/CSV" data pipeline tool. During the import phase, the raw schema was routed through the Power Query Editor to manually switch the `Credit Card` column format from a number to flat text, permanently blocking Excel from truncating the 16-digit card digits.
*   **3. Cleaning:** I removed exact row duplicates to prevent data skewing. Next, I stripped text characters from numeric columns using the `=LEFT()` formula to make the data mathematical, standardized messy job titles by pairing `=TRIM()` and `=PROPER()` to fix spacing and casing errors, and deployed the **Go To Special > Blanks** tool to mass-fill missing fields with an `"Unknown"` label using `Ctrl + Enter`.
*   **4. Transformation:** I completely reshaped the dataset layout by using the `Text-to-Columns` wizard tool. First, I applied a custom keyboard line-break command (**`Ctrl + J`**) to unpack crammed, multi-line address blocks into four clean fields (`Street Address`, `City`, `State`, and `Zip Code`), and then used standard period-delimiters to break bundled system codes down into a structured, three-tier product category tree.
*   **5. Analysis:** I utilized visual, aggregation-based analytical methods strictly within Excel. I built dynamic Pivot Tables to automatically group the rows and calculate the sum of total purchases, and then implemented a descending sort hierarchy based on total spending to immediately isolate the top three customer demographics.
*   **6. Output:** The final results take the form of an optimized, master `.xlsx` workbook containing pure numeric tables, live sorting formulas, and a presentation-ready dashboard tab with dark KPI metrics cards and horizontal bar charts. It also includes an interview-ready PDF case study documentation file for your portfolio.

---

## 📖 Data Model & Schema

### Dataset: `Cleaned_ECommerce_Transactions`

*   **Street_Address** (string): Unpacked physical street number and avenue delivery data. 
    *   *Example Value:* `"16629 Pace Camp Apt. 448"`
*   **City** (string): Extracted geographical municipality name for regional tracking. 
    *   *Example Value:* `"Alexisborough"`
*   **State** (string): Extracted two-letter regional state code. 
    *   *Example Value:* `"NE"`
*   **Zip_Code** (string): Isolated baseline or long-tail postal shipping code. 
    *   *Example Value:* `"77130-7478"`
*   **Lot** (int): Cleaned numeric item dimension or batch metric code. 
    *   *Example Value:* `46`
*   **AM or PM** (string): Time-of-day purchase window indicator. 
    *   *Example Value:* `"PM"`
*   **Browser Info** (string): Raw user-agent device string capturing the client's software setup. 
    *   *Example Value:* `"Opera/9.56 (X11; Linux x86_64)"`
*   **Company** (string): Standardized title-case name of the purchasing business account. 
    *   *Example Value:* `"Martinez-Herman"`
*   **Credit Card** (string): Intercepted 16-digit payment card identification string. 
    *   *Example Value:* `"6011929061123406"`
*   **CC Exp Date** (date): The credit card expiration date formatted as month/year. 
    *   *Example Value:* `"02/20"`
*   **CC Security Code** (string): The 3-to-4 digit card verification security pin value. 
    *   *Example Value:* `"900"`
*   **CC Provider** (string): The specific financial payment gateway network issuer. 
    *   *Example Value:* `"JCB 16 digit"`
*   **Email** (string): Registered primary text electronic mailing contact address. 
    *   *Example Value:* `"pdunlap@yahoo.com"`
*   **Job** (string): Polished title-case classification of the buyer's profession. 
    *   *Example Value:* `"Scientist, Product/Process Development"`
*   **IP Address** (string): The network device Internet Protocol location tracking address tag. 
    *   *Example Value:* `"149.146.147.205"`
*   **Language** (string): Two-letter ISO primary web layout communication language code. 
    *   *Example Value:* `"el"`
*   **Purchase Price** (float): Decentered transaction metric currency cost. 
    *   *Example Value:* `98.14`

> **Row count (approx.):** Thousands of entries (depending on your full platform export).
> **Date range:** Historical transaction capture window.
> **Key structural notes:** This is a flat, single-table analytical database matrix. The columns `Street_Address`, `City`, `State`, and `Zip_Code` were systematically parsed from an originally bundled raw multi-line system cell block using Excel delimiters.

---

## 📈 Analysis & Metrics

*   **Total Gross Revenue**
    *   *Definition:* This represents the cumulative sum of all financial transactions processed on the platform during the recorded export window. It determines the absolute financial scale of the transactional dataset and serves as the baseline target for demographic performance ranking.
*   **Transaction Volume**
    *   *Definition:* This tracks the total absolute count of individual purchase records logged across the database. It measures buyer frequency and baseline order activity, helping to separate high-value single transactions from frequent, lower-value purchasing behavior.
*   **Demographic Purchase Power**
    *   *Definition:* This measures the total gross spending aggregation grouped specifically by unique customer occupational titles. It identifies which professional industries generate the highest financial value, allowing marketing teams to optimize their customer acquisition targeting budgets.

---

---

## 🔍 Key Insights

**Insight 1: Primary Platform Revenue Drivers**
The data shows that the highest cumulative purchase values are concentrated within three specific professional occupations—Analytical Chemists, Drilling Engineers, and Town Planners. This indicates that platform revenue is heavily dependent on specific B2B or high-income professional brackets rather than a general, distributed consumer baseline. This suggests that future marketing campaigns should pivot from broad consumer targeting to industry-specific professional ad networks to maximize customer acquisition efficiency.

**Insight 2: Operational Risks From System Default Overrides**
During data profiling, I discovered that approximately 100% of large 16-digit data fields were corrupted due to standard spreadsheet software loading configurations. By defaulting to numeric scales, the system permanently rounded critical payment codes, causing total tracking confusion for transaction validation. This indicates a significant risk of silent data loss across the operations department, highlighting an urgent need to enforce explicit text-formatting rules at the ingestion phase.

**Insight 3: Visual Readability Gaps in Standard Reporting**
An evaluation of standard reporting columns showed that horizontal text space limits heavily clipped long-tail customer descriptions and job titles. When summaries were rendered as vertical column charts, text titles overlapped, rendering the report completely unreadable for quick executive assessments. Pivoting the reporting environment to sorted horizontal clustered bar layouts instantly resolved visual clutter, proving that data presentation choice directly impacts decision-making speed.

**Insight 4: Administrative Typos Invalidate Demographic Segmentation**
A deep scan of customer profiles revealed a high frequency of mixed capitalization, trailing double spaces, and random abbreviation variants across self-reported entries. These data discrepancies caused Excel to treat identical professional segments as completely separate customer rows, creating a fractured view of actual customer performance. This means that top-level company analytics have likely been missing true macro-trends due to a lack of automated validation constraints at the initial customer registration point.

---

## 💡 Recommendations

*   **High Priority Recommendation:** Adjust digital ad spend and targeting parameters to prioritize marketing campaigns toward the **top three highest-spending job categories** discovered in the sorted dashboard bar chart.
    *   **Based On:** Demographic Purchase Power Insights (Top Professional Spenders)
    *   **Suggested Owner:** Paid Acquisition / Marketing Team
<br>
*   **Medium Priority Recommendation:** Reconfigure the **upstream data export parameters** inside the primary e-commerce system platform to automatically output raw ID data fields as flat Text properties rather than scientific numbers. This stops data truncation errors before reports are ever downloaded by the team.
    *   **Based On:** Power Query Schema Ingestion Phase (Data Loss Discovery)
    *   **Suggested Owner:** IT / Systems Engineering Team
<br>
*   **Low Priority Recommendation:** Replace open-ended user text input boxes on the customer checkout interface with **standardized drop-down selection options** for region fields and industry titles. This stops structural typos at the source, preventing messy text rows in future database exports.
    *   **Based On:** Address and Job Title Text Wrangling Phase (Mismatched Casing and Bundled Strings)
    *   **Suggested Owner:** UI / UX Development Team

---

## ⚠️ Assumptions & Limitations

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

## 🚀 Future Enhancements

- [ ] **Automate the Data Processing Pipeline:** Build an Excel macro script or transition the cleaning workflow into a reusable Python/Pandas script to automatically ingest and sanitize new transaction exports instantly.
- [ ] **Integrate Refund and Status Tracking:** Blend historical product return logs into this dataset to cross-reference customer professions against actual net revenue, neutralizing the current limitation regarding missing return variables.
- [ ] **Establish Upstream UI Input Validation:** Deploy an absolute input constraint (such as standardized drop-down select buttons) on the user profile signup screen to fix customer typos and mixed spacing attributes at the point of data entry.
- [ ] **Conduct Time-Series Spending Analysis:** Expand the collection window across multiple consecutive quarters to track how specific high-spending occupational demographics behave during seasonal market fluctuations or holiday spikes



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
