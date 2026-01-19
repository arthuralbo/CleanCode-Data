# CleanCode Data Engine 🛡️

**A Mathematical Gatekeeper for Automated Data Engineering.**

CleanCode Data is a non-destructive data validation and engineering extension built to bridge the gap between "messy" raw datasets and model-ready analysis. It enforces mathematical integrity at the source, ensuring that datasets are structurally sound before they enter a Data Science pipeline.

---

## 🚀 The Core Problem
Most data science errors occur during the preprocessing phase due to:
* **Mixed Data Types:** Text strings hidden in numerical columns (e.g., "1,200$" vs 1200).
* **Date Fragmentation:** Inconsistent US/EU formats preventing time-series analysis.
* **Outlier Noise:** Extreme values skewing statistical distributions.
* **Missing Value Bias:** Naive imputation (like just filling zeros) corrupting the mean.

---

## ✨ Features

### 1. Strict Validation Phase (The Gatekeeper)
Unlike standard tools that "guess" data types, CleanCode Data runs a pre-scan validation. It identifies the exact row and column of any deal-breaking type mismatches, allowing users to fix structural errors before applying transformations.

### 2. Intelligent Data Modules
* **Date Standardization:** Smart detection of locales (US/EU/ISO) to convert messy strings into unified ISO-8601 or Unix timestamps.
* **Statistical Imputation:** Choose between Mean, Median, Mode, or Forward Fill logic based on the distribution of the variable.
* **Numeric Scaling:** Built-in Z-Score (Standardization), Min-Max (Normalization), and Winsorization (Outlier Capping) to prepare features for machine learning.
* **Categorical Encoding:** One-Hot Encoding (Dummy variables) and Ordinal Label Encoding with a custom drag-and-drop UI for rank-based variables.

### 3. Non-Destructive Workflow
The engine never overwrites original data. It generates new, labeled columns (e.g., `Price_standardized`) to maintain full data lineage and auditability.

---

## 🛠 Technical Stack

* **Logic:** JavaScript / Google Apps Script
* **Methodology:** Statistical outlier detection (3-Sigma Rule), Data Normalization, and Boolean Validation.

---

## ⚙️ How it Works
1. **Schema Mapping:** The tool scans headers and suggests data types (Numeric, Categorical, Date).
2. **Diagnostic Scan:** Validates every cell in the selection against the chosen schema.
3. **Module Activation:** Unlocks specific modules (Scaling, Cleanup, Encoding) only when relevant issues are detected.
4. **Transformation:** Executes batch operations via the Google Sheets API to generate a clean, analysis-ready sheet.

---

## 📈 Use Cases
* **Data Prep for ML:** Scaling features and handling missing values for Scikit-Learn.
* **Business Intelligence:** Cleaning manual-entry spreadsheets for PowerBI or Tableau.

---

## 📫 Contact
**Arthur Albo** *Math & Stats @ Concordia University* [LinkedIn](https://www.linkedin.com/in/arthuralbo) | [Portfolio Website](https://arthur-albo-portfolio.vercel.app)
