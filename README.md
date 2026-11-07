# product_sales
data cleaning assignment2
📊 Data Cleaning & Transformation — Product Inventory Dataset

A comprehensive data cleaning, transformation, and standardization project performed using **Microsoft Excel** and **Power Query Editor**. This repository documents the end-to-end workflow applied to a 34-row product inventory dataset, covering missing value imputation, text standardization, column splitting/merging, deduplication, number formatting, and conditional formatting.

---

## 📋 Project Overview

Raw datasets often contain inconsistencies, missing entries, duplicates, and unformatted fields. This project demonstrates a structured approach to transforming raw product inventory data into an analysis-ready format through the following stages:

1. **Missing Value Treatment** (Price & Category Imputation)
2. **Text Standardization** (Find & Replace Operations)
3. **Splitting & Merging Columns** (Excel Text Formulas)
4. **Number & Date Formatting** (Currency & `DD-MM-YYYY` Standardization)
5. **Conditional Formatting** (Color Scales & Custom Rules)
6. **Duplicate Removal** (Power Query M Language)

---

## 🛠️ Tools & Technologies Used

- **Microsoft Excel** (Formulas, Find & Replace, Conditional Formatting, Number Formatting)
- **Power Query Editor** (`Table.Distinct` M-code transformations)

---

## 🔄 Data Cleaning & Transformation Workflow

### 1. Handling Missing Values
An audit of the 34-row dataset identified missing values in two key columns:

* **Missing Prices (3 records):** *Sony Headphones, Coleman Camping Tent, Ray-Ban Sunglasses*
  * **Formula Applied:**
    ```excel
    =ROUND(IF(ISBLANK([@price]), AVERAGE([price]), [@price]), 0)
    ```
  * **Evaluation & Recommendations:** While global mean imputation (`$309`) fills blanks quickly, it can distort item-specific pricing (e.g., pricing a tent equal to luxury sunglasses). **Category-Based Mean** or **Historical/Match Lookup** (e.g., matching Bose headphones at `$250`) are recommended for higher accuracy.

* **Missing Categories (4 records):** *North Face Backpack, Adidas Sneakers, Nespresso Coffee Maker, Xiaomi Fitness Tracker*
  * **Formula Applied:**
    ```excel
    =IF(ISBLANK([@Category]), "Electronics", [@Category])
    ```
  * **Evaluation & Recommendations:** Static fallback (`"Electronics"`) misclassifies items like backpacks and coffee makers. Recommended alternatives include **Keyword Mapping** (`"Coffee Maker"` → `Kitchen`), **Brand Lookup** (`North Face` → `Outdoor`), and **Text Clean-up**.

---

### 2. Correcting Inconsistent Data (Find & Replace)
Three sequential Find & Replace operations were executed to standardize text formatting and fix typos:

| Step | Target Column | Find What | Replace With | Purpose / Result |
| :--- | :--- | :--- | :--- | :--- |
| 1 | `Product Name` | `headphones` | `Headphones` | Standardized lowercase entries to Title Case |
| 2 | `Product Name` | `smartphone` | `Smartphone` | Corrected lowercase entries to capitalized format |
| 3 | `Category` | `Electronicscscs` / `Electroni` | `Electronics` | Fixed 6 corrupted/truncated category spelling errors |

---

### 3. Splitting and Merging Data
Extracted structured attributes from `Product ID` and combined product/brand fields:

| Target Column | Excel Formula | Operation Description |
| :--- | :--- | :--- |
| **Manufacturing Date** | `=LEFT(A2,2)` | Extracted the 2-digit manufacturing day from `Product ID` |
| **Country Code** | `=RIGHT(A2,2)` | Extracted the 2-letter country code from `Product ID` |
| **Product Brand** | `=CONCATENATE(D2," ", F2)` | Merged `Product Name` and `Brand Name` with a space separator |

---

### 4. Number & Date Formatting
* **Price Column:** Formatted as Currency (`$#,##0.00`).
* **Manufacturing Date Column:** Converted and formatted to `DD-MM-YYYY` using:
  ```excel
  =DATEVALUE(LEFT(A2,6)&"-2026")
5. Conditional Formatting
Applied visual cues to highlight pricing tiers and specific product categories:

Rule	Target Column	Formatting Type	Condition Details
1	Price (IMPUTED)	3-Color Scale	Green ($30–$80), Peach/Orange ($90–$500), Red ($700–$1000)
2	Category (IMPUTED)	Custom Highlight Rule	Highlight cells equal to "Electronics" with a light red fill
6. Duplicate Removal (Power Query Editor)
Seven sequential Table.Distinct steps were applied in Power Query (table_products) to audit and remove duplicates across columns, reducing the dataset from 34 rows to 3 distinct records:

= Table.Distinct(#"Changed Type", {"Product ID"})
= Table.Distinct(#"Removed Duplicates", {"Product Name"})
= Table.Distinct(#"Removed Duplicates1", {"Brand Name"})
= Table.Distinct(#"Removed Duplicates2", {"Price ($)"})
= Table.Distinct(#"Removed Duplicates3", {"Quantity"})
= Table.Distinct(#"Removed Duplicates4", {"Category"})
= Table.Distinct(#"Removed Duplicates5", {"Category (IMPUTED)"})
📁 Repository Structure
├── Assignment_2_-_Full_Report___Dataset.pdf                    # Complete dataset & missing value strategy
├── Assignment_2_-_Data_Cleaning_and_Transformation.pdf         # Dataset record summary & audit log
├── Assignment_2_-_Complete_Data_Cleaning_Report.pdf            # Post-transformation dataset report
├── Assignment_2_-_Data_Cleaning___Inconsistent_Data_Report.pdf # Find & Replace audit log
├── Power_Query_Editor_-_Duplicate_Removal_Report.pdf           # Power Query M-code deduplication audit
├── Splitting_and_Merging_Data_Report.pdf                       # Column split/merge formulas & output
├── Conditional_Formatting_Report.pdf                           # Color scale & custom highlight rules
├── Product_Inventory_Data.pdf                                  # Final currency & date formatted dataset
└── README.md                                                   # Project documentation
🚀 Key Takeaways
Data Integrity: Imputation strategies should align with domain context (e.g., category-level averages over global averages).
Standardization: Consistent casing and typo resolution prevent downstream grouping errors in pivot tables and queries.
Automation: Combining Excel formulas with Power Query M-scripts creates a repeatable, auditable data cleaning pipeline.
🎯 Conclusion
This project demonstrates how structured data cleaning and transformation workflows turn raw, inconsistent inventory records into reliable, analysis-ready datasets. By combining Excel formulas, Find & Replace standardization, conditional formatting, and Power Query M-script deduplication, the pipeline resolves missing values, eliminates formatting discrepancies, and ensures data integrity across all 34 product records. Moving forward, replacing static fallbacks and global mean imputations with context-aware lookups (such as category-based pricing and brand-to-category mapping) will further improve accuracy for downstream reporting and business intelligence.
