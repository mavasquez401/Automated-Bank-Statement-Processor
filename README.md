# 🏦 Statement Processor - Blue Prism RPA Automation

This Blue Prism RPA project automates the validation and routing of bank statement transactions stored in Excel spreadsheets. Designed without any external VBOs or skills from the Digital Exchange, the solution leverages native Excel VBO actions and core Blue Prism functionalities.

---

## 📌 Features

- ✅ **Reads bank statements** from an Excel workbook  
- 🔍 **Validates transactions** based on business logic (e.g., amount > 0)  
- 📤 **Routes valid transactions** to a `ProcessedStatements.xlsx` file  
- ⚠️ **Logs invalid transactions** with contextual error messages  
- 🧠 **Fully automated decision flow** using collections, loops, and expressions  
- 📂 **No external VBOs or custom assets** required (DX not used)

---

## 🛠️ Tools & Technologies

| Tool             | Description                                 |
|------------------|---------------------------------------------|
| **Blue Prism**   | Version 7.4.0 (no DX dependencies)          |
| **Excel VBO**    | Used for reading/writing/managing workbooks |
| **Collections**  | Used for row-level validation and processing |
| **Decisions**    | Conditional routing (e.g., amount validation) |
| **Loops**        | Iterate through transactions                |
| **Exception Handling** | Custom paths for invalid entries     |

---

## 📁 Input & Output

### Input File (`BankStatements.xlsx`)

| Customer Name | Date       | Amount | Type      |
|---------------|------------|--------|-----------|
| John Smith    | 5/1/2024   | 150    | Deposit   |
| Jane Doe      | 5/2/2024   | -25    | Withdrawal |
| ...           | ...        | ...    | ...       |

### Output Files

- `ProcessedStatements.xlsx`: Valid entries (amount > 0)
- `ErrorLog.xlsx`: Invalid entries with appended "Error Message" field

---

## 🔄 Process Flow Overview

```plaintext
Start
 └── Create Excel Instance
 └── Open Source Workbook
 └── Read Bank Data (from Excel to Collection)
 └── Copy Row Structure to CurrentRow
 └── Loop Through SourceData
     ├── Copy Current Row
     ├── Decision: Is Amount > 0?
     │   ├── Yes → Append to ProcessedStatements.xlsx
     │   └── No  → Add Error Message + Append to ErrorLog.xlsx
 └── Close All Workbooks
 └── Release Excel Instance
End
