# 📊 Multi-File ETL Automation — Excel VBA

> Automated pipeline that processes **70+ Excel distributor files** and compiles them into a single structured database in **under 4 minutes** — replacing a manual process that previously took several hours each month.

---

## 🧠 Business Context

Developed during my role as **Senior Commercial Performance & BI Analyst at Danone Sub-Saharan Africa**, where I managed a ~€120M revenue scope across 15+ markets.

Each month, 70+ distributors submitted individual VMI (Vendor-Managed Inventory) files tracking deliveries, sales and stock levels. These files shared an identical column structure but arrived unnormalized — requiring manual processing before any analysis or Power BI ingestion could take place.

This tool was built to **eliminate that bottleneck entirely**.

---

## ⚙️ What the Script Does

```
[70+ raw distributor .xlsx files]
         ↓
[STEP 1] Open master Template (formula normalization logic)
         ↓
[STEP 2] Reset output BDD — delete & recreate from scratch (no duplicates)
         ↓
[STEP 3] Loop through every .xlsx in the source folder:
          • Apply template formula range → normalize structure
          • Save normalized file
          • Append clean values to BDD (paste values only, no formulas)
          • Close file → move to next
         ↓
[STEP 4] Save final BDD → ready for Power BI or further analysis
```

### Key design decisions

| Decision | Reason |
|---|---|
| BDD deleted & recreated on every run | Guarantees zero duplicate data across monthly runs |
| `PasteSpecial xlPasteValues` only | Keeps the output file lightweight and tool-independent |
| Template-based normalization | One single update to the template propagates to all 70+ files |
| `Dir()` loop with self-exclusion | Output file is in the same folder — script skips it automatically |

---

## 📈 Results

| Metric | Before | After |
|---|---|---|
| Monthly processing time | Several hours (manual) | **< 4 minutes** |
| Risk of human error | High | Near zero |
| Data delivery date | ~22nd of the month | **2nd of the month** |
| Manual handling rate | ~90% | **< 10%** |

---

## 🗂️ Repository Structure

```
📁 ETL-MultiFile-VBA/
├── ETL_MultiFile_Compiler.bas       ← Main VBA script (importable into any .xlsm)
├── README.md                        ← This file
└── sample_data/                     ← (Optional) Anonymized sample files for testing
```

---

## 🚀 How to Use

### Prerequisites
- Microsoft Excel (2016 or later recommended)
- Macro execution enabled (`File → Options → Trust Center → Enable Macros`)

### Setup
1. Clone or download this repository
2. Open your Excel workbook and press `Alt + F11` to open the VBA editor
3. Go to `File → Import File` and select `ETL_MultiFile_Compiler.bas`
4. Update the three configuration variables at the top of the script:

```vba
cheminDossier     = "C:\YourPath\SourceFiles\"          ' Folder with all source files
nomFichierTemplate = "YOUR_TEMPLATE.xlsm"               ' Your master template
nomFichierBDD      = "YOUR_OUTPUT_BDD.xlsx"             ' Desired output filename
```

5. Update the worksheet name and cell ranges to match your file structure:

```vba
' Normalization range (template → source files)
classeurTemplate.Sheets("YourSheetName").Range("Q9:AM475")

' Compilation range (source files → BDD)
classeurCourant.Sheets("YourSheetName").Range("Q11:AH475")
```

6. Run the macro: `Alt + F8` → Select `TraiterEtCompiler` → `Run`

---

## 🔧 Customization

The script is designed to be easily adapted:

- **Different sheet names** → update `Sheets("Sales")` references
- **Different data ranges** → update `Range("Q9:AM475")` to match your structure
- **File type** → change `"*.xlsx"` to `"*.xlsm"` or `"*.csv"` as needed
- **Add column headers to BDD** → copy header row before the loop starts

---

## 👤 Author

**Serge NGUENA** — Senior Commercial Performance & BI Analyst  
[LinkedIn](https://linkedin.com/in/serge-nguena) • Montreal, QC • Permanent Resident

*Microsoft Certified: Data Analyst Associate (Power BI)*

---

## 📄 License

MIT License — free to use, adapt and share with attribution.
