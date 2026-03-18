# Excel Skills Reference

> Quick guide for users of the repository to locate specific Excel techniques
> used across the files in this repository.

---

## Skills Index

### 🔗 Nested Functions
| File | Sheet | Cell(s) | Description |
|------|-------|---------|-------------|
| Excel_Dashboard| data_Exp_Imp_per_country | A3 | UNIQUE function nested with VSTACK -> to retrieve the unique countries in the 2 different columns ({exporter_country, importer_country}) from the Packages_Database table |
| Excel_Dashboard | data_Exp_Imp_per_country | E3 | LET function with TAKE, SORT, VSTACK and UNIQUE -> to take the top 25 countries with more export total value and the top 25 countries with more import total value. Then Combine the 2 lists vertically, remove duplicates and sort the countries in ascending order by total value of export.  |

---

### 🔍 XLOOKUP
| File | Sheet | Cell(s) | Description |
|------|-------|---------|-------------|
| Excel_Dashboard | data_brokers | C5:C44 | Lookup broker name by ID from another workbook: brokers_database.xlsx |
| Excel_Dashboard | data_brokers | D5:D44 | Lookup broker department by ID from another workbook: brokers_database.xlsx |

---

### 👨‍💻 Power Query
| File | Sheet/Query Name | Description |
|------|-----------------|-------------|
| Excel_Dashboard | shipment_database | Imported data from Packages_database.csv  |
| Excel_Dashboard | Shipments_database | Changed the columns types (using local types when necessary) and sort rows |
| Excel_Dashboard | Shipments_database | Added column with M code to have: "package_priority" - priority correlated to the package value |
| Excel_Dashboard | Shipments_database | Added column with M code: "hold_import_responsibility" - to understand who is responsible for the package not being delivered |

---
### 🔍 PivotTables:

| File | Sheet | Cell(s) | Description |
|------|-------|---------|-------------|
| Excel_Dashboard | data_pivot_tables | A2:K10 | Pivot table using data from shipment_database sheet to report: in a monthly data frame, the total and % of shipments by import_query [{Delayed, Delivered, Missing Docs, Missing Info, Processing, Waiting Payment}] |
| Excel_Dashboard | data_pivot_tables | A14:D22 | Pivot table using data from shipment_database sheet to report: the responsibility of the packages not delivered yet by the value category [{Low value, Medium value, High value, Premium value}] in %. The responsibility is from Company if the package is [{Delayed, Processing}] or from Customer if the package is [{Missing Docs, Missing Info, Waiting Payment}]

---
## Other Skills Used
- Conditional Formatting
- Dynamic Charts (clustered bar + average line)
- Simple functions (IFS, SUMIF)



---

## About This Project
This Repo is about a fiction logistics company that import and export packages around the globe that I build to train and show my skills with Excel.