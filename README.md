# Excel Dashboard Practice

A personal project to practice and improve Excel skills including Functions, Power Query, and VBA Macros, using real-world data from multiple sources.

---

## Goals
- Practice and document Excel functions
- Build Power Query pipelines for data transformation
- Develop VBA macros for automation
- Create a dynamic dashboard from multiple data sources

---

## Project Structure
```
Excel_Dashboard/
├── data/               # Working data files
├── dashboard           # Processed data into dashboard
├── src/                # VBA modules exported as .bas files
├── docs/               # Notes and documentation
└── README.md
```

---

## Setup

1. Clone the repo
```bash
git clone <repo-url>
cd Excel_Dashboard
```

---

## Skills Practiced

| Area        | Topics                                        |
|-------------|-----------------------------------------------|
| Functions   | XLOOKUP, INDEX/MATCH, SUMIFS, dynamic arrays  |
| Power Query | Data import, transformations, merging sources |
| VBA         | Macros, automation, UserForms                 |

---

## Data Sources
- `data/Packages_database.csv` — Shipments database
- *(more sources to be added)*

---

## Notes
- Never edit files in `raw/` directly
- Export VBA modules as `.bas` files into `src/` before committing