Project Management GUI (PyQt6)

A Python-based GUI application for managing products and projects in NPI (New Product Introduction) workflows. Supports storing project details, MES entries, build/machine/assembly matrices, and checklists in a SQLite database, with Excel fallback.

Features
Add, view, and update projects for different products.
Maintain project details: FG Part Number, PCBA Part Number, Start/End Dates, BOM, NPI Engineer.
Record MES information: LOT ID, Workflows, Work Orders, PO Numbers.
Manage tables:

Build matrix (Component / Make)
Assembly matrix (Drawing / Drawing Name)
Machine matrix (Machine Name / Program Name)
Load projects from database or fallback to Excel if DB record not found.
Initialize and manage project checklists.
User-friendly PyQt6 GUI.
Handover Documents Manager
Handover Checklist Manager

**MODULES TO INSTALL** - pip install PyQt6 pandas openpyxl python-docx reportlab
| Module / Import                                                  | Package to Install                         |
| ---------------------------------------------------------------- | ------------------------------------------ |
| `PyQt6.*`                                                        | PyQt6                                      |
| `pandas`                                                         | pandas                                     |
| `openpyxl`                                                       | openpyxl                                   |
| `docx` (from `python-docx`)                                      | python-docx                                |
| `reportlab.*`                                                    | reportlab                                  |
| `os, sys, shutil, zipfile, traceback, sqlite3, datetime, typing` | Built-in Python modules, no install needed |


**PROJECT STRUCTURE**

project_management_gui/
│
├─ main.py                   # Entry point of the application
├─ db_manager.py             # Handles SQLite DB operations
├─ ui/
│   ├─ main_window.py        # Main GUI window and tab widgets
│   ├─ handover_tab.py       # Handover documents tab
│   └─ checklist_tab.py      # Checklist tab
├─ resources/
│   ├─ EXCEL_FILE.xlsx       # Optional Excel fallback
│   └─ databse.db   
├─ requirements.txt          # Required Python packages
└─ README.md

