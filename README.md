**About This Project**

Handover Checklist & NPI Project Manager is a Python-based GUI application built using PyQt6 to help manage New Product Introduction (NPI) workflows efficiently. It allows engineers to add, view, and update projects, track MES information, and maintain detailed build, assembly, and machine matrices.

The application stores all project data in a SQLite database, with an optional Excel fallback if the database record is not found. Users can also manage checklists, handover documents, and generate reports in PDF or Word format.

This project was developed as a comprehensive tool for product engineers and NPI teams, simplifying project tracking, documentation, and data management in a single, user-friendly interface.

Key Features:
Add, view, and update project details (FG/PCBA numbers, BOM, dates, NPI engineer)
Track MES entries (LOT ID, Work Orders, PO numbers)
Maintain build, assembly, and machine matrices
Load projects from database or Excel fallback
Manage project-specific checklists
Generate PDF/Word reports of project data

This project demonstrates a practical end-to-end Python GUI solution for engineering workflow management
A Python-based GUI application for managing products and projects in NPI (New Product Introduction) workflows. Supports storing project details, MES entries, build/machine/assembly matrices, and checklists in a SQLite database, with Excel fallback.

Features
Project Management: Add, view, and update project details for different products.
Build Matrix: Track components and their respective makes (Component / Make).
Assembly Matrix: Manage assembly drawings and their names (Drawing / Drawing Name).
Machine Matrix: Maintain machine programs and their associated machines (Machine Name / Program Name).
Database & Excel Integration: Load projects from SQLite database or fallback to Excel if the DB record is not found.
MES Tracking: Record Manufacturing Execution System entries such as LOT IDs, Work Orders, and PO Numbers.
Checklists: Initialize, manage, and update project-specific checklists.
Handover Documents Manager: Organize and maintain project handover documents.
Handover Checklist Manager: Track and manage quality/process checklists.
Reports: Generate PDF or Word reports of project data.
User-Friendly GUI: Built with PyQt6 for a responsive, intuitive interface.

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

