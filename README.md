# GenericXMLManager
A lightweight Python GUI tool for managing and editing structured XML datasets

🗂️ Universal XML Database Controller
A desktop application built with Python for managing structured XML files through a user-friendly table interface. Ideal for general-purpose use across various domains involving XML data.

✨ Features
Tabular display of XML datasets

Add, delete, and edit records with real-time GUI interaction

Batch record addition with dropdown field population

Filter and search capabilities for specific fields

Supports saving back to original XML with formatting

Customizable table names and structure

Designed for extensibility and portability

🛠️ Tech Stack
Python 3.9+

tkinter for GUI

xml.etree.ElementTree, minidom for XML processing

Packaged using PyInstaller

📦 File Structure
your-project/
├── data/                  # Place your XML files here
├── XMLDatabaseApp.py      # Main application
└── README.md

🚀 How to Use
Place your structured XML files in the data/ folder

Launch the application via Python or use the packaged .exe

Interact with tables, perform edits, and save changes

📦 Packaging (Windows)
If you want to share the application as an executable:
pyinstaller --noconsole --add-data "data;data" XMLDatabaseApp.py
Make sure the data/ folder stays alongside the .exe file.

📄 License
MIT License – free to use and modify.
