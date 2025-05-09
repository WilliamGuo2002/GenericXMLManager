# GenericXMLManager
A lightweight Python GUI tool for managing and editing structured XML datasets
📂 General XML Database Controller
A user-friendly Python desktop application for managing structured XML datasets with tabular interfaces. Designed to support read, edit, add, delete, and filter operations across multiple XML tables.

🔧 Features
GUI-based data management for multiple XML tables

Real-time editing, adding, and deleting of XML nodes

Support for GB2312 encoding (e.g., Chinese datasets)

Dropdown menus dynamically populated from related tables

Filter and search functionality (e.g., by EmployeeID, ItemID)

Batch addition support for summary tables

Safe saving and XML pretty-print formatting

Designed to be repurposable for any structured XML schema

🧪 Tech Stack
Python 3.9+

tkinter (built-in GUI)

xml.etree.ElementTree and minidom for XML parsing and formatting

Packaged with PyInstaller for distribution

📁 Structure


📦 Packaging & Distribution
Built using PyInstaller:
pyinstaller --noconsole --add-data "data;data" SRRCDataBrowser.py
The executable works as long as the data/ folder is in the same directory.

💡 Customization
To generalize:

Rename XML files and table labels to Table1, Table2, etc.

Update code logic to reflect your schema.

📃 License
This project is licensed under the MIT License.
