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

-----------------------------------------------

# GenericXMLManager 通用 XML 管理工具

一个轻量级的 Python 图形界面工具，用于管理和编辑结构化的 XML 数据集。

📁 **通用 XML 数据库控制器**  
这是一个使用 Python 构建的桌面应用程序，支持通过用户友好的表格界面来管理结构化 XML 文件。适用于各类场景中的通用数据编辑和管理。

## ✨ 功能特色

- 表格形式展示 XML 数据
- 实时添加、删除、编辑记录
- 支持通过下拉菜单批量添加记录
- 支持字段筛选和过滤
- 支持格式化保存回原始 XML 文件
- 可自定义表格名称和结构
- 注重可拓展性与可移植性

## 🛠 技术栈

- Python 3.9+
- `tkinter` 作为 GUI 框架
- `xml.etree.ElementTree` 和 `minidom` 进行 XML 解析
- 使用 `PyInstaller` 打包为可执行文件

## 📁 项目结构
your-project/
├── data/ # 请将 XML 文件放置在此文件夹中
├── XMLDatabaseApp.py # 主程序文件
└── README.md

## 🚀 使用方式

1. 将结构化的 XML 文件放入 `data/` 文件夹中  
2. 使用 Python 或已打包的 `.exe` 文件启动应用  
3. 在界面中交互式管理表格，编辑并保存更改

## 📦 Windows 打包方式（可选）

如需将项目打包为 `.exe` 可执行程序：

pyinstaller --noconsole --add-data "data;data" XMLDatabaseApp.py

⚠️ 请确保打包后的 .exe 文件与 data/ 文件夹处于同一目录。

📄 许可协议
MIT License — 可自由使用与修改。
