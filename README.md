#Allows you to work interactively with LLMs complex financial models via JSON conversion
# XLSX to JSON Structure & Values Exporter

A Python-based GUI utility to extract data from Excel workbooks into two distinct JSON files: one for the **structure** (formulas and constants) and one for the **cached values**.

## Features
* **Drag-and-Drop:** Powered by `tkinterdnd2` for easy file selection.
* **Dual Export:** * `_structure.json`: Maps out cell formulas, constants, and number formats.
    * `_values.json`: Captures the last calculated values saved in the Excel file.
* **Range Specific:** Target specific ranges (e.g., `A1:DN500`) across all sheets.
* **Defined Names:** Automatically extracts Excel named ranges.

## Getting Started

### Prerequisites
* Python 3.x
* `tkinterdnd2`
* `openpyxl`

### Installation
1. Clone the repo:
   ```bash
   git clone [https://github.com/zarikkhimani/xlsx-to-json-gui.git](https://github.com/zarikkhimani/xlsx-to-json-gui.git)

## Commercial Use & Licensing
This project is licensed under the **PolyForm Noncommercial License 1.0.0**. 

**What this means:**
* **Personal/Academic Use:** You are free to use, study, and modify this tool for personal projects or schoolwork.
* **Commercial Use:** You may **not** use this software for any activities intended for or directed toward commercial advantage or monetary compensation. This includes use within a for-profit corporation or for processing client data in a professional setting.


**Interested in commercial use?** If you would like to use this tool at your firm or integrate it into a commercial product, please contact me to discuss a commercial license.
