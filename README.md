# excel-unlocker
A simple Python tool to unlock password-protected Excel sheets and enable cell editing.This tool removes sheet protection from Excel files (.xlsx), allowing you to edit previously locked cells. It works on files with sheet-level protection (not file-level encryption).

## 🔧 Installation

### Prerequisites

- Python 3.6 or higher
- pip

### Install Dependencies
```bash
pip install -r requirements.txt
```

Or manually:
```bash
pip install openpyxl
```

## 🚀 Usage

### Basic Usage
```bash
python unlock_excel.py protected_file.xlsx
```

### Example
```bash
python unlock_excel.py Business-plan-exemple-Excel.xlsx
```

**Output:**
```
[+] plan_affaires débloquée
[+] budget débloquée
[+] previsions débloquée
[✓] Sauvegardé: Business-plan-exemple-Excel_UNLOCKED.xlsx
```

The unlocked file will be saved with `_UNLOCKED` suffix.
