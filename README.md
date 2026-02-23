# Excel Sync Master

This Python script monitors a source Excel file for changes and automatically synchronizes specific cell and table data across multiple target Excel workbooks upon saving. It uses `watchdog` to detect file modifications and `openpyxl` for reading and writing Excel data.

## Features
* **Real-time Monitoring:** Instantly detects when the source file is saved.
* **Smart Debouncing:** Prevents the script from firing multiple times or crashing due to OS-level temporary file saves.
* **Custom Data Mapping:** Maps specific cells and dynamically iterating table rows to different target coordinates depending on the workbook.

## Prerequisites

You need Python installed on your system, along with the following third-party libraries:

```bash
pip install -r requirements.txt
```

## Usage

1. **Configure the source file.** In `src/main.py`, set `SOURCE_FILE_NAME` to the name of the Excel file you want to monitor.

2. **Define your cell mappings.** Edit the `general_info_mapping` dictionary to map source cells to target cells. One-to-many mappings are supported by using a list as the value:
   ```python
   general_info_mapping = {
       "B1": "C3",            # One-to-one
       "B2": ["G3", "H3"],    # One-to-many
   }
   ```

3. **Define your table mappings.** Edit `tracklist_table_mapping` to map source column indices to target column indices. Set a value to `None` to skip a column.

4. **Add target workbooks.** Add entries to the `TARGETS` list. Each entry specifies the target filename, sheet name, row offset, and which mappings to use.

5. **Run the script** from the directory containing your Excel files:
   ```bash
   python src/main.py
   ```

6. The script will watch for changes to the source file. Every time you save it, the configured targets will be updated automatically. Press `Ctrl+C` to stop.

## Project Structure

```
excel/
├── src/
│   └── main.py        # Main script with configuration and sync logic
├── .gitignore
├── requirements.txt
└── README.md
```