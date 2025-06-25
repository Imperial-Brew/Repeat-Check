# Part Number Checker

A utility for checking part numbers from RFQs against the Made2Manage ERP database.

## Overview

This tool loads part numbers from a CSV file, queries the Made2Manage database for information about those parts, and saves the results to a CSV file. It's designed to help identify which parts have been previously manufactured.

## Features

- Load part numbers from CSV files
- Query Made2Manage database using Windows authentication
- Process large numbers of part numbers efficiently
- Generate detailed reports with part information
- Configurable input/output paths and column names
- Comprehensive error handling and logging

## Project Structure

```
Repeat-Check/
│
├── data/                      # Original data files (CSV shortcuts)
├── data_files/                # Actual data files (CSV)
│
├── src/                       # Source code
│   ├── database.py            # Database connection and query functions
│   ├── file_handler.py        # CSV file operations
│   └── logger.py              # Logging functionality
│
├── output/                    # Generated reports
│
├── logs/                      # Log files
│
├── .env                       # Environment variables
├── main.py                    # Main entry point
└── README.md                  # This file
```

## Requirements

- Python 3.6+
- pandas
- pyodbc
- python-dotenv

## Installation

1. Clone this repository
2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```
3. Ensure your `.env` file is configured with the correct database connection parameters:
   ```
   DB_DRIVER={SQL Server}
   DB_SERVER=your_server_name
   DB_NAME=your_database_name
   ```

## Usage

Run the script with default parameters:
```
python main.py
```

Specify custom input and output files:
```
python main.py --input path\to\your\input.csv --output path\to\your\output.csv
```

Specify a different column name for part numbers:
```
python main.py --column your_part_number_column
```

Set logging level:
```
python main.py --log-level DEBUG
```

Disable logging to file:
```
python main.py --no-log-file
```

## Command Line Arguments

| Argument | Short | Default | Description |
|----------|-------|---------|-------------|
| `--input` | `-i` | `data\quote_items_7900_7950_complete.csv` | Input CSV file |
| `--output` | `-o` | `output\matched_parts_output.csv` | Output CSV file |
| `--column` | `-c` | `part_number` | Column name containing part numbers |
| `--log-level` | `-l` | `INFO` | Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL) |
| `--no-log-file` | | | Disable logging to file |

## License

This project is proprietary and confidential.