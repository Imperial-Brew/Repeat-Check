# Part History Checker

This script checks part numbers from a CSV file against the Made2Manage ERP database to find:
1. Manufacturing history (if we've made it before)
2. Cost information (how much it cost to manufacture)
3. Sales history (when and to whom we've sold it)

All data is limited to the last 5 years.

## Features

- Loads part numbers from a CSV file
- Queries the Made2Manage database for manufacturing history, costs, and sales data
- Calculates average costs based on historical job orders
- Exports results to an Excel file with multiple sheets
- Provides detailed logging

## Requirements

- Python 3.6+
- pandas
- pyodbc
- sqlalchemy
- python-dotenv
- openpyxl

## Usage

Run the script with default parameters:
```
python part_history_checker.py
```

Specify a custom CSV file:
```
python part_history_checker.py path\to\your\input.csv
```

Additional command-line options:
```
python part_history_checker.py --help
usage: part_history_checker.py [-h] [--column PART_COLUMN] [--output OUTPUT_PATH] [--years YEARS] [--part PART_NUMBER] [csv_file]

Check part manufacturing and sales history

positional arguments:
  csv_file              Path to CSV file containing part numbers

options:
  -h, --help            show this help message and exit
  --column PART_COLUMN, -c PART_COLUMN
                        Name of the column containing part numbers (default: part_number)
  --output OUTPUT_PATH, -o OUTPUT_PATH
                        Path to save the output Excel file
  --years YEARS, -y YEARS
                        Number of years of history to retrieve (default: 5)
  --part PART_NUMBER, -p PART_NUMBER
                        Generate detailed summary for a specific part number
```

Generate a detailed summary for a specific part:
```
python part_history_checker.py --part "0020-48796"
```

## Output

The script can produce two types of output depending on how it's run:

### Excel Report (Default)

When run with a CSV file of part numbers, the script generates an Excel file in the `output` directory with the following sheets:

1. **Summary** - Overview of the number of records and unique parts found
2. **Manufacturing History** - Detailed job order history for each part
3. **Sales History** - Sales order history for each part
4. **Cost Analysis** - Standard and average costs for each part

### Part Summary (--part option)

When run with the `--part` option, the script generates a detailed console output for the specified part number:

```
Part # 0020-48796
Athena has built this part 23 times in the past 5 years for an avg cost of $xxxx
the previous 5 sales orders (out of 39 total SOs) were:
OrderDate	OrderedQty	SalesOrderNumber	UnitPrice
12/19/2022	1		050259			0
11/8/2022	1		049958			0
6/16/2022	8		048659			326.88
6/16/2022	42		048659			1716.12
6/16/2022	42		048659			1716.12
```

This format provides a quick overview of the part's manufacturing history, average cost, and recent sales orders.

## Database Queries

The script performs three main queries:

1. **Manufacturing History** - Queries the JOMAST and JOPACT tables to find job orders for the parts
2. **Sales History** - Queries the SOMAST, SOITEM, and SORELS tables to find sales orders for the parts
3. **Average Cost** - Uses a complex query to calculate average costs based on historical job orders

## Environment Variables

The script requires the following environment variables to be set in a `.env` file:

```
DB_DRIVER={SQL Server}
DB_SERVER=your_server_name
DB_NAME=your_database_name
```

## Error Handling

The script includes comprehensive error handling for:
- File not found errors
- Database connection issues
- Query execution errors
- Data processing errors

All errors are logged to both the console and a log file in the `logs` directory.

## Recent Improvements

The script has been enhanced with the following improvements:

1. **Comprehensive Documentation**
   - Added detailed module-level docstring
   - Improved function docstrings with parameter and return value information
   - Added comments explaining complex SQL queries

2. **Progress Indicators**
   - Added progress bars for long-running database queries
   - Improved console output with status messages

3. **Command-line Arguments**
   - Added support for specifying the column name containing part numbers
   - Added option to specify output file path
   - Added option to configure the number of years of history to retrieve
   - Added option to generate a detailed summary for a specific part number

4. **Input Validation**
   - Added validation for CSV file existence
   - Added check for empty part number lists
   - Improved error messages

## Recommendations for Future Updates

1. **Performance Optimizations**
   - Consider using parallel processing for database queries
   - Implement caching for frequently accessed data
   - Optimize SQL queries for better performance

2. **User Interface Improvements**
   - Add a simple GUI for non-technical users
   - Implement interactive data visualization
   - Add ability to filter and sort results

3. **Data Analysis Features**
   - Add trend analysis for costs over time
   - Implement anomaly detection for unusual cost patterns
   - Add forecasting capabilities for future costs

4. **Integration Options**
   - Add support for exporting to other formats (CSV, JSON, etc.)
   - Implement API endpoints for integration with other systems
   - Add email notification capabilities for completed reports
