"""
Part History Checker

This script analyzes manufacturing and sales history for a list of parts from a CSV file.
It queries a database for manufacturing history, sales history, and cost analysis data,
then compiles the results into a multi-sheet Excel report.

The script can also generate a detailed summary for a specific part number, showing
how many times it has been built, the average cost, and the previous 5 sales orders.

Usage:
    python part_history_checker.py [csv_file_path]

    If csv_file_path is not provided, it defaults to '../data/quote_items_7900_7950_complete.csv'

    For a detailed summary of a specific part:
    python part_history_checker.py --part "0020-48796"

Output:
    - When processing multiple parts: Excel file with four sheets:
      - Summary: Overview of records found for each category
      - Manufacturing History: Details of manufacturing jobs for the parts
      - Sales History: Details of sales orders for the parts
      - Cost Analysis: Cost information including standard and average costs

    - When using --part option: Console output with a detailed summary showing:
      - Number of times the part has been built in the past 5 years
      - Average manufacturing cost
      - Previous 5 sales orders with details (date, quantity, order number, price)

Requirements:
    - .env file with database connection parameters (DB_DRIVER, DB_SERVER, DB_NAME)
    - pandas, sqlalchemy, pyodbc, openpyxl packages
    - CSV file with part numbers (default column name: 'part_number')

Author: Dustin Drab
Date: June 2025
"""

import os
import sys
import pandas as pd
import logging
from datetime import datetime
from dotenv import load_dotenv
from sqlalchemy import create_engine
from tqdm import tqdm

# Create logs directory if it doesn't exist
script_dir = os.path.dirname(os.path.abspath(__file__))
logs_dir = os.path.join(script_dir, 'logs')
os.makedirs(logs_dir, exist_ok=True)

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(
            os.path.join(
                logs_dir,
                f"part_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            )
        ),
        logging.StreamHandler()
    ]
)

# Load environment variables
load_dotenv()

def connect_to_database():
    """
    Establish connection to the M2M database using Windows authentication.

    Requires environment variables:
    - DB_DRIVER: Database driver (e.g., 'SQL Server')
    - DB_SERVER: Server name
    - DB_NAME: Database name

    Returns:
        sqlalchemy.engine.Engine: Database engine object for executing queries

    Raises:
        Exception: If connection fails due to missing environment variables or connection issues
    """
    try:
        conn_str = (
            f"DRIVER={os.getenv('DB_DRIVER')};"
            f"SERVER={os.getenv('DB_SERVER')};"
            f"DATABASE={os.getenv('DB_NAME')};"
            f"Trusted_Connection=yes;"
        )
        logging.info(f"Connecting to database {os.getenv('DB_NAME')} on server {os.getenv('DB_SERVER')}")
        connection_url = f"mssql+pyodbc:///?odbc_connect={conn_str.replace(';', '%3B')}"
        engine = create_engine(connection_url)
        logging.info("Database connection successful")
        return engine

    except Exception as e:
        logging.error(f"Database connection failed: {e}")
        raise

def load_part_numbers(csv_path, part_number_column='part_number'):
    """
    Load part numbers from a CSV file.

    Args:
        csv_path (str): Path to the CSV file containing part numbers
        part_number_column (str, optional): Name of the column containing part numbers.
                                           Defaults to 'part_number'.

    Returns:
        list: List of unique part numbers

    Raises:
        FileNotFoundError: If the CSV file doesn't exist
        ValueError: If the specified part_number_column is not found in the CSV
    """
    logging.info(f"Loading data from {csv_path}")
    if not os.path.exists(csv_path):
        logging.error(f"CSV file not found: {csv_path}")
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    df = pd.read_csv(csv_path)
    if part_number_column not in df.columns:
        available = ', '.join(df.columns)
        logging.error(f"Column '{part_number_column}' not found. Available: {available}")
        raise ValueError(f"Column '{part_number_column}' not found in CSV")
    part_numbers = df[part_number_column].dropna().unique().tolist()
    logging.info(f"Loaded {len(df)} rows, found {len(part_numbers)} unique part numbers")
    return part_numbers

def chunk(lst, size=1000):
    """
    Break a list into chunks of specified size.

    This is a generator function that yields smaller chunks of a large list
    to avoid memory issues when processing large datasets.

    Args:
        lst (list): The list to be chunked
        size (int, optional): Maximum size of each chunk. Defaults to 1000.

    Yields:
        list: A chunk of the original list with at most 'size' elements
    """
    for i in range(0, len(lst), size):
        yield lst[i:i + size]

def query_part_manufacturing_history(engine, part_numbers):
    """
    Query the database for part manufacturing history.

    Retrieves manufacturing job information for the specified part numbers,
    including job details, costs, and unit costs.

    Args:
        engine (sqlalchemy.engine.Engine): Database connection engine
        part_numbers (list): List of part numbers to query

    Returns:
        pandas.DataFrame: DataFrame containing manufacturing history data

    Raises:
        Exception: If the database query fails
    """
    if not part_numbers:
        logging.warning("No part numbers provided for manufacturing history")
        return pd.DataFrame()
    results = []
    try:
        # Process part numbers in chunks to avoid query size limitations
        chunks = list(chunk(part_numbers))
        for part_chunk in tqdm(chunks, desc="Manufacturing History", unit="chunk"):
            part_list = ",".join(f"'{p}'" for p in part_chunk)
            logging.info(f"Querying manufacturing history for {len(part_chunk)} parts")
            query = f"""
                SELECT
                    jm.fjobno   AS JobNumber,
                    jm.fpartno  AS PartNumber,
                    jm.fpartrev AS Revision,
                    jm.fddue_date AS DueDate,
                    jm.fquantity  AS Quantity,
                    jm.fcus_id    AS Customer,
                    jm.fstatus    AS Status,
                    jm.fact_rel   AS ReleaseDate,
                    jp.flabact    AS Labor,
                    jp.FMATLACT   AS Material,
                    jp.FOVHDACT   AS Overhead,
                    jp.FSETUPACT  AS Setup,
                    jp.FSUBACT    AS Subcontract,
                    jp.FOTHRACT   AS Other,
                    (jp.flabact+jp.FMATLACT+jp.FOVHDACT+jp.FSETUPACT+jp.FSUBACT+jp.FOTHRACT)
                        AS TotalCost,
                    CASE
                        WHEN jm.fquantity=0 THEN NULL
                        ELSE (jp.flabact+jp.FMATLACT+jp.FOVHDACT+jp.FSETUPACT+jp.FSUBACT+jp.FOTHRACT)
                             / jm.fquantity
                    END AS UnitCost
                FROM JOMAST jm
                LEFT JOIN JOPACT jp ON jm.fjobno=jp.fjobno
                WHERE jm.fpartno IN ({part_list})
                  AND jm.fact_rel >= DATEADD(YEAR, -5, GETDATE())
                  AND jm.fstatus IN ('CLOSED','RELEASED')
                ORDER BY jm.fpartno, jm.fact_rel DESC
            """
            conn = engine.raw_connection()
            try:
                df_chunk = pd.read_sql(query, conn)
            finally:
                conn.close()
            logging.info(f"Manufacturing query returned {len(df_chunk)} rows")
            results.append(df_chunk)
        return pd.concat(results, ignore_index=True) if results else pd.DataFrame()
    except Exception as e:
        logging.error(f"Manufacturing history query failed: {e}")
        raise

def query_part_sales_history(engine, part_numbers):
    """
    Query the database for part sales history.

    Retrieves sales order information for the specified part numbers,
    including customer details, quantities, prices, and order dates.

    Args:
        engine (sqlalchemy.engine.Engine): Database connection engine
        part_numbers (list): List of part numbers to query

    Returns:
        pandas.DataFrame: DataFrame containing sales history data

    Raises:
        Exception: If the database query fails
    """
    if not part_numbers:
        logging.warning("No part numbers provided for sales history")
        return pd.DataFrame()
    results = []
    try:
        # Process part numbers in chunks to avoid query size limitations
        chunks = list(chunk(part_numbers))
        for part_chunk in tqdm(chunks, desc="Sales History", unit="chunk"):
            part_list = ",".join(f"'{p}'" for p in part_chunk)
            logging.info(f"Querying sales history for {len(part_chunk)} parts")
            query = f"""
                SELECT
                    S.FSONO    AS SalesOrderNumber,
                    S.FCUSTNO  AS CustomerNumber,
                    S.FCOMPANY AS CustomerName,
                    I.FPARTNO  AS PartNumber,
                    I.FPARTREV AS Revision,
                    I.FCITEMSTATUS AS ItemStatus,
                    I.FQUANTITY    AS OrderedQty,
                    CASE WHEN I.FQUANTITY=0 THEN 0 ELSE R.FNETPRICE/I.FQUANTITY END AS UnitPrice,
                    R.FNETPRICE    AS TotalValue,
                    S.FORDERDATE   AS OrderDate
                FROM SOMAST S
                JOIN SOITEM I  ON S.FSONO=I.FSONO
                JOIN SORELS R  ON I.FSONO=R.FSONO AND I.FENUMBER=R.FENUMBER
                WHERE I.FPARTNO IN ({part_list})
                  AND S.FORDERDATE >= DATEADD(YEAR, -5, GETDATE())
                ORDER BY I.FPARTNO, S.FORDERDATE DESC
            """
            conn = engine.raw_connection()
            try:
                df_chunk = pd.read_sql(query, conn)
            finally:
                conn.close()
            logging.info(f"Sales query returned {len(df_chunk)} rows")
            results.append(df_chunk)
        return pd.concat(results, ignore_index=True) if results else pd.DataFrame()
    except Exception as e:
        logging.error(f"Sales history query failed: {e}")
        raise

def query_part_average_cost(engine, part_numbers):
    """
    Query the database for part average cost information.

    Calculates average manufacturing costs for the specified part numbers
    based on recent job history, excluding outliers.

    Args:
        engine (sqlalchemy.engine.Engine): Database connection engine
        part_numbers (list): List of part numbers to query

    Returns:
        pandas.DataFrame: DataFrame containing cost analysis data including
                         standard costs and calculated average costs

    Raises:
        Exception: If the database query fails
    """
    if not part_numbers:
        logging.warning("No part numbers provided for average cost")
        return pd.DataFrame()
    results = []
    try:
        # Process part numbers in chunks to avoid query size limitations
        chunks = list(chunk(part_numbers))
        for part_chunk in tqdm(chunks, desc="Cost Analysis", unit="chunk"):
            part_list = ",".join(f"'{p}'" for p in part_chunk)
            logging.info(f"Querying average cost for {len(part_chunk)} parts")
            # This complex query calculates average costs while excluding outliers
            query = f"""
                SELECT
                    m.fpartno   AS PartNumber,
                    m.frev      AS Revision,
                    m.fdescript AS Description,
                    m.fstdcost  AS StandardCost,  -- Standard cost from inventory master
                    subq.Average_Cost,            -- Calculated average cost
                    subq.JobCount                 -- Number of jobs used in calculation
                FROM INMAST m
                LEFT JOIN (
                    -- Calculate average cost from filtered job costs
                    SELECT tmp2.fpartno, tmp2.fpartrev,
                           AVG(tmp2.total_cost) AS Average_Cost,  -- Average of unit costs
                           COUNT(tmp2.fpartno)  AS JobCount       -- Count of jobs used
                    FROM (
                        -- Calculate unit cost and rank jobs by cost
                        SELECT m.fjobno, m.fpartno, m.fpartrev, m.fact_rel,
                               -- Calculate unit cost (total cost / quantity)
                               CASE WHEN m.fquantity=0 THEN NULL ELSE
                                    (a.fmatlact+a.fsubact+a.fsetupact+a.flabact+a.fovhdact+a.fothract)
                                    / m.fquantity END   AS total_cost,
                               -- Rank jobs by unit cost to identify outliers
                               ROW_NUMBER() OVER (
                                   PARTITION BY m.fpartno
                                   ORDER BY CASE WHEN m.fquantity=0 THEN 0 ELSE
                                        (a.fmatlact+a.fsubact+a.fsetupact+a.flabact
                                         +a.fovhdact+a.fothract)/m.fquantity END
                               ) AS rn
                        FROM JOMAST m
                        JOIN JOPACT a ON m.fjobno=a.fjobno
                        JOIN (
                            -- Get the 10 most recent jobs for each part
                            SELECT m.fjobno, m.fpartno, m.fpartrev, m.fact_rel,
                                   ROW_NUMBER() OVER (
                                       PARTITION BY m.fpartno
                                       ORDER BY m.fact_rel DESC  -- Sort by release date descending
                                   ) AS rn1
                            FROM JOMAST m
                            JOIN JOPACT a ON m.fjobno=a.fjobno
                            WHERE m.fstatus='closed'             -- Only include closed jobs
                              AND m.fquantity<>0                 -- Avoid division by zero
                              AND m.fact_rel>DATEADD(YEAR,-5,GETDATE())  -- Last 5 years only
                        ) tmp_filtered ON tmp_filtered.fjobno=m.fjobno
                        WHERE tmp_filtered.rn1 <= 10  -- Limit to 10 most recent jobs
                    ) tmp2
                    WHERE tmp2.rn BETWEEN 2 AND 9  -- Exclude lowest and highest cost jobs (outliers)
                    GROUP BY tmp2.fpartno, tmp2.fpartrev
                ) subq
                  ON subq.fpartno=m.fpartno AND subq.fpartrev=m.frev
                WHERE m.fpartno IN ({part_list})
                ORDER BY m.fpartno
            """
            conn = engine.raw_connection()
            try:
                df_chunk = pd.read_sql(query, conn)
            finally:
                conn.close()
            logging.info(f"Average cost query returned {len(df_chunk)} rows")
            results.append(df_chunk)
        return pd.concat(results, ignore_index=True) if results else pd.DataFrame()
    except Exception as e:
        logging.error(f"Average cost query failed: {e}")
        raise

def generate_part_summary(engine, part_number):
    """
    Generate a detailed summary for a specific part number.

    Creates a formatted string with information about:
    - How many times the part has been built in the past 5 years
    - Average cost of manufacturing
    - Previous 5 sales orders with details

    Args:
        engine (sqlalchemy.engine.Engine): Database connection engine
        part_number (str): The part number to generate summary for

    Returns:
        str: Formatted summary text for the part
    """
    # Query manufacturing history for this part
    manu_df = query_part_manufacturing_history(engine, [part_number])
    job_count = len(manu_df) if not manu_df.empty else 0

    # Query cost information for this part
    cost_df = query_part_average_cost(engine, [part_number])
    avg_cost = cost_df['Average_Cost'].iloc[0] if not cost_df.empty and not cost_df['Average_Cost'].isna().all() else 0

    # Query sales history for this part
    sales_df = query_part_sales_history(engine, [part_number])
    sales_count = len(sales_df) if not sales_df.empty else 0

    # Format the sales history for display (last 5 orders)
    sales_history = ""
    if not sales_df.empty:
        # Sort by OrderDate descending to get the most recent orders
        sales_df = sales_df.sort_values('OrderDate', ascending=False)
        # Take the first 5 rows
        recent_sales = sales_df.head(5)

        # Create a header for the sales history table
        sales_history = "OrderDate\tOrderedQty\tSalesOrderNumber\tUnitPrice\n"

        # Add each row to the sales history
        for _, row in recent_sales.iterrows():
            order_date = row['OrderDate'].strftime('%m/%d/%Y') if pd.notna(row['OrderDate']) else 'N/A'
            ordered_qty = int(row['OrderedQty']) if pd.notna(row['OrderedQty']) else 0
            sales_order = row['SalesOrderNumber'] if pd.notna(row['SalesOrderNumber']) else 'N/A'
            unit_price = row['UnitPrice'] if pd.notna(row['UnitPrice']) else 0

            sales_history += f"{order_date}\t{ordered_qty}\t\t{sales_order}\t\t\t{unit_price}\n"

    # Format the summary text
    summary = f"""
Part # {part_number}
Athena has built this part {job_count} times in the past 5 years for an avg cost of ${avg_cost:.2f}
the previous 5 sales orders (out of {sales_count} total SOs) were:
{sales_history}
"""
    return summary

def save_results(manufacturing_df, sales_df, cost_df, output_path):
    """
    Save query results to an Excel file with multiple sheets.

    Creates an Excel workbook with four sheets:
    - Summary: Overview of record counts and unique part counts
    - Manufacturing History: Manufacturing job details (if available)
    - Sales History: Sales order details (if available)
    - Cost Analysis: Cost information (if available)

    Args:
        manufacturing_df (pandas.DataFrame): Manufacturing history data
        sales_df (pandas.DataFrame): Sales history data
        cost_df (pandas.DataFrame): Cost analysis data
        output_path (str): Path where the Excel file will be saved

    Returns:
        str: Path to the saved Excel file

    Raises:
        Exception: If there's an error creating or saving the Excel file
    """
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        logging.info(f"Saving results to {output_path}")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            summary = {
                'Category': ['Manufacturing History','Sales History','Cost Analysis'],
                'Records': [
                    len(manufacturing_df), len(sales_df), len(cost_df)
                ],
                'Unique Parts': [
                    manufacturing_df['PartNumber'].nunique() if not manufacturing_df.empty else 0,
                    sales_df['PartNumber'].nunique()         if not sales_df.empty         else 0,
                    cost_df['PartNumber'].nunique()          if not cost_df.empty          else 0,
                ]
            }
            pd.DataFrame(summary).to_excel(writer, sheet_name='Summary', index=False)
            if not manufacturing_df.empty:
                manufacturing_df.to_excel(writer, sheet_name='Manufacturing History', index=False)
            if not sales_df.empty:
                sales_df.to_excel(writer, sheet_name='Sales History', index=False)
            if not cost_df.empty:
                cost_df.to_excel(writer, sheet_name='Cost Analysis', index=False)
        logging.info("Results successfully saved")
        return output_path
    except Exception as e:
        logging.error(f"Failed to save results: {e}")
        raise

def main():
    """
    Main function to execute the part history check process.

    Parses command-line arguments, loads part numbers, queries the database,
    and saves results to an Excel file.

    Returns:
        int: 0 for success, 1 for failure
    """
    import argparse

    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Check part manufacturing and sales history')
    parser.add_argument('csv_file', nargs='?', 
                        default=os.path.join('..', 'data', 'quote_items_7900_7950_complete.csv'),
                        help='Path to CSV file containing part numbers')
    parser.add_argument('--column', '-c', dest='part_column', default='part_number',
                        help='Name of the column containing part numbers (default: part_number)')
    parser.add_argument('--output', '-o', dest='output_path',
                        help='Path to save the output Excel file')
    parser.add_argument('--years', '-y', type=int, default=5,
                        help='Number of years of history to retrieve (default: 5)')
    parser.add_argument('--part', '-p', dest='part_number',
                        help='Generate detailed summary for a specific part number')
    args = parser.parse_args()

    # Set output path if not specified
    if not args.output_path:
        output_dir = os.path.join('..', 'output')
        os.makedirs(output_dir, exist_ok=True)
        args.output_path = os.path.join(output_dir, 
                                       f"part_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    engine = None
    try:
        logging.info("Starting part history check process")

        # Validate input file
        if not os.path.exists(args.csv_file):
            raise FileNotFoundError(f"CSV file not found: {args.csv_file}")

        # Load part numbers
        part_numbers = load_part_numbers(args.csv_file, args.part_column)
        if not part_numbers:
            logging.warning("No part numbers found in the CSV file")
            print("\n⚠️ Warning: No part numbers found in the CSV file")
            return 0

        # Connect to database
        engine = connect_to_database()

        # Check if a specific part number was requested
        if args.part_number:
            print(f"\nGenerating detailed summary for part {args.part_number}...")
            summary = generate_part_summary(engine, args.part_number)
            print(summary)
            logging.info(f"✅ Part summary generated for {args.part_number}")
            return 0

        # Query database for part history
        print("\nQuerying database for part history...")
        manu_df = query_part_manufacturing_history(engine, part_numbers)
        sales_df = query_part_sales_history(engine, part_numbers)
        cost_df = query_part_average_cost(engine, part_numbers)

        # Save results
        out_file = save_results(manu_df, sales_df, cost_df, args.output_path)
        logging.info("✅ Process completed successfully")
        print(f"\n✅ Done! Output saved to '{out_file}'")
        return 0

    except (FileNotFoundError, ValueError) as e:
        logging.error(e)
        print(f"\nError: {e}")
        return 1
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        print("\nError: An unexpected error occurred. See log for details.")
        return 1
    finally:
        if engine:
            try:
                engine.dispose()
                logging.info("Database connection closed")
            except Exception as close_err:
                logging.warning(f"Error closing connection: {close_err}")

if __name__ == "__main__":
    sys.exit(main())
