"""
Part History Checker

This script analyzes manufacturing and sales history for a list of parts from a CSV file.
It queries a database for manufacturing history, sales history, and cost analysis data,
then compiles the results into a multi-sheet Excel report.

The script can also generate a detailed summary for a specific part number, showing
manufacturing metrics, sales history, and business metrics like margins and risk assessments.

Usage:
    python part_history_checker.py [csv_file_path]

    If csv_file_path is not provided, it defaults to '../data/quote_items_7000_8067_complete.csv'

    For a detailed summary of a specific part:
    python part_history_checker.py --part "0020-48796"

    For a detailed summary in JSON format:
    python part_history_checker.py --part "0020-48796" --json

    For processing parts in batches:
    python part_history_checker.py --batch 5

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

    - When using --part with --json option: JSON output with comprehensive metrics:
      - Basic part information (part number, current revision)
      - Manufacturing metrics (total builds, builds by revision, average costs)
      - Sales information (recent sales orders, annual revenue)
      - RFQ information from the CSV file
      - Calculated business metrics (margins, risk assessments)

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
            ),
            encoding='utf-8'
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
                WITH RankedReleases AS (
                    SELECT 
                        R.FSONO, 
                        R.FENUMBER, 
                        R.FNETPRICE,
                        ROW_NUMBER() OVER (PARTITION BY R.FSONO, R.FENUMBER ORDER BY R.FRELEASE DESC) AS ReleaseRank
                    FROM SORELS R
                )
                SELECT
                    S.FSONO    AS SalesOrderNumber,
                    S.FCUSTNO  AS CustomerNumber,
                    S.FCOMPANY AS CustomerName,
                    I.FPARTNO  AS PartNumber,
                    I.FPARTREV AS Revision,
                    I.FCITEMSTATUS AS ItemStatus,
                    I.FQUANTITY    AS OrderedQty,
                    CASE 
                        WHEN I.FQUANTITY=0 THEN 0 
                        ELSE R.FNETPRICE/I.FQUANTITY 
                    END AS UnitPrice,
                    R.FNETPRICE AS TotalValue,
                    S.FORDERDATE   AS OrderDate
                FROM SOMAST S
                JOIN SOITEM I ON S.FSONO=I.FSONO
                JOIN RankedReleases R ON S.FSONO=R.FSONO AND I.FENUMBER=R.FENUMBER AND R.ReleaseRank=1
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

def generate_part_summary_dict(engine, part_number, csv_data=None):
    """
    Generate a detailed summary dictionary for a specific part number.

    Creates a dictionary with comprehensive information about:
    - Basic part information (part number, current revision)
    - Manufacturing metrics (total builds, builds by revision, average costs)
    - Sales information (recent sales orders, annual revenue)
    - Calculated business metrics (margins, risk assessments)

    Args:
        engine (sqlalchemy.engine.Engine): Database connection engine
        part_number (str): The part number to generate summary for
        csv_data (pandas.DataFrame, optional): DataFrame containing RFQ data

    Returns:
        dict: Dictionary containing detailed part metrics
    """
    # Query manufacturing history for this part
    manu_df = query_part_manufacturing_history(engine, [part_number])

    # Query cost information for this part
    cost_df = query_part_average_cost(engine, [part_number])

    # Query sales history for this part
    sales_df = query_part_sales_history(engine, [part_number])

    # Get the revision from the CSV file instead of SQL data
    csv_revision = "05"  # Default to 05 as specified in notes.txt
    if csv_data is not None:
        # Filter CSV data for this part number
        part_rows = csv_data[csv_data['part_number'] == part_number]
        if not part_rows.empty and 'revision' in part_rows.columns:
            # Use the first revision found
            if not part_rows['revision'].isna().all():
                csv_revision = part_rows['revision'].iloc[0]

    result = {
        "PartNumber": part_number,
        "CurrentRevision": csv_revision,
    }

    # Manufacturing metrics
    total_builds = len(manu_df) if not manu_df.empty else 0
    result["TotalBuilds"] = total_builds

    # Calculate builds by revision
    builds_by_revision = {}
    if not manu_df.empty and 'Revision' in manu_df.columns:
        rev_counts = manu_df['Revision'].value_counts().to_dict()
        builds_by_revision = {str(rev): count for rev, count in rev_counts.items()}
    result["BuildsByRevision"] = builds_by_revision

    # Average costs - Fix the calculation to ensure it's not NaN
    avg_cost_all_revs = 0
    if not cost_df.empty and 'Average_Cost' in cost_df.columns:
        valid_costs = cost_df['Average_Cost'].dropna()
        if not valid_costs.empty:
            avg_cost_all_revs = float(valid_costs.mean())

    # Recent standard cost (from most recent job)
    recent_std_cost = 0
    if not cost_df.empty and 'StandardCost' in cost_df.columns:
        valid_costs = cost_df['StandardCost'].dropna()
        if not valid_costs.empty:
            recent_std_cost = float(valid_costs.iloc[0])

    # Recent sales orders
    recent_sales_orders = []
    if not sales_df.empty:
        # Sort by OrderDate descending to get the most recent orders
        recent_sales = sales_df.sort_values('OrderDate', ascending=False).head(5)

        for _, row in recent_sales.iterrows():
            order_date = row['OrderDate'].strftime('%Y-%m-%d') if pd.notna(row['OrderDate']) else None
            ordered_qty = int(row['OrderedQty']) if pd.notna(row['OrderedQty']) else 0
            sales_order = row['SalesOrderNumber'] if pd.notna(row['SalesOrderNumber']) else None
            total_value = float(row['TotalValue']) if pd.notna(row['TotalValue']) else 0

            # Calculate unit price as TotalValue/Qty
            unit_price = 0
            if ordered_qty > 0:
                unit_price = total_value / ordered_qty

            recent_sales_orders.append({
                "OrderDate": order_date,
                "Qty": ordered_qty,
                "SONumber": sales_order,
                "UnitPrice": round(unit_price, 2)
            })
    result["RecentSalesOrders"] = recent_sales_orders

    # Annual revenue
    annual_revenue = {}
    current_year = datetime.now().year
    # Initialize with zeros for the last 6 years
    for year in range(current_year - 5, current_year + 1):
        annual_revenue[year] = 0.0

    if not sales_df.empty and 'OrderDate' in sales_df.columns and 'TotalValue' in sales_df.columns:
        # Add year column to sales_df
        sales_df['Year'] = sales_df['OrderDate'].dt.year
        # Group by year and sum TotalValue
        yearly_revenue = sales_df.groupby('Year')['TotalValue'].sum().to_dict()
        # Update annual_revenue with actual values
        for year, revenue in yearly_revenue.items():
            if year in annual_revenue:
                annual_revenue[year] = float(revenue)
    result["AnnualRevenue"] = annual_revenue

    # Average annual revenue
    avg_annual_revenue = sum(annual_revenue.values()) / 6  # Average over 6 years
    result["AvgAnnualRevenue"] = round(avg_annual_revenue, 2)

    # RFQ quantity
    rfq_qty = 0
    if csv_data is not None:
        # Filter CSV data for this part number
        part_rows = csv_data[csv_data['part_number'] == part_number]
        if not part_rows.empty and 'quantity' in part_rows.columns:
            # Use the first quantity found (or could use max, sum, etc.)
            rfq_qty = int(part_rows['quantity'].iloc[0])
    result["RFQQty"] = rfq_qty

    # Recent sales metrics - Find the most recent non-zero unit price
    recent_so_qty = 0
    recent_so_date = None
    recent_unit_price = 0.0

    if recent_sales_orders:
        # Look for the first non-zero unit price
        for sale in recent_sales_orders:
            if sale["UnitPrice"] > 0:
                recent_so_qty = sale["Qty"]
                recent_so_date = sale["OrderDate"]
                recent_unit_price = sale["UnitPrice"]
                break

    result["RecentSOQty"] = int(recent_so_qty)
    result["RecentSODate"] = recent_so_date
    result["RecentUnitPrice"] = round(recent_unit_price, 2)
    result["SQLCost"] = round(avg_cost_all_revs, 2)  # Moved from above
    result["recentSTDcost"] = round(recent_std_cost, 2)  # Moved from above

    # Calculated business metrics
    potential_revenue = rfq_qty * recent_unit_price
    result["PotentialRevenue"] = round(potential_revenue, 2)

    # Calculate estimated margin as 1 - (SQLCost / recent unit price)
    estimated_margin = 0
    if recent_unit_price > 0:
        estimated_margin = 1 - (avg_cost_all_revs / recent_unit_price)
    result["estimatedmargin"] = round(estimated_margin, 2)

    # Risk assessments
    # Define risk thresholds
    potential_thresholds = [1000, 10000, 50000]  # New thresholds
    avg_annual_thresholds = [10000, 50000, 250000]   # New thresholds

    # Determine risk by potential revenue
    if potential_revenue < potential_thresholds[0]:
        risk_by_potential = "Low"
    elif potential_revenue < potential_thresholds[1]:
        risk_by_potential = "Medium"
    elif potential_revenue < potential_thresholds[2]:
        risk_by_potential = "High"
    else:
        risk_by_potential = "Very High"

    # Determine risk by average annual revenue
    if avg_annual_revenue < avg_annual_thresholds[0]:
        risk_by_avg_annual = "Low"
    elif avg_annual_revenue < avg_annual_thresholds[1]:
        risk_by_avg_annual = "Medium"
    elif avg_annual_revenue < avg_annual_thresholds[2]:
        risk_by_avg_annual = "High"
    else:
        risk_by_avg_annual = "Very High"

    result["RiskByPotential"] = risk_by_potential
    result["RiskByAvgAnnual"] = risk_by_avg_annual

    # Round all numeric values in the result dictionary to 2 decimal places
    for key, value in result.items():
        if isinstance(value, float):
            result[key] = round(value, 2)

    return result

def generate_part_summary(engine, part_number, csv_data=None):
    """
    Generate a detailed summary for a specific part number.

    Creates a formatted string with information about:
    - How many times the part has been built in the past 5 years
    - Average cost of manufacturing
    - Previous 5 sales orders with details

    Args:
        engine (sqlalchemy.engine.Engine): Database connection engine
        part_number (str): The part number to generate summary for
        csv_data (pandas.DataFrame, optional): DataFrame containing RFQ data

    Returns:
        str: Formatted summary text for the part
    """
    # Get the detailed summary dictionary
    summary_dict = generate_part_summary_dict(engine, part_number, csv_data)

    # Format the sales history for display
    sales_history = "OrderDate\tOrderedQty\tSalesOrderNumber\tUnitPrice\n"
    for sale in summary_dict["RecentSalesOrders"]:
        order_date = sale["OrderDate"] if sale["OrderDate"] else "N/A"
        ordered_qty = sale["Qty"]
        sales_order = sale["SONumber"] if sale["SONumber"] else "N/A"
        unit_price = sale["UnitPrice"]

        sales_history += f"{order_date}\t{ordered_qty}\t\t{sales_order}\t\t\t{unit_price:.2f}\n"

    # Format the summary text
    summary = f"""
Part # {summary_dict["PartNumber"]}
Athena has built this part {summary_dict["TotalBuilds"]} times in the past 5 years for an avg cost of ${summary_dict["SQLCost"]:.2f}
the previous 5 sales orders (out of {len(summary_dict["RecentSalesOrders"])} total SOs) were:
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
                        default=os.path.join('..', 'data', 'quote_items_7000_8067_complete.csv'),
                        help='Path to CSV file containing part numbers')
    parser.add_argument('--column', '-c', dest='part_column', default='part_number',
                        help='Name of the column containing part numbers (default: part_number)')
    parser.add_argument('--output', '-o', dest='output_path',
                        help='Path to save the output Excel file')
    parser.add_argument('--years', '-y', type=int, default=5,
                        help='Number of years of history to retrieve (default: 5)')
    parser.add_argument('--part', '-p', dest='part_number',
                        help='Generate detailed summary for a specific part number (use with --json for comprehensive metrics)')
    parser.add_argument('--json', '-j', action='store_true',
                        help='Output part summary as JSON (only with --part)')
    parser.add_argument('--batch', '-b', type=int, default=0,
                        help='Process parts in batches of specified size (default: process all at once)')
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

            # Load CSV data for RFQ information
            csv_data = None
            try:
                csv_data = pd.read_csv(args.csv_file)
            except Exception as e:
                logging.warning(f"Could not load CSV data for RFQ information: {e}")
                print(f"\n⚠️ Warning: Could not load CSV data for RFQ information: {e}")

            if args.json:
                # Generate and output JSON summary
                summary_dict = generate_part_summary_dict(engine, args.part_number, csv_data)
                import json
                json_summary = json.dumps(summary_dict, indent=2, default=str)
                print(json_summary)

                # Save JSON to file
                output_dir = os.path.join('..', 'output')
                os.makedirs(output_dir, exist_ok=True)
                json_file = os.path.join(output_dir, f"{args.part_number}_summary.json")
                with open(json_file, 'w') as f:
                    f.write(json_summary)
                logging.info(f"✅ JSON part summary generated for {args.part_number} and saved to {json_file}")
                print(f"\n✅ JSON summary saved to '{json_file}'")
            else:
                # Generate and output text summary
                summary = generate_part_summary(engine, args.part_number, csv_data)
                print(summary)
                logging.info(f"✅ Part summary generated for {args.part_number}")

            return 0

        # Query database for part history
        print("\nQuerying database for part history...")

        # Process parts in batches if batch size is specified
        if args.batch > 0:
            print(f"Processing parts in batches of {args.batch}...")
            # Divide parts into batches
            batches = list(chunk(part_numbers, args.batch))

            # Initialize empty DataFrames for results
            manu_df = pd.DataFrame()
            sales_df = pd.DataFrame()
            cost_df = pd.DataFrame()

            # Process each batch
            for i, batch in enumerate(batches):
                print(f"\nProcessing batch {i+1} of {len(batches)} ({len(batch)} parts)...")

                # Query database for this batch
                batch_manu_df = query_part_manufacturing_history(engine, batch)
                batch_sales_df = query_part_sales_history(engine, batch)
                batch_cost_df = query_part_average_cost(engine, batch)

                # Combine results
                manu_df = pd.concat([manu_df, batch_manu_df], ignore_index=True) if not batch_manu_df.empty else manu_df
                sales_df = pd.concat([sales_df, batch_sales_df], ignore_index=True) if not batch_sales_df.empty else sales_df
                cost_df = pd.concat([cost_df, batch_cost_df], ignore_index=True) if not batch_cost_df.empty else cost_df

                print(f"Batch {i+1} complete. Found {len(batch_manu_df)} manufacturing records, {len(batch_sales_df)} sales records, and {len(batch_cost_df)} cost records.")
        else:
            # Process all parts at once (original behavior)
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
