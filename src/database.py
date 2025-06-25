import os
import pandas as pd
import pyodbc
import logging
from dotenv import load_dotenv
from sqlalchemy import create_engine

def connect_to_database():
    """Establish connection to the M2M database using Windows authentication.

    Returns:
        SQLAlchemy engine for database connection
    """
    try:
        # Load environment variables
        load_dotenv()

        # Create connection string for pyodbc
        conn_str = (
            f"DRIVER={os.getenv('DB_DRIVER')};"
            f"SERVER={os.getenv('DB_SERVER')};"
            f"DATABASE={os.getenv('DB_NAME')};"
            f"Trusted_Connection=yes;"
        )

        # Create SQLAlchemy engine using the pyodbc connection
        logging.info(f"Connecting to database {os.getenv('DB_NAME')} on server {os.getenv('DB_SERVER')}")
        connection_url = f"mssql+pyodbc:///?odbc_connect={conn_str.replace(';', '%3B')}"
        engine = create_engine(connection_url)
        logging.info("Database connection successful")
        return engine

    except Exception as e:
        logging.error(f"Database connection failed: {str(e)}")
        raise

def chunk(lst, size=1000):
    """Break a list into chunks of specified size.

    Args:
        lst: The list to chunk
        size: Maximum chunk size (default: 1000)

    Yields:
        Chunks of the list
    """
    for i in range(0, len(lst), size):
        yield lst[i:i + size]

def query_part_data(engine, part_numbers):
    """Query the database for part information.

    Args:
        engine: SQLAlchemy engine for database connection
        part_numbers: List of part numbers to query

    Returns:
        DataFrame containing part information
    """
    if not part_numbers:
        logging.warning("No part numbers provided for query")
        return pd.DataFrame()

    results = []

    try:
        for part_chunk in chunk(part_numbers):
            part_list = ','.join(f"'{p}'" for p in part_chunk)

            logging.info(f"Querying database for {len(part_chunk)} part numbers")

            query = f"""
            WITH latest_so AS (
              SELECT 
                FSONO,
                FPARTNO,
                FPARTREV,
                FPRICE AS SO_PRICE,
                FQUANTITY,
                ROW_NUMBER() OVER (PARTITION BY FPARTNO ORDER BY FSONO DESC) AS rn
              FROM SOITEM
              WHERE FPARTNO IN ({part_list})
            )
            SELECT 
              i.FPARTNO,
              i.FREV,
              i.FPRICE AS BASE_PRICE,
              i.FONHAND,
              i.FONORDER,
              i.FBOOK,
              i.FDISPLCOST,
              i.FDISPMCOST,
              i.FDISPOCOST,
              i.FDESCript AS DESCRIPTION,
              s.FSONO,
              s.FPARTREV AS LAST_ORDER_REV,
              s.SO_PRICE,
              s.FQUANTITY AS LAST_ORDER_QTY
            FROM INMAST i
            LEFT JOIN latest_so s ON i.FPARTNO = s.FPARTNO AND s.rn = 1
            WHERE i.FPARTNO IN ({part_list})
            """

            # Use pandas read_sql with SQLAlchemy connection
            with engine.connect() as connection:
                chunk_df = pd.read_sql(query, connection)
                logging.info(f"Query returned {len(chunk_df)} records")
                results.append(chunk_df)

        if results:
            final_df = pd.concat(results, ignore_index=True)
            return final_df
        else:
            logging.warning("No results returned from database")
            return pd.DataFrame()

    except pyodbc.Error as e:
        logging.error(f"Database query failed: {str(e)}")
        raise
    except Exception as e:
        logging.error(f"Unexpected error during database query: {str(e)}")
        raise
