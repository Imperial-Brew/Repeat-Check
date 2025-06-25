import os
import argparse
import logging
import sys
from src.logger import setup_logger
from src.database import connect_to_database, query_part_data
from src.file_handler import load_part_numbers, save_results

def parse_arguments():
    """Parse command line arguments.

    Returns:
        Parsed arguments
    """
    parser = argparse.ArgumentParser(description='Check part numbers against M2M database')
    parser.add_argument('--input', '-i', default='data\\quote_items_7900_7950_complete.csv',
                        help='Input CSV file (default: data\\quote_items_7900_7950_complete.csv)')
    parser.add_argument('--output', '-o', default='output\\matched_parts_output.csv',
                        help='Output CSV file (default: output\\matched_parts_output.csv)')
    parser.add_argument('--column', '-c', default='part_number',
                        help='Column name containing part numbers (default: part_number)')
    parser.add_argument('--log-level', '-l', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                        default='INFO', help='Logging level (default: INFO)')
    parser.add_argument('--no-log-file', action='store_true',
                        help='Disable logging to file')

    return parser.parse_args()

def main():
    """Main function to orchestrate the process."""
    # Parse command line arguments
    args = parse_arguments()

    # Set up logging
    log_level = getattr(logging, args.log_level)
    logger = setup_logger(log_level=log_level, log_to_file=not args.no_log_file)

    # Initialize engine variable for proper cleanup
    engine = None

    try:
        # Load part numbers from CSV
        logger.info(f"Starting part number check process")
        part_numbers = load_part_numbers(args.input, part_number_column=args.column)

        # Connect to database
        engine = connect_to_database()

        # Query part data
        results_df = query_part_data(engine, part_numbers)

        # Save results
        output_path = save_results(results_df, args.output)

        logger.info(f"✅ Process completed successfully")
        print(f"\n✅ Done! Output saved to '{output_path}'")

        return 0

    except FileNotFoundError as e:
        logger.error(str(e))
        print(f"\nError: {str(e)}")
        return 1
    except ValueError as e:
        logger.error(str(e))
        print(f"\nError: {str(e)}")
        return 1
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        print(f"\nError: An unexpected error occurred. See log for details.")
        return 1
    finally:
        # Close database connection if it was opened
        if engine:
            try:
                engine.dispose()
                logger.info("Database connection closed")
            except Exception as e:
                logger.warning(f"Error closing database connection: {str(e)}")

if __name__ == "__main__":
    sys.exit(main())
