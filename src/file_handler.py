import os
import pandas as pd
import logging

def load_part_numbers(csv_path, part_number_column='part_number'):
    """Load part numbers from CSV file.
    
    Args:
        csv_path: Path to the CSV file
        part_number_column: Column name containing part numbers (default: 'part_number')
        
    Returns:
        List of unique part numbers
    """
    try:
        logging.info(f"Loading data from {csv_path}")
        
        if not os.path.exists(csv_path):
            logging.error(f"CSV file not found: {csv_path}")
            raise FileNotFoundError(f"CSV file not found: {csv_path}")
            
        df = pd.read_csv(csv_path)
        
        if part_number_column not in df.columns:
            available_columns = ', '.join(df.columns)
            logging.error(f"Column '{part_number_column}' not found in CSV. Available columns: {available_columns}")
            raise ValueError(f"Column '{part_number_column}' not found in CSV")
            
        # Extract unique part numbers
        part_numbers = df[part_number_column].dropna().unique().tolist()
        logging.info(f"Loaded {len(df)} rows, found {len(part_numbers)} unique part numbers")
        
        return part_numbers
        
    except pd.errors.EmptyDataError:
        logging.error(f"CSV file is empty: {csv_path}")
        raise
    except pd.errors.ParserError:
        logging.error(f"Error parsing CSV file: {csv_path}")
        raise
    except Exception as e:
        logging.error(f"Unexpected error loading CSV: {str(e)}")
        raise

def save_results(df, output_path):
    """Save results to CSV file.
    
    Args:
        df: DataFrame containing results
        output_path: Path to save the CSV file
        
    Returns:
        Path to the saved file
    """
    try:
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        logging.info(f"Saving {len(df)} records to {output_path}")
        df.to_csv(output_path, index=False)
        logging.info(f"Results successfully saved to {output_path}")
        
        return output_path
        
    except PermissionError:
        logging.error(f"Permission denied when writing to {output_path}")
        raise
    except Exception as e:
        logging.error(f"Failed to save results: {str(e)}")
        raise