import os
import sys
import pandas as pd
import logging
from io import StringIO # Import StringIO for df.info() capture

# --- Configure logging ---
# Create a logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO) # Set the minimum level of messages to log

# Create a file handler
log_file = 'lang.log'
file_handler = logging.FileHandler(log_file)
file_handler.setLevel(logging.INFO) # Log INFO messages and above to the file

# Create a formatter and set it for the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# Add the file handler to the logger
logger.addHandler(file_handler)

# IMPORTANT: The console handler is removed from here to stop console logging
# console_handler = logging.StreamHandler(sys.stdout)
# console_handler.setLevel(logging.INFO)
# console_handler.setFormatter(formatter)
# logger.addHandler(console_handler)


def load_excel_file(file_name):
    """
    Loads an XLSX Excel file into a pandas DataFrame.

    Args:
        file_name (str): The path to the XLSX file.

    Returns:
        pandas.DataFrame: The loaded DataFrame, or None if an error occurs.
    """
    try:
        df = pd.read_excel(file_name)
        logger.info(f"Successfully loaded '{file_name}' into a DataFrame.")
        logger.info("--- First 5 rows of the DataFrame ---")
        logger.info("\n" + df.head().to_markdown(index=False, numalign="left", stralign="left")) # Log DataFrame head

        logger.info("--- DataFrame Information (Columns and Data Types) ---")
        # Use a StringIO object to capture df.info() output
        buffer = StringIO()
        df.info(buf=buffer)
        logger.info(buffer.getvalue())

        return df

    except FileNotFoundError:
        logger.error(f"Error: The file '{file_name}' was not found. Please ensure the file is in the correct directory.")
        return None
    except Exception as e:
        logger.error(f"An error occurred while loading the Excel file: {e}", exc_info=True) # exc_info=True to log traceback
        return None

