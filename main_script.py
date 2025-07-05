import os
import sys # Import sys for command line arguments
from excel_loader import load_excel_file, logger
from export_to_pptx import export_dataframe_to_pptx

if __name__ == "__main__":
    excel_file_name = "/mnt/c/code/language/words.xlsx"
    pptx_output_file = "words_data.pptx" # Define the output PowerPoint file name
    rows_per_slide = 5 # Default value

    # Check for command-line argument for rows_per_slide
    if len(sys.argv) > 1:
        try:
            requested_rows = int(sys.argv[1])
            if 1 <= requested_rows <= 5:
                rows_per_slide = requested_rows
                logger.info(f"Using {rows_per_slide} words per slide as specified from command line.")
            else:
                logger.warning("Invalid number of words per slide. Please provide a number between 1 and 5. Using default of 5.")
        except ValueError:
            logger.warning("Invalid input for words per slide. Please provide an integer. Using default of 5.")
    else:
        logger.info(f"No words per slide specified. Using default of {rows_per_slide}.")


    logger.info(f"Attempting to load Excel file: {excel_file_name}")
    words_dataframe = load_excel_file(excel_file_name)

    if words_dataframe is not None:
        logger.info("Excel file loaded successfully.")
        logger.info(f"Attempting to export data to PowerPoint: {pptx_output_file}")
        export_dataframe_to_pptx(words_dataframe, pptx_output_file, slide_title="Words Data from Excel", rows_per_slide=rows_per_slide)
    else:
        logger.warning("Failed to load Excel file. PowerPoint export skipped.")

    # --- Clean up dummy file (remains the same) ---
    if os.path.exists(excel_file_name) and excel_file_name == '/mnt/c/code/language/words.xlsx':
        # os.remove(excel_file_name) # Uncomment if you want to auto-delete the dummy
        pass
