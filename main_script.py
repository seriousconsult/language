import os
import sys
import pandas as pd
import logging
from io import StringIO
import json
# Assuming excel_loader.py has the load_excel_file function
# Assuming export_to_pptx.py has the export_dataframe_to_pptx function
from excel_loader import load_excel_file # This is your load_excel_file with the df.head(3) change
from export_to_pptx import export_dataframe_to_pptx

# --- Configuration Constants ---
EXCEL_DATA_FILE = 'words.xlsx' # Your Excel file name (main vocabulary source for both flashcards and PPTX)
JSON_DATA_FILE = 'flashcards_state.json' # Where your program state (card states + session) will be saved
LOG_FILE = 'lang.log' # Log file name
PPTX_OUTPUT_FILE = "words_data.pptx" # Define the output PowerPoint file name

SESSION_INTERVALS = {
    'New': 0,
    'Learning1': 1,
    'Learning2': 3,
    'Known': 5,
    'Mastered': 10
}
STATE_ORDER = ['New', 'Learning1', 'Learning2', 'Known', 'Mastered']

# --- Configure logging ---
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler(LOG_FILE)
file_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# --- Flashcard Class ---
class Flashcard:
    # --- MODIFIED: Added 'id' to __init__ ---
    def __init__(self, english, chinese, pinyin, id=None, state='New', last_reviewed_session=0, original_index=None):
        self.id = id # Store the ID
        self.english = english
        self.chinese = chinese
        self.pinyin = pinyin
        self.state = state
        self.last_reviewed_session = last_reviewed_session
        self.original_index = original_index # To track the row in the DataFrame (relevant during initial load)

    def __repr__(self):
        # --- MODIFIED: Added 'id' to __repr__ (optional, but good for debugging) ---
        return (f"Flashcard(ID='{self.id}', English='{self.english}', Chinese='{self.chinese}', Pinyin='{self.pinyin}', "
                f"State='{self.state}', LastSession={self.last_reviewed_session}, OriginalIndex={self.original_index})")

    def to_dict(self):
        """Converts the Flashcard object to a dictionary for JSON serialization."""
        # --- MODIFIED: Added 'id' to to_dict ---
        return {
            'id': self.id,
            'english': self.english,
            'chinese': self.chinese,
            'pinyin': self.pinyin,
            'state': self.state,
            'last_reviewed_session': self.last_reviewed_session
            # original_index is not saved as it's only relevant during initial Excel import
        }

    @staticmethod
    def from_dict(data):
        """Creates a Flashcard object from a dictionary (for loading from JSON)."""
        # --- MODIFIED: Added 'id' to from_dict ---
        return Flashcard(
            id=data.get('id'), # Use .get() for robustness, in case old JSON lacks 'id'
            english=data['english'],
            chinese=data['chinese'],
            pinyin=data['pinyin'],
            state=data['state'],
            last_reviewed_session=data['last_reviewed_session'],
            original_index=None # Not relevant when loading from JSON state
        )

    @staticmethod
    def from_dataframe_row(row, index):
        """Creates a Flashcard object from a pandas DataFrame row (for initial Excel load)."""
        # When loading from Excel to initialize a new card, state and last_reviewed_session
        # are defaulted if not present in the Excel file itself.
        # This is primarily for the *first* run or new cards added to Excel.
        return Flashcard(
            id=row.get('number',''), # You already correctly added this.
            english=row.get('English', ''),
            chinese=row.get('Chinese', ''),
            pinyin=row.get('Pinyin', ''),
            state=row.get('State', 'New'),
            last_reviewed_session=row.get('LastReviewedSession', 0),
            original_index=index
        )

    def get_next_review_session(self):
        """Calculates the session number when this card should be reviewed next."""
        interval = SESSION_INTERVALS.get(self.state, 0)
        return self.last_reviewed_session + interval

    def advance_state(self, current_session_num):
        """Moves the card to the next state upon correct answer."""
        try:
            current_idx = STATE_ORDER.index(self.state)
            if current_idx < len(STATE_ORDER) - 1:
                self.state = STATE_ORDER[current_idx + 1]
            else:
                self.state = STATE_ORDER[-1] # Stays in the highest state
            self.last_reviewed_session = current_session_num
        except ValueError:
            logger.warning(f"Unknown state '{self.state}' for card: {self.english}. Resetting to 'New'.")
            self.state = 'New'
            self.last_reviewed_session = current_session_num

    def regress_state(self):
        """Moves the card back one state upon incorrect answer."""
        try:
            current_idx = STATE_ORDER.index(self.state)
            if current_idx > 0:
                self.state = STATE_ORDER[current_idx - 1]
            else:
                self.state = 'New' # If already in 'New', it stays in 'New'
        except ValueError:
            logger.warning(f"Unknown state '{self.state}' for card: {self.english}. Resetting to 'New'.")
            self.state = 'New'


# --- Data Management Functions ---

# NOTE: The load_excel_file function is assumed to be in excel_loader.py
# and modified there to log only the first 3 rows as discussed previously.
# I'm not including it here to avoid duplication, assuming it's in its own file.

def load_program_state():
    """
    Loads flashcard content from Excel and then merges with saved state from JSON.
    This function manages the reconciliation of vocabulary from Excel and learning state from JSON.
    """
    cards_df_from_excel = load_excel_file(EXCEL_DATA_FILE)
    
    # Initialize default state and session number
    saved_cards_state = {}
    current_session = 0

    # Try to load saved state from JSON
    if os.path.exists(JSON_DATA_FILE):
        try:
            with open(JSON_DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
                # Load individual card states from JSON into a lookup dictionary
                for card_data in data.get('cards', []):
                    # Use a tuple of content as a key for reconciliation
                    # --- MODIFIED: Ensuring 'id' is part of the key and it's present ---
                    key = (card_data.get('id', ''), card_data.get('english', ''), card_data.get('chinese', ''), card_data.get('pinyin', ''))
                    if all(val is not None and val != '' for val in key): # Ensure all parts of the key are non-empty/non-None
                        saved_cards_state[key] = {
                            'state': card_data.get('state', 'New'),
                            'last_reviewed_session': card_data.get('last_reviewed_session', 0)
                        }
                current_session = data.get('current_session', 0)
                logger.info(f"Loaded {len(saved_cards_state)} card states and session {current_session} from '{JSON_DATA_FILE}'.")
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logger.error(f"Error loading JSON data from '{JSON_DATA_FILE}': {e}. Starting with default state and session 0.", exc_info=True)
            # Reset if JSON is corrupt or missing
            saved_cards_state = {}
            current_session = 0
    else:
        logger.info(f"'{JSON_DATA_FILE}' not found. Starting with all cards 'New' and session 0.")

    # Reconcile cards from Excel with saved states
    active_cards = []
    if cards_df_from_excel is not None: # Ensure DataFrame was loaded successfully
        for index, row in cards_df_from_excel.iterrows():
            excel_id = row.get('number','') # Get ID from 'number' column
            excel_english = row.get('English', '')
            excel_chinese = row.get('Chinese', '')
            excel_pinyin = row.get('Pinyin', '')
            
            card_key = (excel_id, excel_english, excel_chinese, excel_pinyin)
            
            initial_state = 'New'
            initial_last_reviewed_session = 0

            # If a matching card state was found in JSON, use it
            if card_key in saved_cards_state:
                state_info = saved_cards_state[card_key]
                initial_state = state_info['state']
                initial_last_reviewed_session = state_info['last_reviewed_session']
                logger.debug(f"Matched Excel card ID: {excel_id}, English: '{excel_english}' with saved state: {initial_state}")
            else:
                logger.debug(f"Excel card ID: {excel_id}, English: '{excel_english}' not found in saved state. Initializing as New.")

            # Create Flashcard object with combined data
            card = Flashcard(
                id=excel_id, # Pass the ID to the Flashcard constructor
                english=excel_english,
                chinese=excel_chinese,
                pinyin=excel_pinyin,
                state=initial_state,
                last_reviewed_session=initial_last_reviewed_session
            )
            active_cards.append(card)
    else:
        logger.error("No Excel data available to load flashcards.")

    logger.info(f"Final active cards in memory: {len(active_cards)}. Current session: {current_session}")
    return active_cards, current_session


def save_program_state(cards, current_session):
    """Saves flashcards' current state and session number to the JSON_DATA_FILE."""
    data = {
        'cards': [card.to_dict() for card in cards],
        'current_session': current_session
    }
    try:
        with open(JSON_DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        logger.info(f"Program state saved successfully to '{JSON_DATA_FILE}'.")
    except IOError as e:
        logger.error(f"Error saving program state to '{JSON_DATA_FILE}': {e}", exc_info=True)


# --- Core Program Logic (Flashcards) ---

def review_session(cards, current_session_num):
    """Conducts a review session based on card states."""
    logger.info(f"\n--- Starting Review Session {current_session_num} ---")
    print(f"\n--- Starting Review Session {current_session_num} ---")

    cards_to_review = [
        card for card in cards
        if card.get_next_review_session() <= current_session_num
    ]

    if not cards_to_review:
        print("No cards due for review this session. Good job!")
        logger.info("No cards due for review this session.")
        return

    cards_to_review.sort(key=lambda card: (STATE_ORDER.index(card.state), card.last_reviewed_session))

    print(f"You have {len(cards_to_review)} cards to review.")
    logger.info(f"Reviewing {len(cards_to_review)} cards.")

    for i, card in enumerate(cards_to_review):
        print(f"\n--- Card {i + 1}/{len(cards_to_review)} ---")
        print(f"ID: {card.id}") # Optional: Display ID
        print(f"Current State: {card.state}")
        print(f"English: {card.english}")
        input("Press Enter to reveal Chinese/Pinyin...")
        print(f"Chinese: {card.chinese}")
        print(f"Pinyin: {card.pinyin}")

        while True:
            response = input("Did you get it right? (y/n/skip): ").strip().lower()
            if response == 'y':
                card.advance_state(current_session_num)
                print(f"Correct! Card moved to state: {card.state}")
                logger.info(f"Card '{card.english}' (ID: {card.id}) correct. New state: {card.state}")
                break
            elif response == 'n':
                card.regress_state()
                print(f"Incorrect. Card moved to state: {card.state}")
                logger.info(f"Card '{card.english}' (ID: {card.id}) incorrect. New state: {card.state}")
                break
            elif response == 'skip':
                print("Skipping this card for now.")
                logger.info(f"Card '{card.english}' (ID: {card.id}) skipped.")
                break
            else:
                print("Invalid input. Please enter 'y' for yes, 'n' for no, or 'skip'.")

    print("\n--- Review Session Ended ---")
    logger.info("Review session ended.")

def list_all_cards(cards):
    """Prints a list of all flashcards and their current states."""
    logger.info("\n--- Listing All Flashcards ---")
    print("\n--- All Flashcards ---")
    if not cards:
        print("No flashcards found in 'words.xlsx'. Please add data there.")
        logger.info("No flashcards to list.")
        return

    sorted_cards = sorted(cards, key=lambda c: c.english.lower()) # Still sort by English for consistency

    for card in sorted_cards:
        print(f"ID: {card.id} | '{card.english}' -> '{card.chinese}' ({card.pinyin}) | State: {card.state} | Last Reviewed Session: {card.last_reviewed_session} | Next Review: Session {card.get_next_review_session()}")
    print("----------------------")
    logger.info(f"Listed {len(cards)} flashcards.")

# This function now correctly takes a DataFrame for PPTX export
def export_to_pptx_menu_option(dataframe_to_export, rows_per_slide_val):
    """Handles the PPTX export option from the menu."""
    if dataframe_to_export is not None and not dataframe_to_export.empty:
        logger.info(f"Attempting to export data to PowerPoint: {PPTX_OUTPUT_FILE}")
        export_dataframe_to_pptx(dataframe_to_export, PPTX_OUTPUT_FILE, slide_title="Words Data from Excel", rows_per_slide=rows_per_slide_val)
        print(f"Exported to Pptx file: {PPTX_OUTPUT_FILE}.")
        logger.info(f"Exported to Pptx file: {PPTX_OUTPUT_FILE}.")
    else:
        print("No data available to export to PowerPoint or Excel file was empty.")
        logger.warning("No data available to export to PowerPoint or Excel file was empty. PPTX export skipped.")

# --- Main Program Execution ---

def main():
    # Load flashcard states and session from JSON, reconciling with Excel content
    cards, current_session = load_program_state()

    # --- PowerPoint Export Configuration (happens once at start, but parameters set here) ---
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
    
    # Load the Excel data *once* for potential PPTX export
    # This DataFrame will be passed to the export_to_pptx_menu_option function when needed.
    words_dataframe_for_pptx = load_excel_file(EXCEL_DATA_FILE)
    if words_dataframe_for_pptx is None or words_dataframe_for_pptx.empty:
        logger.warning("Initial Excel load for PPTX export resulted in no data.")


    # --- Flashcard Program Menu Loop ---
    while True:
        print("\n--- Language Learning Program Menu ---")
        print(f"Current Flashcard Session Number: {current_session}")
        print("1. Start Flashcard Review Session")
        print("2. List All Flashcards")
        print("3. Increment Flashcard Session Number (without review)")
        print("4. Export to Pptx file.")
        print("5. Exit")
        choice = input("Enter your choice: ").strip()

        if choice == '1':
            current_session += 1
            review_session(cards, current_session)
            save_program_state(cards, current_session) # Save state after review
        elif choice == '2':
            list_all_cards(cards)
        elif choice == '3':
            current_session += 1
            print(f"Flashcard session number incremented to {current_session}.")
            save_program_state(cards, current_session) # Save state after increment
        elif choice == '4':
            # Call the new wrapper function, passing the DataFrame loaded earlier
            export_to_pptx_menu_option(words_dataframe_for_pptx, rows_per_slide)
            # Removed 'break' here, as the user might want to do other things after export
            # and it should not exit the program just because they exported a PPTX.
        elif choice == '5':
            save_program_state(cards, current_session) # Final save before exit
            print("Exiting Language Learning Program. Goodbye!")
            logger.info("Language learning program exited.")
            break
        else:
            print("Invalid choice. Please try again.") # Corrected typo "Language learning program exited."
            logger.warning(f"Invalid menu choice: {choice}")

if __name__ == "__main__":
    main()