import os
import sys
import pandas as pd
import logging
from io import StringIO
from excel_loader import load_excel_file
import json

# --- Configuration Constants ---
EXCEL_DATA_FILE = 'words.xlsx' # Your Excel file name (initial load only)
DATA_FILE = 'flashcards_state.json' # Where your program state (cards + session) will be saved
LOG_FILE = 'lang.log' # Log file name

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
    def __init__(self, english, chinese, pinyin, state='New', last_reviewed_session=0, original_index=None):
        self.english = english
        self.chinese = chinese
        self.pinyin = pinyin
        self.state = state
        self.last_reviewed_session = last_reviewed_session
        self.original_index = original_index # To track the row in the DataFrame (relevant during initial load)

    def __repr__(self):
        return (f"Flashcard(English='{self.english}', Chinese='{self.chinese}', Pinyin='{self.pinyin}', "
                f"State='{self.state}', LastSession={self.last_reviewed_session}, OriginalIndex={self.original_index})")

    def to_dict(self):
        """Converts the Flashcard object to a dictionary for JSON serialization."""
        return {
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
        return Flashcard(
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
        return Flashcard(
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

def load_excel_initial_data(file_name):
    """
    Loads initial vocabulary from an XLSX Excel file into a pandas DataFrame.
    This is only used if the main JSON state file doesn't exist yet.
    """
    try:
        df = pd.read_excel(file_name)
        logger.info(f"Successfully loaded initial vocabulary from '{file_name}'.")

        # Ensure 'State' and 'LastReviewedSession' columns exist for new cards
        if 'State' not in df.columns:
            df['State'] = 'New'
        if 'LastReviewedSession' not in df.columns:
            df['LastReviewedSession'] = 0

        return df

    except FileNotFoundError:
        logger.error(f"Initial Excel file '{file_name}' not found. Please create it or add cards manually.", exc_info=False)
        return pd.DataFrame(columns=['English', 'Chinese', 'Pinyin', 'State', 'LastReviewedSession']) # Return empty DataFrame
    except Exception as e:
        logger.error(f"An error occurred while loading initial Excel file: {e}", exc_info=True)
        return pd.DataFrame(columns=['English', 'Chinese', 'Pinyin', 'State', 'LastReviewedSession']) # Return empty DataFrame


def load_program_state():
    """
    Loads program state (flashcards and current_session) from the JSON DATA_FILE.
    If the JSON file doesn't exist, it attempts to load initial vocabulary from Excel.
    """
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                cards = [Flashcard.from_dict(d) for d in data.get('cards', [])]
                current_session = data.get('current_session', 0)
                logger.info(f"Loaded program state from '{DATA_FILE}'. {len(cards)} cards, Session: {current_session}")
                return cards, current_session
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logger.error(f"Error loading JSON data from '{DATA_FILE}': {e}. Attempting initial load from Excel.", exc_info=True)
            pass # Fall through to Excel load if JSON fails

    # If JSON doesn't exist or failed to load, try to load from Excel for initial setup
    logger.info(f"'{DATA_FILE}' not found or corrupted. Attempting initial load from '{EXCEL_DATA_FILE}'.")
    df_initial = load_excel_file(EXCEL_DATA_FILE)
    cards = []
    for index, row in df_initial.iterrows():
        card = Flashcard.from_dataframe_row(row, index)
        cards.append(card)
    current_session = 0 # Start session at 0 for a new setup

    logger.info(f"Initial load from Excel completed. {len(cards)} cards created.")
    return cards, current_session

def save_program_state(cards, current_session):
    """Saves flashcards and current session number to the JSON DATA_FILE."""
    data = {
        'cards': [card.to_dict() for card in cards],
        'current_session': current_session
    }
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        logger.info(f"Program state saved successfully to '{DATA_FILE}'.")
    except IOError as e:
        logger.error(f"Error saving program state to '{DATA_FILE}': {e}", exc_info=True)


# --- Main Program Logic ---
# The add_new_card function is now removed as it's no longer a menu option.
# If you wanted to re-introduce adding cards, it would be by editing flashcards.xlsx

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
                logger.info(f"Card '{card.english}' correct. New state: {card.state}")
                break
            elif response == 'n':
                card.regress_state()
                print(f"Incorrect. Card moved to state: {card.state}")
                logger.info(f"Card '{card.english}' incorrect. New state: {card.state}")
                break
            elif response == 'skip':
                print("Skipping this card for now.")
                logger.info(f"Card '{card.english}' skipped.")
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
        print("No flashcards added yet.")
        logger.info("No flashcards to list.")
        return

    sorted_cards = sorted(cards, key=lambda c: c.english.lower())

    for card in sorted_cards:
        print(f"'{card.english}' -> '{card.chinese}' ({card.pinyin}) | State: {card.state} | Last Reviewed Session: {card.last_reviewed_session} | Next Review: Session {card.get_next_review_session()}")
    print("----------------------")
    logger.info(f"Listed {len(cards)} flashcards.")

def main():
    cards, current_session = load_program_state()

    while True:
        print("\n--- Flashcard Program Menu ---")
        print(f"Current Session Number: {current_session}")
        print("1. Start Review Session")
        # Option 2 (Add New Flashcard) removed
        print("2. List All Cards") # Now option 2
        print("3. Increment Session Number (without review)") # Now option 3
        print("4. Exit") # Now option 4
        choice = input("Enter your choice: ").strip()

        if choice == '1':
            current_session += 1
            review_session(cards, current_session)
            save_program_state(cards, current_session)
        elif choice == '2': # This was '3' before
            list_all_cards(cards)
        elif choice == '3': # This was '4' before
            current_session += 1
            print(f"Session number incremented to {current_session}.")
            save_program_state(cards, current_session)
        elif choice == '4': # This was '5' before
            save_program_state(cards, current_session)
            print("Exiting Flashcard Program. Goodbye!")
            logger.info("Flashcard program exited.")
            break
        else:
            print("Invalid choice. Please try again.")
            logger.warning(f"Invalid menu choice: {choice}")

if __name__ == "__main__":
    main()