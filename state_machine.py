import json
import os
from datetime import datetime

''' to run
python3 state_machine.py


approximate logic
State: New (Just added)
  -> On correct answer: Learning1 (See again next session)
  -> On incorrect answer: New (Stay here)

State: Learning1 (Seen once, got right)
  -> On correct answer: Learning2 (See again in three sessions)
  -> On incorrect answer: New (Go back to start)

State: Learning2 (Seen twice, got right)
  -> On correct answer: Known (See again in 5 sessions)
  -> On incorrect answer: Learning1 (Go back one step)

State: Known (Mastered)
  -> On correct answer: Known (see again in 10 sessions)
  -> On incorrect answer: Learning2 (Go back to intermediate step)
  '''




# --- Configuration Constants ---
DATA_FILE = 'flashcards_data.json' # Where your flashcard data will be saved
SESSION_INTERVALS = {
    'New': 0,           # Reviewed in the current session if still 'New' after first wrong
    'Learning1': 1,     # Reviewed 1 session after getting it right from 'New'
    'Learning2': 3,     # Reviewed 3 sessions after getting it right from 'Learning1'
    'Known': 5,         # Reviewed 5 sessions after getting it right from 'Learning2'
    'Mastered': 10      # Reviewed 10 sessions after getting it right from 'Known'
}
# Define the order of states for easy progression and regression
STATE_ORDER = ['New', 'Learning1', 'Learning2', 'Known', 'Mastered']

# --- Flashcard Class ---
class Flashcard:
    def __init__(self, english, chinese, pinyin, state='New', last_reviewed_session=0):
        self.english = english
        self.chinese = chinese
        self.pinyin = pinyin
        self.state = state
        self.last_reviewed_session = last_reviewed_session # The session number when this card was last reviewed successfully

    def __repr__(self):
        return f"Flashcard(English='{self.english}', Chinese='{self.chinese}', Pinyin='{self.pinyin}', State='{self.state}', LastSession={self.last_reviewed_session})"

    def to_dict(self):
        """Converts the Flashcard object to a dictionary for JSON serialization."""
        return {
            'english': self.english,
            'chinese': self.chinese,
            'pinyin': self.pinyin,
            'state': self.state,
            'last_reviewed_session': self.last_reviewed_session
        }

    @staticmethod
    def from_dict(data):
        """Creates a Flashcard object from a dictionary."""
        return Flashcard(
            english=data['english'],
            chinese=data['chinese'],
            pinyin=data['pinyin'],
            state=data['state'],
            last_reviewed_session=data['last_reviewed_session']
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
            # Update last_reviewed_session only if successfully moved forward
            self.last_reviewed_session = current_session_num
        except ValueError:
            print(f"Warning: Unknown state '{self.state}' for card: {self.english}")

    def regress_state(self):
        """Moves the card back one state upon incorrect answer."""
        try:
            current_idx = STATE_ORDER.index(self.state)
            if current_idx > 0:
                self.state = STATE_ORDER[current_idx - 1]
            else:
                # If already in 'New', it stays in 'New'
                self.state = 'New'
        except ValueError:
            print(f"Warning: Unknown state '{self.state}' for card: {self.english}")

# --- Data Management Functions ---
def load_cards():
    """Loads flashcards from the data file."""
    if not os.path.exists(DATA_FILE):
        return [], 0 # No cards, current session number is 0

    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            cards = [Flashcard.from_dict(d) for d in data.get('cards', [])]
            current_session = data.get('current_session', 0)
            print(f"Loaded {len(cards)} cards. Current session number: {current_session}")
            return cards, current_session
    except (json.JSONDecodeError, FileNotFoundError) as e:
        print(f"Error loading data: {e}. Starting with an empty set of cards.")
        return [], 0

def save_cards(cards, current_session):
    """Saves flashcards and current session number to the data file."""
    data = {
        'cards': [card.to_dict() for card in cards],
        'current_session': current_session
    }
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print("Cards saved successfully.")

# --- Main Program Logic ---
def add_new_card(cards):
    """Prompts the user to add a new flashcard."""
    print("\n--- Add New Flashcard ---")
    english = input("Enter English word/phrase: ").strip()
    if not english:
        print("English word cannot be empty. Aborting.")
        return

    chinese = input("Enter Chinese character(s): ").strip()
    pinyin = input("Enter Pinyin (with tone marks if possible): ").strip()

    new_card = Flashcard(english, chinese, pinyin)
    cards.append(new_card)
    print(f"Added new card: '{new_card.english}'")

def review_session(cards, current_session_num):
    """Conducts a review session based on card states."""
    print(f"\n--- Starting Review Session {current_session_num} ---")

    # Filter cards that are due for review in this session
    cards_to_review = [
        card for card in cards
        if card.get_next_review_session() <= current_session_num
    ]

    if not cards_to_review:
        print("No cards due for review this session. Good job!")
        return

    # Sort cards: New cards first, then by how long ago they were reviewed (older first)
    # This ensures new cards are seen quickly and forgotten cards appear sooner.
    cards_to_review.sort(key=lambda card: (STATE_ORDER.index(card.state), card.last_reviewed_session))

    print(f"You have {len(cards_to_review)} cards to review.")

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
                break
            elif response == 'n':
                card.regress_state()
                print(f"Incorrect. Card moved to state: {card.state}")
                break
            elif response == 'skip':
                print("Skipping this card for now.")
                break # Just break without changing state or last_reviewed_session
            else:
                print("Invalid input. Please enter 'y' for yes, 'n' for no, or 'skip'.")

    print("\n--- Review Session Ended ---")

def list_all_cards(cards):
    """Prints a list of all flashcards and their current states."""
    print("\n--- All Flashcards ---")
    if not cards:
        print("No flashcards added yet.")
        return

    # Sort cards for consistent display (e.g., by English word)
    sorted_cards = sorted(cards, key=lambda c: c.english.lower())

    for card in sorted_cards:
        print(f"'{card.english}' -> '{card.chinese}' ({card.pinyin}) | State: {card.state} | Last Reviewed Session: {card.last_reviewed_session} | Next Review: Session {card.get_next_review_session()}")
    print("----------------------")

def main():
    cards, current_session = load_cards()

    while True:
        print("\n--- Flashcard Program Menu ---")
        print(f"Current Session Number: {current_session}")
        print("1. Start Review Session")
        print("2. Add New Flashcard")
        print("3. List All Cards")
        print("4. Increment Session Number (without review)")
        print("5. Exit")
        choice = input("Enter your choice: ").strip()

        if choice == '1':
            current_session += 1 # Increment session for the review
            review_session(cards, current_session)
            save_cards(cards, current_session) # Save after each session
        elif choice == '2':
            add_new_card(cards)
            save_cards(cards, current_session) # Save after adding new card
        elif choice == '3':
            list_all_cards(cards)
        elif choice == '4':
            current_session += 1
            print(f"Session number incremented to {current_session}.")
            save_cards(cards, current_session)
        elif choice == '5':
            save_cards(cards, current_session)
            print("Exiting Flashcard Program. Goodbye!")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()