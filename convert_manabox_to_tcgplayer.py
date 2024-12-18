import csv
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from rapidfuzz import process, fuzz

# Get input file path using a file dialog
Tk().withdraw()  # Hide the root window
manabox_csv = askopenfilename(title="Select the Manabox CSV File", filetypes=[("CSV Files", "*.csv")])
if not manabox_csv:
    print("No file selected. Exiting.")
    exit()

reference_csv = "TCGplayer__Pricing_Custom_Export_20241213_025218.csv"
tcgplayer_csv = "tcgplayer_staged.csv"

def normalize_key(card_name, set_name, condition, number):
    """
    Normalize card name, set name, condition, and card number for consistent key matching.
    Removes any suffixes in parentheses initially, reappending them after matching.
    """
    suffix = ""
    if "(" in card_name and ")" in card_name:
        suffix_match = re.search(r"\((.*?)\)$", card_name)
        if suffix_match:
            suffix = suffix_match.group(0)
        card_name = re.sub(r"\((.*?)\)$", "", card_name).strip()

    card_name = card_name.split('//')[0].strip()  # Use only text before '//'
    normalized_card_name = re.sub(r"[^a-zA-Z0-9 ,'-]", "", card_name).strip().lower()
    if "foil" in condition.lower():
        normalized_card_name += " foil"
    normalized_set_name = re.sub(r"[^a-zA-Z0-9 ]", "", set_name).strip().lower()
    if normalized_set_name in ["plst", "the list"]:
        normalized_set_name = "the list reprints"

    if normalized_set_name == "the list":
        number = number.split("-")[-1]  # Take only the numeric part of the number

    normalized_number = re.sub(r"^[A-Za-z\-]*", "", str(number).strip()) if number else ""
    return (normalized_card_name, normalized_set_name, normalized_number, condition.lower(), suffix)

# Map condition codes to human-readable condition strings
CONDITION_MAP = {
    "near_mint": "Near Mint",
    "lightly_played": "Lightly Played",
    "moderately_played": "Moderately Played",
    "heavily_played": "Heavily Played",
    "damaged": "Damaged"
}

def load_reference_data(reference_csv):
    """
    Load reference data from the CSV file into a dictionary.
    """
    reference_data = {}
    try:
        with open(reference_csv, mode='r', newline='', encoding='utf-8') as ref_file:
            ref_reader = csv.DictReader(ref_file)
            for row in ref_reader:
                key = normalize_key(
                    row.get("Product Name", ""),
                    row.get("Set Name", ""),
                    row.get("Condition", ""),
                    row.get("Number", "")
                )
                reference_data[key[:4]] = row
    except FileNotFoundError:
        print(f"Reference file {reference_csv} not found. Exiting.")
        exit()
    return reference_data

reference_data = load_reference_data(reference_csv)
confirmed_matches = {}

def find_best_match(normalized_key, reference_keys):
    """
    Find the best match for a given key using rapidfuzz matching.
    """
    normalized_key_str = " ".join(normalized_key)
    matches = process.extract(normalized_key_str, [" ".join(k) for k in reference_keys], scorer=fuzz.ratio, limit=10)
    return matches

def map_fields(manabox_row):
    """
    Converts a row from the Manabox CSV format into the TCGPlayer staged inventory format.
    """
    card_name = manabox_row.get("Name", "").strip()
    set_name = manabox_row.get("Set name", "").strip()
    card_number = re.sub(r"^[A-Za-z\-]*", "", manabox_row.get("Collector number", "").strip().split("-")[-1])
    condition_code = manabox_row.get("Condition", "near_mint").strip().lower()
    foil = "Foil" if manabox_row.get("Foil", "normal").lower() == "foil" else ""
    condition = CONDITION_MAP.get(condition_code, "Near Mint")

    if foil:
        condition += " Foil"

    if not card_name or not set_name:
        print(f"Skipping row with missing fields: {manabox_row}")
        return None

    normalized_result = normalize_key(card_name, set_name, condition, card_number)
    key = normalized_result[:4]

    # Check confirmed matches first
    if key in confirmed_matches:
        ref_row = confirmed_matches[key]
    else:
        matches = find_best_match(key, reference_data.keys())
        for match, score, _ in matches:
            if score == 100:
                ref_row_key = next((k for k in reference_data if " ".join(k) == match), None)
                if ref_row_key:
                    ref_row = reference_data[ref_row_key]
                    confirmed_matches[key] = ref_row
                    print(f"Auto-confirmed match with 100% confidence: {key} -> {confirmed_matches[key]}")
                    break
            else:
                print(f"Ambiguous Match Found for: {card_name} - {set_name} - {card_number}")
                print("Possible Match:", match, f"Score: {score}")
                user_input = input("Is this the correct match? (yes/no): ").lower()
                if user_input == 'yes':
                    ref_row_key = next((k for k in reference_data if " ".join(k) == match), None)
                    if ref_row_key:
                        ref_row = reference_data[ref_row_key]
                        confirmed_matches[key] = ref_row
                        print(f"Confirmed and added match: {key} -> {confirmed_matches[key]}")
                        break
                else:
                    print("Skipping this match.")
        else:
            print("No suitable match confirmed. Skipping entry.")
            return None

    if not ref_row:
        print(f"No match found for key: {key}")
        return None

    # Correct the condition if it doesn't match the row data
    if "foil" in ref_row.get("Product Name", "").lower() and "foil" not in condition.lower():
        condition += " Foil"

    tcgplayer_id = ref_row.get("TCGplayer Id", "Not Found")
    return {
        "TCGplayer Id": tcgplayer_id,
        "Product Line": ref_row.get("Product Line", "Magic: The Gathering"),
        "Set Name": ref_row.get("Set Name", ""),
        "Product Name": ref_row.get("Product Name", "") + normalized_result[4],
        "Title": ref_row.get("Title", ""),
        "Number": ref_row.get("Number", ""),
        "Rarity": ref_row.get("Rarity", ""),
        "Condition": condition,
        "TCG Market Price": ref_row.get("TCG Market Price", ""),
        "TCG Direct Low": ref_row.get("TCG Direct Low", ""),
        "TCG Low Price With Shipping": ref_row.get("TCG Low w/ Shipping", ""),
        "TCG Low Price": ref_row.get("TCG Low Price", ""),
        "Total Quantity": ref_row.get("Total Quantity", ""),
        "Add to Quantity": manabox_row.get("Quantity", "1"),
        "TCG Marketplace Price": manabox_row.get("Purchase price", "0.00"),
        "Photo URL": ref_row.get("Photo URL", ""),
        "My Store Reserve Quantity": ref_row.get("My Store Reserve Quantity", ""),
        "My Store Price": ref_row.get("My Store Price", "")
    }

def update_tcgplayer_ids(output_file, reference_data):
    try:
        with open(output_file, mode='r', newline='', encoding='utf-8') as infile:
            rows = list(csv.DictReader(infile))

        with open(output_file, mode='w', newline='', encoding='utf-8') as outfile:
            fieldnames = rows[0].keys()
            writer = csv.DictWriter(outfile, fieldnames=fieldnames)
            writer.writeheader()

            for row in rows:
                key = normalize_key(row["Product Name"], row["Set Name"], row["Condition"], row["Number"])
                if key[:4] in reference_data:
                    row["TCGplayer Id"] = reference_data[key[:4]].get("TCGplayer Id", "Not Found")
                else:
                    print(f"No TCGplayer ID match during update for key: {key[:4]}")
                writer.writerow(row)
    except Exception as e:
        print(f"Error during TCGplayer ID update: {e}")

try:
    with open(manabox_csv, mode='r', newline='', encoding='utf-8') as infile:
        with open(tcgplayer_csv, mode='w', newline='', encoding='utf-8') as outfile:
            reader = csv.DictReader(infile)
            fieldnames = [
                "TCGplayer Id", "Product Line", "Set Name", "Product Name", "Title", "Number", "Rarity",
                "Condition", "TCG Market Price", "TCG Direct Low", "TCG Low Price With Shipping", "TCG Low Price",
                "Total Quantity", "Add to Quantity", "TCG Marketplace Price", "Photo URL",
                "My Store Reserve Quantity", "My Store Price"
            ]
            writer = csv.DictWriter(outfile, fieldnames=fieldnames)
            writer.writeheader()

            for row in reader:
                tcgplayer_row = map_fields(row)
                if tcgplayer_row:
                    writer.writerow(tcgplayer_row)

    update_tcgplayer_ids(tcgplayer_csv, reference_data)
    print(f"Conversion complete. Output saved to {tcgplayer_csv}")

except FileNotFoundError as e:
    print(f"Error: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
