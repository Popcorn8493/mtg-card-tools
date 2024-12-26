import csv
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import pandas as pd
from rapidfuzz import fuzz

# Option to filter out prerelease cards
FILTER_PRERELEASE = True  # Set to True to exclude prerelease cards
PRERELEASE_PENALTY = -20  # Penalty to prerelease cards
EXACT_MATCH_BOOST = 200  # Boost for exact matches

# Alias mappings for sets
SET_ALIAS = {
    "Universes Beyond: The Lord of the Rings: Tales of Middle-earth": "LTR",
    "Commander: The Lord of the Rings: Tales of Middle-earth": "LTC",
    "the list": "The List"
}

CONDITION_MAP = {
    "near_mint": "Near Mint",
    "lightly_played": "Lightly Played",
    "moderately_played": "Moderately Played",
    "heavily_played": "Heavily Played",
    "damaged": "Damaged"
}

# Keep track of cards the user has given up on
given_up_cards = []
# Record user-confirmed matches for learning purposes
confirmed_matches = {}


def normalize_key(card_name, set_name, condition, number):
    """
    Normalize card name, set name, condition, and card number for consistent key matching.
    Removes any suffixes in parentheses initially, reappending them after matching.
    """
    suffix = ""
    if "(" in card_name and ")" in card_name:
        card_name = re.sub(r"\(.*?\)", "", card_name).strip()  # Remove content within parentheses

    card_name = card_name.split('//')[0].strip()  # Use only text before '//'
    normalized_card_name = re.sub(r"[^a-zA-Z0-9 ,'-]", "", card_name).strip().lower()
    if "foil" in condition.lower():
        normalized_card_name += " foil"
    normalized_set_name = re.sub(r"[^a-zA-Z0-9 ]", "", set_name).strip().lower()
    if normalized_set_name in ["plst", "the list"]:
        normalized_set_name = "the list reprints"

    # Exclude prerelease sets explicitly during normalization
    if "prerelease cards" in normalized_set_name:
        return None

    if normalized_set_name == "the list":
        number = number.split("-")[-1]  # Take only the numeric part of the number

    normalized_number = re.sub(r"^[A-Za-z\-]*", "", str(number).strip()) if number else ""
    return normalized_card_name, normalized_set_name, normalized_number, condition.lower(), suffix


def load_reference_data(reference_csv):
    """
    Load and clean reference data for card matching.
    """
    try:
        # Load the reference data with explicit dtype for "Number"
        ref_data = pd.read_csv(reference_csv, dtype={"Number": "str"})

        # Filter out rows without a valid "Number" or "Set Name"
        ref_data = ref_data[ref_data["Number"].notnull() & ref_data["Set Name"].notnull()]

        if FILTER_PRERELEASE:
            # Identify prerelease cards and log excluded rows
            prerelease_mask = ref_data["Product Name"].str.contains("Prerelease", case=False, na=False)
            prerelease_count = prerelease_mask.sum()
            excluded_prerelease = ref_data[prerelease_mask]
            print(f"Excluding {prerelease_count} prerelease cards based on FILTER_PRERELEASE setting.")
            print("Excluded prerelease entries:")
            print(excluded_prerelease["Product Name"].head(10))
            ref_data = ref_data[~prerelease_mask]

        # Log remaining entries for validation
        print("Sample of remaining entries after filtering:")
        print(ref_data["Product Name"].head(10))

        # Normalize data into a dictionary
        reference_data = {}
        for _, row in ref_data.iterrows():
            key = normalize_key(
                row.get("Product Name", ""),
                row.get("Set Name", ""),
                row.get("Condition", "Near Mint"),
                row.get("Number", "")
            )
            if key:  # Exclude invalid keys
                reference_data[key] = row.to_dict()

        # Validate prerelease exclusion in reference_data
        prerelease_in_dict = [k for k in reference_data.keys() if
                              "prerelease" in reference_data[k]["Product Name"].lower()]
        if prerelease_in_dict:
            print(
                f"Warning: {len(prerelease_in_dict)} prerelease entries remain in reference_data. Investigate keys: {prerelease_in_dict[:5]}")
        else:
            print("No prerelease entries remain in reference_data.")

        return reference_data
    except FileNotFoundError:
        print(f"Reference file {reference_csv} not found. Exiting.")
        exit()


def find_best_match(normalized_key, reference_data):
    """
    Find the best match for the given key in reference data.
    Apply penalties for mismatched attributes and boost scores for exact matches.
    """
    matches = []

    for ref_key in reference_data.keys():
        base_score = fuzz.ratio(normalized_key[0], ref_key[0])  # Name similarity score

        # Boost for exact matches on name, set, collector number, and condition
        if normalized_key[1] == ref_key[1]:
            base_score += 50  # Set match boost
        if normalized_key[2] == ref_key[2]:
            base_score += 100  # Collector number match boost
        if normalized_key[3] == ref_key[3]:
            base_score += 50  # Condition match boost (e.g., "Foil")

        # Penalize mismatched attributes
        if normalized_key[2] != ref_key[2]:  # Mismatched collector number
            base_score -= 30
        if normalized_key[3] != ref_key[3]:  # Mismatched condition
            base_score -= 20

        # Skip prerelease matches explicitly
        if "prerelease" in reference_data[ref_key]["Product Name"].lower() or "prerelease cards" in \
                reference_data[ref_key]["Set Name"].lower():
            continue

        matches.append((ref_key, base_score))

    # Log adjusted matches
    matches.sort(key=lambda x: x[1], reverse=True)  # Sort by adjusted score
    return matches


def confirm_and_iterate_match(normalized_key, matches, reference_data):
    """
    Iterate through matches until a correct match is confirmed.
    Automatically confirm matches with a score >= 250.
    Allow users to give up ("g") on a match.
    """
    print(
        f"Attempting to match card: {normalized_key[0]} from set {normalized_key[1]} with number {normalized_key[2]}.")
    for i, (match, adjusted_score) in enumerate(matches, start=1):
        ref_row = reference_data.get(match, {})
        if adjusted_score >= 250:
            print(
                f"Automatically confirming: {ref_row.get('Product Name', 'Unknown')} | Set: {ref_row.get('Set Name', 'Unknown')} | Card Number: {ref_row.get('Number', 'Unknown')} | Adjusted Score: {adjusted_score}")
            confirmed_matches[normalized_key] = match
            return match

        print(
            f"Potential match {i}: {ref_row.get('Product Name', 'Unknown')} | Set: {ref_row.get('Set Name', 'Unknown')} | Card Number: {ref_row.get('Number', 'Unknown')} | Adjusted Score: {adjusted_score}")
        response = input("Is this correct? (y/n/g): ").strip().lower()
        if response == "y":
            confirmed_matches[normalized_key] = match
            return match
        elif response == "g":
            given_up_cards.append({"Name": normalized_key[0], "Set": normalized_key[1], "Number": normalized_key[2]})
            return None
    print("No match confirmed.")
    return None


def map_fields(manabox_row, reference_data):
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
        return None

    normalized_result = normalize_key(card_name, set_name, condition, card_number)
    if not normalized_result:
        print(f"Skipping invalid or prerelease card: {card_name} from set {set_name}")
        return None

    key = normalized_result[:4]

    if key in confirmed_matches:
        ref_row = reference_data[confirmed_matches[key]]
        return {
            "TCGplayer Id": ref_row.get("TCGplayer Id", "Not Found"),
            "Product Line": ref_row.get("Product Line", "Magic: The Gathering"),
            "Set Name": ref_row.get("Set Name", ""),
            "Product Name": ref_row.get("Product Name", "") + normalized_result[4],
            "Number": ref_row.get("Number", ""),
            "Rarity": ref_row.get("Rarity", ""),
            "Condition": condition,
            "Add to Quantity": int(manabox_row.get("Quantity", "1")),
            "TCG Marketplace Price": manabox_row.get("Purchase price", "0.00")
        }

    matches = find_best_match(key, reference_data)
    if matches:
        confirmed_match = confirm_and_iterate_match(key, matches, reference_data)
        if confirmed_match:
            ref_row = reference_data[confirmed_match]
            return {
                "TCGplayer Id": ref_row.get("TCGplayer Id", "Not Found"),
                "Product Line": ref_row.get("Product Line", "Magic: The Gathering"),
                "Set Name": ref_row.get("Set Name", ""),
                "Product Name": ref_row.get("Product Name", "") + normalized_result[4],
                "Number": ref_row.get("Number", ""),
                "Rarity": ref_row.get("Rarity", ""),
                "Condition": condition,
                "Add to Quantity": int(manabox_row.get("Quantity", "1")),
                "TCG Marketplace Price": manabox_row.get("Purchase price", "0.00")
            }
    print(
        f"No match found for card: {normalized_result[0]} from set {normalized_result[1]} with number {normalized_result[2]}.")
    return None


def merge_entries(cards):
    """
    Merge duplicate entries based on card ID, condition, and other attributes.
    """
    merged = {}
    for card in cards:
        key = (card['TCGplayer Id'], card['Condition'])
        if key in merged:
            merged[key]['Add to Quantity'] += card['Add to Quantity']
        else:
            merged[key] = card
    return list(merged.values())


def auto_confirm_high_score(cards):
    """
    Automatically confirm cards with a score of 250 or higher.
    """
    confirmed = []
    for card in cards:
        if card.get('Score', 0) >= 250:
            confirmed.append(card)
    return confirmed


# Get input and output file paths using a file dialog
Tk().withdraw()  # Hide the root window
manabox_csv = askopenfilename(title="Select the Manabox CSV File", filetypes=[("CSV Files", "*.csv")])
if not manabox_csv:
    print("No file selected. Exiting.")
    exit()

tcgplayer_csv = "tcgplayer_staged.csv"
reference_csv = "REFERENCE.csv"

reference_data = load_reference_data(reference_csv)

try:
    with open(manabox_csv, mode='r', newline='', encoding='utf-8') as infile, open(tcgplayer_csv, mode='w', newline='',
                                                                                   encoding='utf-8') as outfile:
        reader = csv.DictReader(infile)
        fieldnames = ["TCGplayer Id", "Product Line", "Set Name", "Product Name", "Number", "Rarity", "Condition",
                      "Add to Quantity", "TCG Marketplace Price"]
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()

        cards = []
        for row in reader:
            tcgplayer_row = map_fields(row, reference_data)
            if tcgplayer_row:
                cards.append(tcgplayer_row)

        merged_cards = merge_entries(cards)
        auto_confirmed = auto_confirm_high_score(merged_cards)

        for card in merged_cards:
            writer.writerow(card)

    print("Conversion complete. Output saved to tcgplayer_staged.csv")

    if given_up_cards:
        print("Cards given up on:")
        for card in given_up_cards:
            print(f"Name: {card['Name']}, Set: {card['Set']}, Number: {card['Number']}")

except FileNotFoundError as e:
    print(f"Error: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
