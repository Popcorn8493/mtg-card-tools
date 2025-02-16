import csv
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from rapidfuzz import fuzz

# Option to filter out prerelease cards
FILTER_PRERELEASE = True  # Exclude prerelease cards
PRERELEASE_PENALTY = -20
EXACT_MATCH_BOOST = 200

# Alias mappings for sets (if needed)
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

# Set a floor price for tokens (if no valid price is found)
FLOOR_PRICE = 0.10

# Track cards the user has given up on and confirmed matches.
given_up_cards = []
confirmed_matches = {}


def is_double_sided_candidate(product_name):
    """Returns True if the product name appears to be double-sided."""
    pn = product_name.lower()
    return '//' in pn or ('double' in pn and 'sided' in pn)


def get_market_price(manabox_row, ref_row=None):
    """Determines a valid market price using multiple candidate fields."""
    candidate_fields = ["TCG Marketplace Price", "List Price", "Retail Price"]
    if ref_row:
        for field in candidate_fields:
            price = str(ref_row.get(field, "")).strip()
            try:
                if price and float(price) > 0:
                    return price
            except ValueError:
                continue
    csv_candidate_fields = ["Purchase price"]
    for field in csv_candidate_fields:
        price = str(manabox_row.get(field, "")).strip()
        try:
            if price and float(price) > 0:
                return price
        except ValueError:
            continue
    return f"{FLOOR_PRICE:.2f}"


def normalize_key(card_name, set_name, condition, number):
    """Normalize card name, set name, condition, and card number for matching."""
    suffix = ""
    if "(" in card_name and ")" in card_name:
        card_name = re.sub(r"\(.*?\)", "", card_name).strip()
    # Use only text before '//' (if present)
    card_name = card_name.split('//')[0].strip()
    normalized_card_name = re.sub(r"[^a-zA-Z0-9 ,'-]", "", card_name).strip().lower()
    if "foil" in condition.lower():
        normalized_card_name += " foil"
    normalized_set_name = re.sub(r"[^a-zA-Z0-9 ]", "", set_name).strip().lower()
    if normalized_set_name in ["plst", "the list"]:
        normalized_set_name = "the list reprints"
    if "prerelease cards" in normalized_set_name:
        return None
    if normalized_set_name == "the list":
        number = number.split("-")[-1] if number else ""
    normalized_number = re.sub(r"^[A-Za-z\-]*", "", str(number).strip()) if number else None
    return normalized_card_name, normalized_set_name, normalized_number, condition.lower(), suffix


def load_reference_data(reference_csv):
    """Load and clean reference data for matching."""
    try:
        ref_df = pd.read_csv(reference_csv, dtype={"Number": "str"})
        ref_df = ref_df[ref_df["Set Name"].notnull()]
        if FILTER_PRERELEASE:
            prerelease_mask = ref_df["Product Name"].str.contains("Prerelease", case=False, na=False)
            prerelease_count = prerelease_mask.sum()
            excluded = ref_df[prerelease_mask]
            print(f"Excluding {prerelease_count} prerelease cards based on FILTER_PRERELEASE setting.")
            print("Excluded prerelease entries:")
            print(excluded["Product Name"].head(10))
            ref_df = ref_df[~prerelease_mask]
        print("Sample of remaining entries after filtering:")
        print(ref_df["Product Name"].head(10))
        ref_data = {}
        for _, row in ref_df.iterrows():
            key = normalize_key(
                row.get("Product Name", ""),
                row.get("Set Name", ""),
                row.get("Condition", "Near Mint"),
                row.get("Number", "")
            )
            if key:
                ref_data[key] = row.to_dict()
        prerelease_keys = [k for k in ref_data.keys() if "prerelease" in ref_data[k]["Product Name"].lower()]
        if prerelease_keys:
            print(
                f"Warning: {len(prerelease_keys)} prerelease entries remain in reference data. Investigate keys: {prerelease_keys[:5]}")
        else:
            print("No prerelease entries remain in reference data.")
        return ref_data
    except FileNotFoundError:
        print(f"Reference file {reference_csv} not found. Exiting.")
        exit()


def find_best_match(normalized_key, ref_data):
    """Find the best match for the given key in reference data using fuzzy matching."""
    matches = []
    for ref_key in ref_data.keys():
        base_score = fuzz.ratio(normalized_key[0], ref_key[0])
        if normalized_key[1] == ref_key[1]:
            base_score += 50
        if normalized_key[2] == ref_key[2] or not normalized_key[2]:
            base_score += 100
        if normalized_key[3] == ref_key[3]:
            base_score += 50
        if normalized_key[2] and normalized_key[2] != ref_key[2]:
            base_score -= 30
        if normalized_key[3] != ref_key[3]:
            base_score -= 20
        if "prerelease" in ref_data[ref_key]["Product Name"].lower() or "prerelease cards" in ref_data[ref_key][
            "Set Name"].lower():
            continue
        matches.append((ref_key, base_score))
    matches.sort(key=lambda x: x[1], reverse=True)
    return matches


def confirm_and_iterate_match(normalized_key, matches, ref_data):
    """Iterate through fuzzy-matched candidates until one is confirmed."""
    print(
        f"Attempting to match card: {normalized_key[0]} from set {normalized_key[1]} with number {normalized_key[2]}.")
    for i, (match, score) in enumerate(matches, start=1):
        candidate = ref_data.get(match, {})
        if score >= 300:
            print(
                f"Automatically confirming: {candidate.get('Product Name', 'Unknown')} | Set: {candidate.get('Set Name', 'Unknown')} | Card Number: {candidate.get('Number', 'Unknown')} | Adjusted Score: {score}")
            confirmed_matches[normalized_key] = match
            return match
        print(
            f"Potential match {i}: {candidate.get('Product Name', 'Unknown')} | Set: {candidate.get('Set Name', 'Unknown')} | Card Number: {candidate.get('Number', 'Unknown')} | Adjusted Score: {score}")
        response = input("Is this correct? (y/n/g): ").strip().lower()
        if response == "y":
            confirmed_matches[normalized_key] = match
            return match
        elif response == "g":
            given_up_cards.append({"Name": normalized_key[0], "Set": normalized_key[1], "Number": normalized_key[2]})
            return None
    print("No match confirmed.")
    return None


def build_standard_entry(ref_row, product_name_suffix, manabox_row, condition):
    """Helper to build a standard entry dictionary."""
    return {
        "TCGplayer Id": ref_row.get("TCGplayer Id", "Not Found"),
        "Product Line": ref_row.get("Product Line", "Magic: The Gathering"),
        "Set Name": ref_row.get("Set Name", ""),
        "Product Name": ref_row.get("Product Name", "") + product_name_suffix,
        "Number": ref_row.get("Number", ""),
        "Rarity": ref_row.get("Rarity", ""),
        "Condition": condition,
        "Add to Quantity": int(manabox_row.get("Quantity", "1")),
        "TCG Marketplace Price": manabox_row.get("Purchase price", "0.00")
    }


def build_token_entry(ref_row, token_set_name, token_product_name, token_number, manabox_row, condition):
    """Helper to build a token entry dictionary from a confirmed match."""
    return {
        "TCGplayer Id": ref_row.get("TCGplayer Id", "Not Found"),
        "Product Line": ref_row.get("Product Line", "Magic: The Gathering"),
        "Set Name": ref_row.get("Set Name", token_set_name),
        "Product Name": ref_row.get("Product Name", token_product_name),
        "Number": ref_row.get("Number", token_number),
        "Rarity": ref_row.get("Rarity", "Token"),
        "Condition": condition,
        "Add to Quantity": int(manabox_row.get("Quantity", "1")),
        "TCG Marketplace Price": get_market_price(manabox_row, ref_row)
    }


def build_token_fallback(token_set_name, token_product_name, card_number, manabox_row, condition):
    """Helper to build a fallback token entry dictionary when no match is found."""
    fallback_price = get_market_price(manabox_row, None)
    return {
        "TCGplayer Id": "Not Found",
        "Product Line": "Magic: The Gathering",
        "Set Name": token_set_name,
        "Product Name": token_product_name,
        "Number": card_number,
        "Rarity": "Token",
        "Condition": condition,
        "Add to Quantity": int(manabox_row.get("Quantity", "1")),
        "TCG Marketplace Price": fallback_price
    }


def process_standard(manabox_row, ref_data, condition, card_name, set_name):
    """Process non-token (standard) card rows."""
    card_number = re.sub(r"^[A-Za-z\-]*", "", manabox_row.get("Collector number", "").strip().split("-")[-1])
    if not card_name or not set_name:
        return None
    normalized_result = normalize_key(card_name, set_name, condition, card_number)
    if not normalized_result:
        print(f"Skipping invalid or prerelease card: {card_name} from set {set_name}")
        return None
    key = normalized_result[:4]
    if key in confirmed_matches:
        ref_row = ref_data[confirmed_matches[key]]
        return build_standard_entry(ref_row, normalized_result[4], manabox_row, condition)
    matches = find_best_match(key, ref_data)
    if matches:
        confirmed_match = confirm_and_iterate_match(key, matches, ref_data)
        if confirmed_match:
            ref_row = ref_data[confirmed_match]
            return build_standard_entry(ref_row, normalized_result[4], manabox_row, condition)
    print(
        f"No match found for card: {normalized_result[0]} from set {normalized_result[1]} with number {normalized_result[2]}.")
    return None


def process_token(manabox_row, ref_data, condition, card_name, set_name):
    """Process token card rows."""
    # Determine canonical token set name.
    if set_name.startswith("T") and re.match(r"^T[A-Z0-9]+$", set_name):
        token_set_name = set_name[1:] + " tokens"
    else:
        token_set_name = set_name
    token_set_base = token_set_name.lower().replace(" tokens", "")
    card_number = manabox_row.get("Collector number", "").strip()
    if "//" in card_name:
        parts = card_name.split("//")
        side1 = parts[0].strip()
        side2 = re.sub(r"double[-\s]?sided token", "", parts[1], flags=re.IGNORECASE).strip()
        token_product_name = f"{side1} // {side2}"
    else:
        token_product_name = card_name

    normalized_token_key = normalize_key(token_product_name, token_set_name, condition, card_number)
    if not normalized_token_key:
        print(f"Skipping invalid or prerelease token: {card_name} from set {set_name}")
        return None

    token_ref_data = {k: v for k, v in ref_data.items()
                      if (("token" in v.get("Set Name", "").lower() or "token" in v.get("Product Name", "").lower())
                          and (token_set_name.lower() in v.get("Set Name", "").lower() or token_set_base in v.get(
                    "Set Name", "").lower()))}
    matches = find_best_match(normalized_token_key[:4], token_ref_data)
    chosen_match = None
    if "//" not in card_name:
        ds_response = input(
            f"Token '{card_name}' from set '{set_name}' does not indicate two sides. Is it a double sided token? (y/n): ").strip().lower()
        if ds_response == "y":
            ds_candidates = []
            scanned_lower = card_name.lower()
            for k, v in token_ref_data.items():
                prod_name = v.get("Product Name", "")
                if is_double_sided_candidate(prod_name):
                    sides = prod_name.split("//")
                    sides = [re.sub(r"doubled?-sided token", "", s, flags=re.IGNORECASE).strip() for s in sides]
                    if any(scanned_lower in side.lower() for side in sides) or any(
                            fuzz.ratio(scanned_lower, side.lower()) > 70 for side in sides):
                        scores = [fuzz.ratio(scanned_lower, side.lower()) for side in sides]
                        ds_candidates.append((k, max(scores)))
            if ds_candidates:
                ds_candidates.sort(key=lambda x: x[1], reverse=True)
                for candidate_key, score in ds_candidates:
                    candidate = token_ref_data[candidate_key]
                    resp = input(
                        f"Is the double sided token candidate '{candidate.get('Product Name')}' (score {score}) correct? (y/n/g): ").strip().lower()
                    if resp == "y":
                        chosen_match = candidate_key
                        break
                    elif resp == "g":
                        chosen_match = None
                        break
            else:
                print("No double sided candidate entries found in reference data.")
        else:
            if matches:
                best_match, best_score = matches[0]
                if best_score >= 250:
                    chosen_match = best_match
    else:
        for m, s in matches:
            candidate = token_ref_data[m]
            if is_double_sided_candidate(candidate.get("Product Name", "")):
                resp = input(
                    f"Double sided token candidate: '{candidate.get('Product Name')}' (score {s}). Is this correct? (y/n/g): ").strip().lower()
                if resp == "y":
                    chosen_match = m
                    break
                elif resp == "g":
                    chosen_match = None
                    break

    if chosen_match:
        ref_row = token_ref_data[chosen_match]
        token_product_name = ref_row.get("Product Name", token_product_name)
        token_number = ref_row.get("Number", card_number)
        return build_token_entry(ref_row, token_set_name, token_product_name, token_number, manabox_row, condition)
    return build_token_fallback(token_set_name, token_product_name, card_number, manabox_row, condition)


def map_fields(manabox_row, ref_data):
    """Converts a row from the Manabox CSV into the TCGPlayer staged inventory format."""
    card_name = manabox_row.get("Name", "").strip()
    set_name = manabox_row.get("Set name", "").strip()
    condition_code = manabox_row.get("Condition", "near_mint").strip().lower()
    foil = "Foil" if manabox_row.get("Foil", "normal").lower() == "foil" else ""
    condition = CONDITION_MAP.get(condition_code, "Near Mint")
    if foil:
        condition += " Foil"

    # Decide if this row is a token or a standard card.
    is_token = False
    if "token" in set_name.lower() or "token" in card_name.lower():
        is_token = True
    elif set_name.startswith("T") and re.match(r"^T[A-Z0-9]+$", set_name):
        is_token = True

    if is_token:
        return process_token(manabox_row, ref_data, condition, card_name, set_name)
    else:
        return process_standard(manabox_row, ref_data, condition, card_name, set_name)


def merge_entries(cards):
    """Merge duplicate entries based on card ID, condition, and other attributes."""
    merged = {}
    for card in cards:
        key = (card['TCGplayer Id'], card['Condition'])
        if key in merged:
            merged[key]['Add to Quantity'] += card['Add to Quantity']
        else:
            merged[key] = card
    return list(merged.values())


def auto_confirm_high_score(cards):
    """Automatically confirm cards with a score of 250 or higher."""
    confirmed = []
    for card in cards:
        if card.get('Score', 0) >= 250:
            confirmed.append(card)
    return confirmed


def select_csv_file(prompt):
    """Helper function to open a file dialog with a given prompt and return the file path."""
    file_path = askopenfilename(title=prompt, filetypes=[("CSV Files", "*.csv")])
    if not file_path:
        print(f"No file selected for {prompt}. Exiting.")
        exit()
    return file_path


# Main file I/O
Tk().withdraw()  # Hide the root window once
manabox_csv = select_csv_file("Select the Manabox CSV File")
reference_csv = select_csv_file("Select the TCGPlayer Reference CSV File")
tcgplayer_csv = "tcgplayer_staged.csv"
ref_data = load_reference_data(reference_csv)

try:
    with open(manabox_csv, mode='r', newline='', encoding='utf-8') as infile, \
            open(tcgplayer_csv, mode='w', newline='', encoding='utf-8') as outfile:
        reader = csv.DictReader(infile)
        fieldnames = ["TCGplayer Id", "Product Line", "Set Name", "Product Name", "Number", "Rarity",
                      "Condition", "Add to Quantity", "TCG Marketplace Price"]
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        cards = []
        for row in reader:
            tcgplayer_row = map_fields(row, ref_data)
            if tcgplayer_row:
                cards.append(tcgplayer_row)
        merged_cards = merge_entries(cards)
        auto_confirm_high_score(merged_cards)
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
