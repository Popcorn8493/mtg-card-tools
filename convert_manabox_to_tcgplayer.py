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
    """
    Returns True if the product name appears to be double sided.
    Checks for a literal '//' or the presence of both 'double' and 'sided'.
    """
    pn = product_name.lower()
    return '//' in pn or ('double' in pn and 'sided' in pn)


def get_market_price(manabox_row, ref_row=None):
    """
    Determines a valid market price using multiple candidate fields.
    Order of precedence:
      1. Check the reference row candidate fields: "TCG Marketplace Price", "List Price", "Retail Price".
      2. Then check the CSV row's "Purchase price".
      3. If none yields a valid (nonzero) price, return the FLOOR_PRICE.
    """
    candidate_fields = ["TCG Marketplace Price", "List Price", "Retail Price"]
    if ref_row:
        for field in candidate_fields:
            price = str(ref_row.get(field, "")).strip()
            try:
                if price and float(price) > 0:
                    return price
            except ValueError:
                continue
    # Next, try the CSV row
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
    """
    Normalize card name, set name, condition, and card number for matching.
    For tokens, only the portion before any '//' is used.
    """
    suffix = ""
    if "(" in card_name and ")" in card_name:
        card_name = re.sub(r"\(.*?\)", "", card_name).strip()

    # For matching, use only text before '//' (if present)
    card_name = card_name.split('//')[0].strip()
    normalized_card_name = re.sub(r"[^a-zA-Z0-9 ,'-]", "", card_name).strip().lower()
    if "foil" in condition.lower():
        normalized_card_name += " foil"
    normalized_set_name = re.sub(r"[^a-zA-Z0-9 ]", "", set_name).strip().lower()
    if normalized_set_name in ["plst", "the list"]:
        normalized_set_name = "the list reprints"

    # Exclude prerelease sets explicitly.
    if "prerelease cards" in normalized_set_name:
        return None

    if normalized_set_name == "the list":
        number = number.split("-")[-1] if number else ""

    normalized_number = re.sub(r"^[A-Za-z\-]*", "", str(number).strip()) if number else None
    return normalized_card_name, normalized_set_name, normalized_number, condition.lower(), suffix


def load_reference_data(reference_csv):
    """
    Load and clean reference data for matching.
    """
    try:
        ref_data = pd.read_csv(reference_csv, dtype={"Number": "str"})
        ref_data = ref_data[ref_data["Set Name"].notnull()]

        if FILTER_PRERELEASE:
            prerelease_mask = ref_data["Product Name"].str.contains("Prerelease", case=False, na=False)
            prerelease_count = prerelease_mask.sum()
            excluded_prerelease = ref_data[prerelease_mask]
            print(f"Excluding {prerelease_count} prerelease cards based on FILTER_PRERELEASE setting.")
            print("Excluded prerelease entries:")
            print(excluded_prerelease["Product Name"].head(10))
            ref_data = ref_data[~prerelease_mask]

        print("Sample of remaining entries after filtering:")
        print(ref_data["Product Name"].head(10))

        reference_data = {}
        for _, row in ref_data.iterrows():
            key = normalize_key(
                row.get("Product Name", ""),
                row.get("Set Name", ""),
                row.get("Condition", "Near Mint"),
                row.get("Number", "")
            )
            if key:
                reference_data[key] = row.to_dict()

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
    Find the best match for the given key in reference data using fuzzy matching.
    """
    matches = []
    for ref_key in reference_data.keys():
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
        if "prerelease" in reference_data[ref_key]["Product Name"].lower() or "prerelease cards" in \
                reference_data[ref_key]["Set Name"].lower():
            continue
        matches.append((ref_key, base_score))
    matches.sort(key=lambda x: x[1], reverse=True)
    return matches


def confirm_and_iterate_match(normalized_key, matches, reference_data):
    """
    Iterate through fuzzy-matched candidates until one is confirmed.
    """
    print(
        f"Attempting to match card: {normalized_key[0]} from set {normalized_key[1]} with number {normalized_key[2]}.")
    for i, (match, adjusted_score) in enumerate(matches, start=1):
        ref_row = reference_data.get(match, {})
        if adjusted_score >= 300:
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
    Converts a row from the Manabox CSV into the TCGPlayer staged inventory format.
    Handles tokens specially. If the set name contains "Tokens" or the set code begins with 'T'
    (e.g. "TMOM"), then the row is treated as a token.
    For tokens, if the CSV does not show two sides (no "//") but the user indicates it is double sided,
    candidate double-sided entries are sought from reference data regardless of which side was scanned.
    The market price is determined by trying candidate fields from the reference row, then CSV,
    and finally falling back to a floor price.
    """
    card_name = manabox_row.get("Name", "").strip()
    set_name = manabox_row.get("Set name", "").strip()
    condition_code = manabox_row.get("Condition", "near_mint").strip().lower()
    foil = "Foil" if manabox_row.get("Foil", "normal").lower() == "foil" else ""
    condition = CONDITION_MAP.get(condition_code, "Near Mint")
    if foil:
        condition += " Foil"

    # Identify token rows.
    is_token = False
    if "token" in set_name.lower() or "token" in card_name.lower():
        is_token = True
    elif set_name.startswith("T") and re.match(r"^T[A-Z0-9]+$", set_name):
        is_token = True

    if is_token:
        # Form a canonical token set name.
        if set_name.startswith("T") and re.match(r"^T[A-Z0-9]+$", set_name):
            token_set_name = set_name[1:] + " tokens"
        else:
            token_set_name = set_name

        # For more-flexible matching on set names, derive a base by removing " tokens"
        token_set_base = token_set_name.lower().replace(" tokens", "")
        card_number = manabox_row.get("Collector number", "").strip()

        # Build the initial token product name.
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

        # Restrict reference data to token entries.
        token_reference_data = {k: v for k, v in reference_data.items()
                                if (("token" in v.get("Set Name", "").lower() or "token" in v.get("Product Name",
                                                                                                  "").lower())
                                    and (token_set_name.lower() in v.get("Set Name",
                                                                         "").lower() or token_set_base in v.get(
                        "Set Name", "").lower()))}

        # Find candidate matches via fuzzy matching.
        matches = find_best_match(normalized_token_key[:4], token_reference_data)
        chosen_match = None

        # DOUBLE-SIDED TOKEN HANDLING:
        if "//" not in card_name:
            ds_response = input(
                f"Token '{card_name}' from set '{set_name}' does not indicate two sides. Is it a double sided token? (y/n): ").strip().lower()
            if ds_response == "y":
                ds_candidates = []
                scanned_lower = card_name.lower()
                for k, v in token_reference_data.items():
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
                        candidate = token_reference_data[candidate_key]
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
                candidate = token_reference_data[m]
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
            ref_row = token_reference_data[chosen_match]
            token_product_name = ref_row.get("Product Name", token_product_name)
            token_number = ref_row.get("Number", card_number)
            market_price = get_market_price(manabox_row, ref_row)
            return {
                "TCGplayer Id": ref_row.get("TCGplayer Id", "Not Found"),
                "Product Line": ref_row.get("Product Line", "Magic: The Gathering"),
                "Set Name": ref_row.get("Set Name", token_set_name),
                "Product Name": token_product_name,
                "Number": token_number,
                "Rarity": ref_row.get("Rarity", "Token"),
                "Condition": condition,
                "Add to Quantity": int(manabox_row.get("Quantity", "1")),
                "TCG Marketplace Price": market_price
            }
        # Fallback: if no candidate is chosen, build a generic token entry using CSV values.
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

    # STANDARD (NON-TOKEN) PROCESSING:
    card_number = re.sub(r"^[A-Za-z\-]*", "", manabox_row.get("Collector number", "").strip().split("-")[-1])
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


# Main file I/O
Tk().withdraw()  # Hide the root window
manabox_csv = askopenfilename(title="Select the Manabox CSV File", filetypes=[("CSV Files", "*.csv")])
if not manabox_csv:
    print("No file selected. Exiting.")
    exit()

tcgplayer_csv = "tcgplayer_staged.csv"
reference_csv = "TCGplayer__Pricing_Custom_Export_20250206_015825.csv"

reference_data = load_reference_data(reference_csv)

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
