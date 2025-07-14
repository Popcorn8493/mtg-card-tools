import csv
import re
import unicodedata
from tkinter import Tk, Toplevel, Listbox, Button, Label, Frame, messagebox, Scrollbar, END
from tkinter.filedialog import askopenfilename
import tkinter.font as tkFont
import pandas as pd
from rapidfuzz import fuzz

# Option to filter out prerelease cards
FILTER_PRERELEASE = True  # Exclude prerelease cards

# Alias mappings for sets (if needed)
SET_ALIAS = {
    "Universes Beyond: The Lord of the Rings: Tales of Middle-earth": "LTR",
    "Commander: The Lord of the Rings: Tales of Middle-earth": "LTC",
    "the list": "The List"
}

# Updated condition mapping (ensure underscores become spaces)
CONDITION_MAP = {
    "near mint": "Near Mint",
    "lightly played": "Lightly Played",
    "moderately played": "Moderately Played",
    "heavily played": "Heavily Played",
    "damaged": "Damaged"
}

# Mapping to rank conditions (lower is better)
condition_rank = {
    "near mint": 0,
    "lightly played": 1,
    "moderately played": 2,
    "heavily played": 3,
    "damaged": 4
}

# Set a floor price for tokens (if no valid price is found)
FLOOR_PRICE = 0.10

# Global lists to track confirmed matches and given-up cards.
given_up_cards = []
confirmed_matches = {}


def remove_accents(text):
    """Convert accented characters to their unaccented equivalents."""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', text)
        if not unicodedata.combining(c)
    )


def is_double_sided_candidate(product_name):
    """Returns True if the product name appears to be double-sided."""
    pn = product_name.lower()
    return '//' in pn or ('double' in pn and 'sided' in pn)


def get_market_price(manabox_row, ref_row=None):
    """Determine a valid market price using multiple candidate fields."""
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
    # Remove accents from card_name
    card_name = remove_accents(card_name)
    card_name = card_name.split('//')[0].strip()  # Use only text before '//' if present.
    normalized_card_name = re.sub(r"[^a-zA-Z0-9 ,'-]", "", card_name).strip().lower()
    # Remove accents from set_name
    set_name = remove_accents(set_name)
    normalized_set_name = re.sub(r"[^a-zA-Z0-9 ]", "", set_name).strip().lower()
    if normalized_set_name in ["plst", "the list"]:
        normalized_set_name = "the list reprints"
    if "prerelease cards" in normalized_set_name:
        return None
    if normalized_set_name == "the list":
        number = number.split("-")[-1] if number else ""
    normalized_number = re.sub(r"[^\d\-]", "", str(number).strip()) if number else None
    if normalized_number == "":
        normalized_number = None
    return normalized_card_name, normalized_set_name, normalized_number, condition.lower(), suffix


def build_given_up_entry(manabox_row, condition, card_name, set_name):
    """Build a fallback entry (in TCGplayer format) for cards given up on."""
    return {
        "TCGplayer Id": "Not Found",
        "Product Line": "Magic: The Gathering",
        "Set Name": set_name,
        "Product Name": card_name,
        "Number": manabox_row.get("Collector number", "").strip(),
        "Rarity": manabox_row.get("Rarity", ""),
        "Condition": condition,
        "Add to Quantity": int(manabox_row.get("Quantity", "1")),
        "TCG Marketplace Price": get_market_price(manabox_row, None)
    }


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


def find_best_match(normalized_key, card_database):
    """Find the best match for the given key in reference data using fuzzy matching."""
    matches = []
    for ref_key in card_database.keys():
        # Quick check: if the first letters differ, skip.
        if normalized_key[0] and ref_key[0] and normalized_key[0][0] != ref_key[0][0]:
            continue

        query_words = normalized_key[0].split()
        candidate_words = ref_key[0].split()
        if len(query_words) == 1 and len(candidate_words) == 1:
            if query_words[0] != candidate_words[0]:
                continue
        elif len(query_words) > 1 and len(candidate_words) > 1:
            if not set(query_words).intersection(set(candidate_words)):
                continue

        base_score = fuzz.ratio(normalized_key[0], ref_key[0])
        if normalized_key[0] in ref_key[0] or ref_key[0] in normalized_key[0]:
            base_score += 20
        if normalized_key[1] == ref_key[1]:
            base_score += 50

        if not normalized_key[2] or not ref_key[2]:
            base_score += 50
        elif normalized_key[2] == ref_key[2]:
            base_score += 100
        else:
            base_score -= 15

        cond1 = normalized_key[3].replace("foil", "").strip()
        cond2 = ref_key[3].replace("foil", "").strip()
        if cond1 in condition_rank and cond2 in condition_rank:
            diff = abs(condition_rank[cond1] - condition_rank[cond2])
            if diff == 0:
                base_score += 50
            elif diff == 1:
                base_score -= 10
            else:
                base_score -= 30
        else:
            if normalized_key[3] != ref_key[3]:
                base_score -= 20

        if ("prerelease" in ref_data[ref_key]["Product Name"].lower() or
                "prerelease cards" in ref_data[ref_key]["Set Name"].lower()):
            continue

        special_print_penalties = {
            "foil": 40,
            "showcase": 30,
            "etched": 30,
            "borderless": 30,
            "extended": 30,
            "gilded": 30
        }
        for term, penalty in special_print_penalties.items():
            in_query = term in normalized_key[3]
            in_ref = term in ref_key[3]
            if in_query != in_ref:
                base_score -= penalty

        matches.append((ref_key, base_score))
    matches.sort(key=lambda x: x[1], reverse=True)
    return matches


def confirm_match_gui(normalized_key, matches, reference_data, title="Select Correct Card"):
    """Opens a centered GUI window for candidate selection."""
    root = Tk()
    root.tk.call('tk', 'scaling', 2.0)
    root.withdraw()
    window = Toplevel(root)
    window.title(title)
    window.resizable(True, True)

    custom_font = tkFont.Font(family="Helvetica", size=14)

    instruction = (
        f"Select the correct match for: {normalized_key[0]} ({normalized_key[3]})\n"
        f"Set: {normalized_key[1]} | Number: {normalized_key[2]}"
    )
    Label(window, text=instruction, padx=10, pady=10, font=custom_font).pack()

    frame = Frame(window)
    frame.pack(padx=10, pady=5, fill="both", expand=True)

    listbox = Listbox(frame, height=10, font=custom_font)
    listbox.pack(side="left", fill="both", expand=True)

    x_scroll = Scrollbar(frame, orient="horizontal", command=listbox.xview)
    x_scroll.pack(side="bottom", fill="x")
    listbox.configure(xscrollcommand=x_scroll.set)

    candidate_options = []
    for idx, (match, score) in enumerate(matches):
        candidate = ref_data.get(match, {})
        candidate_condition = candidate.get("Condition", "Unknown")
        candidate_str = (f"{idx + 1}: {candidate.get('Product Name', 'Unknown')} | "
                         f"Set: {candidate.get('Set Name', 'Unknown')} | "
                         f"Number: {candidate.get('Number', 'Unknown')} | "
                         f"Candidate Condition: {candidate_condition} | "
                         f"Score: {score}")
        listbox.insert(END, candidate_str)
        candidate_options.append(match)
    if candidate_options:
        listbox.selection_set(0)

    selection = {"choice": None}

    def on_confirm():
        selected_indices = listbox.curselection()
        if selected_indices:
            selection["choice"] = candidate_options[selected_indices[0]]
        window.destroy()

    def on_giveup():
        selection["choice"] = None  # Change this from "giveup" to None
        window.destroy()

    button_frame = Frame(window)
    button_frame.pack(pady=10)
    confirm_button = Button(button_frame, text="Confirm Selection", command=on_confirm, font=custom_font)
    confirm_button.pack(side="left", padx=5)
    giveup_button = Button(button_frame, text="Give Up", command=on_giveup, font=custom_font)
    giveup_button.pack(side="left", padx=5)
    window.update_idletasks()
    min_width = 800
    w = window.winfo_width()
    h = window.winfo_height()
    ws = window.winfo_screenwidth()
    hs = window.winfo_screenheight()
    x = (ws // 2) - (max(w, min_width) // 2)
    y = (hs // 2) - (h // 2)
    window.geometry(f"{max(w, min_width)}x{h}+{x}+{y}")
    window.protocol("WM_DELETE_WINDOW", on_giveup)
    window.wait_window()
    root.destroy()
    if selection["choice"] == None:  # Check for None instead of "giveup"
        print(f"User gave up on matching card: {normalized_key[0]} from set {normalized_key[1]}")
        return None
    else:
        return selection["choice"]


def confirm_and_iterate_match(normalized_key, matches, ref_data):
    """
    Improved auto-confirm logic:
      1) If the top match is >= 270, auto-confirm immediately.
      2) If the top match is >= 260 and leads the second-best match by >= 30, auto-confirm.
      3) Otherwise, open the GUI for user selection.
    """
    print(f"Matching card: {normalized_key[0]} from set {normalized_key[1]} (Number: {normalized_key[2]})")
    best_match, best_score = matches[0]
    candidate = ref_data.get(best_match, {})
    second_best_score = matches[1][1] if len(matches) > 1 else 0
    if best_score >= 270:
        print(f"Auto-confirming high score match: {candidate.get('Product Name', 'Unknown')} | "
              f"Candidate Condition: {candidate.get('Condition', 'Unknown')} (Score: {best_score})")
        confirmed_matches[normalized_key] = best_match
        return best_match
    if best_score >= 260 and (best_score - second_best_score) >= 30:
        print(f"Auto-confirming strong lead match: {candidate.get('Product Name', 'Unknown')} | "
              f"Candidate Condition: {candidate.get('Condition', 'Unknown')} "
              f"(Score: {best_score}, 2nd Score: {second_best_score})")
        confirmed_matches[normalized_key] = best_match
        return best_match
    chosen_match = confirm_match_gui(normalized_key, matches, ref_data)
    if chosen_match:
        confirmed_matches[normalized_key] = chosen_match
    return chosen_match


def build_standard_entry(ref_row, product_name_suffix, manabox_row, condition):
    """Build a standard entry dictionary in TCGplayer format."""
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
    """Build a token entry dictionary from a confirmed match."""
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
    """Build a fallback token entry dictionary when no match is found."""
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


def map_fields(manabox_row, card_database):
    """Convert a row from the Manabox CSV into the TCGplayer staged inventory format."""
    card_name = manabox_row.get("Name", "").strip()
    set_name = manabox_row.get("Set name", "").strip()
    condition_code = manabox_row.get("Condition", "near mint").strip().lower().replace("_", " ")
    foil = "Foil" if manabox_row.get("Foil", "normal").lower() == "foil" else ""
    condition = CONDITION_MAP.get(condition_code, "Near Mint")
    if foil:
        condition += " Foil"
    is_token = (
            "token" in set_name.lower() or
            "token" in card_name.lower() or
            (set_name.startswith("T") and re.match(r"^T[A-Z0-9]+$", set_name))
    )
    if is_token:
        return process_token(manabox_row, card_database, condition, card_name, set_name)
    else:
        return process_standard(manabox_row, card_database, condition, card_name, set_name)


def process_standard(manabox_row, card_database, condition, card_name, set_name):
    """Process a standard (non-token) card row."""
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
    confirmed_match = None
    if matches:
        confirmed_match = confirm_and_iterate_match(key, matches, ref_data)
    if confirmed_match:
        ref_row = ref_data[confirmed_match]
        return build_standard_entry(ref_row, normalized_result[4], manabox_row, condition)
    else:
        fallback = build_given_up_entry(manabox_row, condition, card_name, set_name)
        given_up_cards.append(fallback)
        print(f"User gave up on matching card: {normalized_result[0]} from set {normalized_result[1]}")
        return None


def process_token(manabox_row, card_database, condition, card_name, set_name):
    """Process a token card row."""
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

    token_ref_data = {
        k: v for k, v in ref_data.items()
        if (
                ("token" in v.get("Set Name", "").lower() or "token" in v.get("Product Name", "").lower()) and
                (token_set_name.lower() in v.get("Set Name", "").lower() or token_set_base in v.get("Set Name",
                                                                                                    "").lower())
        )
    }
    matches = find_best_match(normalized_token_key[:4], token_ref_data)
    chosen_match = None
    if "//" not in card_name:
        is_ds = messagebox.askyesno(
            "Double Sided Token",
            f"Token '{card_name}' from set '{set_name}' does not indicate two sides. Is it a double sided token?"
        )
        if is_ds:
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
                confirm_match_gui(normalized_token_key, ds_candidates, token_ref_data,
                                  title="Select Double Sided Token")
            else:
                messagebox.showinfo("Info", "No double sided candidate entries found in reference data.")
        else:
            if matches:
                best_match, best_score = matches[0]
                if best_score >= 250:
                    chosen_match = best_match
                else:
                    confirm_match_gui(normalized_token_key, matches, token_ref_data,
                                      title="Select Token Match")
    else:
        ds_matches = [
            (m, s) for m, s in matches
            if is_double_sided_candidate(token_ref_data[m].get("Product Name", ""))
        ]
        if ds_matches:
            confirm_match_gui(normalized_token_key, ds_matches, token_ref_data,
                              title="Select Double Sided Token")
    if chosen_match:
        ref_row = token_ref_data[chosen_match]
        token_product_name = ref_row.get("Product Name", token_product_name)
        token_number = ref_row.get("Number", card_number)
        return build_token_entry(ref_row, token_set_name, token_product_name, token_number, manabox_row, condition)
    else:
        fallback = build_token_fallback(token_set_name, token_product_name, card_number, manabox_row, condition)
        given_up_cards.append(fallback)
        print(f"User gave up on matching token: {card_name} from set {set_name}")
        return None


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
    """Automatically confirm cards with a score of 250 or higher (unused by default)."""
    confirmed = []
    for card in cards:
        if card.get('Score', 0) >= 250:
            confirmed.append(card)
    return confirmed


def select_csv_file(prompt):
    """Open a file dialog with a given prompt and return the file path."""
    file_path = askopenfilename(title=prompt, filetypes=[("CSV Files", "*.csv")])
    if not file_path:
        print(f"No file selected for {prompt}. Exiting.")
        exit()
    return file_path


# Main file I/O
Tk().withdraw()
manabox_csv = select_csv_file("Select the Manabox CSV File")
#reference_csv = select_csv_file("Select the TCGPlayer Reference CSV File")
reference_csv = "TCGplayer__Pricing_Custom_Export_20250615_111402.csv"
tcgplayer_csv = "tcgplayer_staged.csv"
ref_data = load_reference_data(reference_csv)

try:
    with open(manabox_csv, mode='r', newline='', encoding='utf-8') as infile, \
            open(tcgplayer_csv, mode='w', newline='', encoding='utf-8') as outfile:
        reader = csv.DictReader(infile)
        fieldnames = [
            "TCGplayer Id", "Product Line", "Set Name", "Product Name",
            "Number", "Rarity", "Condition", "Add to Quantity", "TCG Marketplace Price"
        ]
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        cards = []
        for row in reader:
            tcgplayer_row = map_fields(row, ref_data)
            if tcgplayer_row:
                cards.append(tcgplayer_row)
        merged_cards = merge_entries(cards)
        for card in merged_cards:
            writer.writerow(card)
    print("Conversion complete. Output saved to tcgplayer_staged.csv")
    if given_up_cards:
        given_up_csv = "tcgplayer_given_up.csv"
        with open(given_up_csv, mode='w', newline='', encoding='utf-8') as gfile:
            gwriter = csv.DictWriter(gfile, fieldnames=fieldnames)
            gwriter.writeheader()
            for entry in given_up_cards:
                gwriter.writerow(entry)
        print(f"Given up cards saved to {given_up_csv}")
except FileNotFoundError as e:
    print(f"Error: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
