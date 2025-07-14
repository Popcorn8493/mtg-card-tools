# MtG-Card-Tools

These scripts help manage your Magic: The Gathering card inventory for TCGplayer.

### Disclaimer 

This is a work in progress. The accuracy of the output is not guaranteed, so please verify the results after running the scripts.

-----

### Setup & Installation

First, set up a virtual environment to keep your project dependencies isolated.

1.  **Create a virtual environment:**

    ```bash
    python -m venv venv
    ```

2.  **Activate the virtual environment:**

      * **Windows:**
        ```bash
        .\venv\Scripts\activate
        ```
      * **macOS/Linux:**
        ```bash
        source venv/bin/activate
        ```

3.  **Install dependencies:**
    Install all the necessary packages from the `requirements.txt` file.

    ```bash
    pip install -r requirements.txt
    ```

-----

### How to Use

####  `convert_manabox_to_tcgplayer.py`

This script converts a CSV export from Manabox to a TCGplayer-compatible format.

1.  **Get your TCGplayer reference file:**

      * From your TCGplayer seller portal, go to the **Pricing** tab.
      * Click **Export Filtered CSV**.
      * Leave the options as default, but make sure **Export only from Live Inventory** is unchecked.
      * Save this file as `REFERENCE.csv` in the same folder as the script.

2.  **Run the script:**

    ```bash
    python convert_manabox_to_tcgplayer.py
    ```

3.  **Follow the prompts:**

      * A window will pop up asking you to select your Manabox CSV file.
      * The script will then try to match each card. You may be prompted to confirm matches:
          * Press **Y** to confirm a match.
          * Press **N** to reject it and see the next suggestion.
          * Press **G** to give up on a card and move to the next.
      * The output will be saved as `tcgplayer_staged.csv` and any cards you gave up on will be in `tcgplayer_given_up.csv`.

####  `update_tcgplayer_prices.py`

This script updates the prices in your TCGplayer inventory CSV based on a variety of parameters that can be adjusted as needed, wanted, or desired. Always double check after running and before uploading for correct quantities, prices, etc.

1.  **Run the script:**

    ```bash
    python update_tcgplayer_prices.py
    ```

2.  **Select your file:**

      * A window will pop up asking you to select your TCGplayer inventory CSV.

3.  **Get the output:**

      * The script will create a new file named `Updated_TCGplayer_Inventory.csv` with the updated prices.
