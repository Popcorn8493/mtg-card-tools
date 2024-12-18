import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Constants
MIN_PRICE = 0.20
PRICE_MULTIPLIER = 1.10

# Open a file dialog to select the input file
Tk().withdraw()  # Hide the root Tkinter window
input_file_path = askopenfilename(title="Select the input CSV file", filetypes=[("CSV files", "*.csv")])
output_file_path = 'Updated_TCGplayer_Inventory.csv'

if not input_file_path:
    print("No file selected. Exiting...")
else:
    try:
        # Load the inventory CSV
        df = pd.read_csv(input_file_path)


        # Update the TCG Marketplace Price column (Column O)
        def calculate_price(row):
            market_price = row.get('TCG Market Price', None)
            low_price = row.get('TCG Low Price', None)
            base_price = market_price if pd.notna(market_price) else low_price
            return max(base_price * PRICE_MULTIPLIER, MIN_PRICE) if pd.notna(base_price) else MIN_PRICE


        df['TCG Marketplace Price'] = df.apply(calculate_price, axis=1)

        # Save the updated file
        df.to_csv(output_file_path, index=False)

        print(f"Updated inventory saved to {output_file_path}")
    except FileNotFoundError:
        print(f"Error: The file at {input_file_path} was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
