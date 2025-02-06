"Download and work within the repo zip
Install dependencies with PIP:
pip install pandas
pip install rapidfuzz

Acquire a TCGplayer reference CSV:
Go to Pricing tab in TCGplayer seller portal, Export Filtered CSV, leave options as default but notably leave Export only from Live Inventory unchecked. Rename that CSV to REFERENCE.csv and save in project folder.

Put your manabox csv somewhere handy like in the project folder.

Run CMD and enter python convert_manabox_to_tcgplayer.py, select your manabox csv and you'll see the script attempt to match each entry across the reference csv. You'll get some prompts asking you to manually confirm a match, press Y to confirm it is correct, N to reject and try the next possible match, or G to give up and move on. For tokens I just press G.

The script will spit out a tcgplayer_staged.csv file. You can use Pricing tab then press Import To Staged to upload."

- /u/nrhjov
