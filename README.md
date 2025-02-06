# MtG-Card-Tools

Work in progress, tool needs some work and improvements and I make no guarantee on accuracies so use at your own risk and verify the information is accurate after using the scripts

Download and work within the repo zip

Install dependencies with pip:
pip install pandas
pip install rapidfuzz

You will need TCGPlayer Level 4 seller to benefit from this, and also to download the CSV reference file.
Go to Pricing tab in TCGplayer seller portal, Export Filtered CSV, Leave options as default but notably leave Export only from Live Inventory unchecked. Rename that CSV to REFERENCE.csv and save next to convert_manabox_to_tcgplayer.py.

Put your manabox csv somewhere handy like in the same folder as the project. 

Run CMD and enter python convert_manabox_to_tcgplayer.py, select your manabox csv and you'll see the script attempt to match each entry across the reference csv. You'll get some prompts asking you to manually confirm a match, press Y to confirm it is correct, N to reject and try the next possible match, or G to give up and move on. For tokens I just press G. 
