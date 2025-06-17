# Add the necessary imports
import pandas as pd
from rapidfuzz import fuzz, process

#  define a constant string for Restricted field
RESTRICTED_FIELD = "Restricted"
DESCRIPTION_FIELD = "Transaction Description"
PURPOSE_FIELD = "Purpose"
AMOUNT_FIELD = "Amount"
DESC_NAME_FIELD = "Description_Name"

def print_progress(message):
    """Prints a progress message to the console."""
    print(f"[INFO] {message}")

def process_data() -> pd.DataFrame:
    # Load the bank statement and members list
    print_progress("Loading bank statement and members list...")
    bank_df = pd.read_excel("bank_statement.xlsx")
    members_df = pd.read_excel("members_list.xlsx")
    print_progress(f"Loaded bank statement data shape: {bank_df.shape}")

    # mark the description field as a string to avoid issues with mixed types
    bank_df[DESCRIPTION_FIELD] = bank_df[DESCRIPTION_FIELD].astype(str)

    print_progress("Cleaning and processing bank statement data...")
    #  Rename "Business a/c" to "Amount" 
    bank_df.rename(columns={'Business a/c': AMOUNT_FIELD}, inplace=True)

    # Drop rows with NaN in Description or Amount
    bank_df.dropna(subset=[DESCRIPTION_FIELD, AMOUNT_FIELD], inplace=True)
    print_progress(f"Bank statement data shape after dropping NaN in Description or Amount: {bank_df.shape}")

    #  Drop STWDSHP rows from DESCRIPTION_FIELD
    bank_df = bank_df[~bank_df[DESCRIPTION_FIELD].str.contains('STWDSHP', case=False, na=False)]
    print_progress(f"Bank statement data shape after dropping STWDSHP rows: {bank_df.shape}")

    # Drop rows if the Purpose colomn does not contain the word 'Tithe' or 'Offering'
    bank_df = bank_df[bank_df[PURPOSE_FIELD].str.contains('Tithe', case=False, na=False) |
                 bank_df[PURPOSE_FIELD].str.contains('Offering', case=False, na=False)]
    print_progress(f"Filtered bank statement data shape after dropping non Tithes and Offering: {bank_df.shape}")
    #  Drop Restricted rows if the Restricted field exists
    if RESTRICTED_FIELD in bank_df.columns.to_list():
        # If the Restricted field exists, filter out rows where it is True
        bank_df = bank_df[~bank_df[RESTRICTED_FIELD].str.contains('Yes', case=False, na=False)]
    print_progress(f"Filtered bank statement data shape after removing restricted rows: {bank_df.shape}")
    
    # Drop all rows where Description is just a number
    bank_df =  bank_df[~pd.to_numeric(bank_df[DESCRIPTION_FIELD], errors='coerce').notnull()]
    print_progress(f"Bank statement data shape after dropping rows with numeric Description: {bank_df.shape}")

    # From the bank description extract the first two words into a Description_Name column
    bank_df[DESC_NAME_FIELD] = bank_df[DESCRIPTION_FIELD].astype(str).apply(lambda x: ' '.join(x.split()[:2]) if isinstance(x, str) else x)

    print_progress("Extracted Description_Name from bank statement data.")

    # Sum amounts by Description_Name
    bank_df = bank_df.groupby(DESC_NAME_FIELD, as_index=False)[AMOUNT_FIELD].sum()
    print_progress("Grouped bank statement data by Description_Name and summed amounts.")
    return bank_df

    #  Clean up the memebers list drop duplicate rowas with the same First Name and Last Name, keep the last occurrence with the most recent Start Time column
    members_df.drop_duplicates(subset=['First Name', 'Last Name'], keep='last', inplace=True)

    # Match the Description_Name with member names in the members list and create a new column Matched Member with the matched name and the ID from the members list
    members_df['Full Name'] = members_df['First Name'].str.strip() + " " + members_df['Last Name'].str.strip()
    member_names = members_df['Full Name'].tolist()
    bank_df['Matched Member'] = bank_df[DESC_NAME_FIELD].astype(str).apply(lambda x: process.extractOne(x, member_names, scorer=fuzz.partial_ratio)[0] if isinstance(x, str) else None)

    bank_df['Matched Member ID'] = bank_df['Matched Member'].apply(lambda x: members_df[members_df['Full Name'] == x]['ID'].values[0] if x in members_df['Full Name'].values else None)

    #  Write the processed bank_df to an Excel file
    bank_df.to_excel("processed_bank_statement.xlsx", index=False)

def load_and_clean_statement(eoy_file: str) -> pd.DataFrame:
    """Load and process the bank statement from the given file."""
    print_progress(f"Loading bank statement from {eoy_file}...")
    bank_df = pd.read_excel(eoy_file, "EOY 2023 MCR", header=None)
    print_progress(f"Loaded bank statement data shape: {bank_df.shape}")
    return bank_df

if __name__ == "__main__":
    bankd_df = process_data()
    print_progress(f"Processed bank statement data shape: {bankd_df.shape}")
    #  Write the processed bank_df to an Excel file
    bankd_df.to_excel("processed_bank_statement.xlsx", index=False)
    print("Data processing complete. Processed data saved to 'processed_bank_statement.xlsx'.")
else:
    print("This script is intended to be run as a standalone program.")
    # If imported, the process_data function can be called directly.
