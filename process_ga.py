# Add the necessary imports
import pandas as pd
from rapidfuzz import fuzz, process

#  define a constant string for Restricted field
RESTRICTED_FIELD = "Restricted"
DESCRIPTION_FIELD = "Transaction Description"
PURPOSE_FIELD = "Purpose"
AMOUNT_FIELD = "Amount"
DESC_NAME_FIELD = "Description_Name"
FIRST_NAME_FIELD = "First Name"
LAST_NAME_FIELD = "Last Name"
FULL_NAME_FIELD = "Full Name"
MATCHED_MEMBER_FIELD = "Matched Member"
MATCHED_MEMBER_ID_FIELD = "Matched Member ID"
MEMBERS_ID_FIELD = "ID"

def print_progress(message):
    """Prints a progress message to the console."""
    print(f"[INFO] {message}")

def process_data() -> pd.DataFrame:
    # Load the bank statement and members list
    print_progress("Loading bank statement and members list...")
    bank_df = load_and_clean_statement("bank_statement.xlsx")
    print_progress(f"Process Data: Loaded bank statement data shape: {bank_df.shape}")

    members_df = pd.read_excel("members_list.xlsx")
    print_progress(f"Process Data: Loaded members list data shape: {members_df.shape}")

    #  Clean up the memebers list drop duplicate rowas with the same First Name and Last Name, keep the last occurrence with the most recent Start Time column
    members_df.drop_duplicates(subset=[FIRST_NAME_FIELD, LAST_NAME_FIELD], keep='last', inplace=True)
    print_progress(f"Process Data: Members list data shape after dropping duplicates: {members_df.shape}")

    # Match the Description_Name with member names in the members list and create a new column Matched Member with the matched name and the ID from the members list
    members_df[FULL_NAME_FIELD] = members_df[FIRST_NAME_FIELD].str.strip() + " " + members_df[LAST_NAME_FIELD].str.strip()
    member_names = members_df[FULL_NAME_FIELD].tolist()
    # bank_df[MATCHED_MEMBER_FIELD] = bank_df[DESC_NAME_FIELD].astype(str).apply(lambda x: process.extractOne(x, member_names, scorer=fuzz.partial_ratio)[0] if isinstance(x, str) else None)
    # bank_df[MATCHED_MEMBER_FIELD] = bank_df[DESC_NAME_FIELD].astype(str).apply(match_member, member_names=member_names)
    # Extract the first name and last name columns from the members_df into a new df
    members_name_df = members_df[[FIRST_NAME_FIELD, LAST_NAME_FIELD, MEMBERS_ID_FIELD]].copy()
    bank_df[MATCHED_MEMBER_FIELD] = bank_df[DESC_NAME_FIELD].astype(str).apply(lambda x: manual_match_member(x, members_df[FIRST_NAME_FIELD].tolist(), members_df[LAST_NAME_FIELD].tolist(), members_name_df) if isinstance(x, str) else None)
    # bank_df[MATCHED_MEMBER_ID_FIELD] = bank_df[MATCHED_MEMBER_FIELD].apply(lambda x: members_df[members_df[FULL_NAME_FIELD] == x]['ID'].values[0] if x in members_df[FULL_NAME_FIELD].values else None)
    return bank_df
    #  Write the processed bank_df to an Excel file
    # bank_df.to_excel("processed_bank_statement.xlsx", index=False)

# Implement extract apply functions to match members
def match_member(description: str, member_names: list) -> str:
    """Match the description with member names using fuzzy matching."""
    if pd.isnull(description):
        return None
    # best_match, score, _ = process.extractOne(description, member_names, scorer=fuzz.partial_ratio)
    best_match = process.extract(description, member_names, scorer=fuzz.token_sort_ratio, limit=3)
    print(f"Description: {description}, Best match: {best_match}, ")
    # return best_match if score >= 80 else None 
    return best_match[0][0] if len(best_match) > 0 else None

def manual_match_member(description: str, first_names: list, last_names: list, members_df: pd.DataFrame) -> str:
    bank_names = description.split()
    if len(bank_names) == 2:
        matched_rows = members_df[members_df[LAST_NAME_FIELD].str.lower() ==  bank_names[1].lower()]
        if len(matched_rows) == 1:
            # If we have a single match, check if the first name matches
            first_name = bank_names[0]
            f_name = matched_rows[FIRST_NAME_FIELD].values[0]
            #  Get the first and last name concatenated with space from the matched rows
            full_name = f"{f_name} {matched_rows[LAST_NAME_FIELD].values[0]}"
            if f_name.startswith(first_name[0]):
                print(f"Description: {description}, Matched first name: {f_name}, Last name: {matched_rows[LAST_NAME_FIELD].values[0]}")
                return   full_name
            return full_name+ " (Check)"
        elif len(matched_rows) > 1:
            print(f"Description: {description}, Matched rows: {matched_rows}, but multiple last names match.")
            # If we have multiple matches, we need to check the first name
            first_name = bank_names[0]
            matched_first_name = members_df[members_df[FIRST_NAME_FIELD].str.lower().str.startswith(first_name.lower())]
            if matched_first_name.empty:
                print(f"Description: {description}, No first name match found.")
                return None
            if len(matched_first_name) == 1:
                # If we have a single match, return the full name
                full_name = f"{matched_first_name[FIRST_NAME_FIELD].values[0]} {matched_first_name[LAST_NAME_FIELD].values[0]}"
                print(f"Description: {description}, Matched first name: {matched_first_name[FIRST_NAME_FIELD].values[0]}, Last name: {matched_first_name[LAST_NAME_FIELD].values[0]}")
                return full_name
            else:
                print(f"Description: {description}, Multiple matches found, but first name does not match.")
                return None
    if len(bank_names) == 2:
        if len(bank_names[0]) == 1:
            #  If the first name is just a single character, we can assume it's a first name
            last_name = bank_names[1]
            # find all last names that match the last name in the bank description
            matched_last_names = [name for name in last_names if name.lower() == last_name.lower()]
            if len(matched_last_names) == 1:
                # Check if the first name starts with the same letter as our first name
                first_name = bank_names[0]
                if matched_last_names[0].startswith(first_name[0]):
                    return matched_last_names[0]
            return None
                
            print(f"Description: {description}, Matched last names: {matched_last_names}")
    #     last_name_match0 = process.extractOne(bank_names[0], last_names, scorer=fuzz.token_set_ratio)
    #     last_name_match1 = process.extractOne(bank_names[1], last_names, scorer=fuzz.token_set_ratio)
    #     print(f"Description: {description}, Last name matches: {last_name_match0}, {last_name_match1}")

    return process.extractOne(description, first_names, scorer=fuzz.partial_ratio)[0] if isinstance(description, str) else None


def load_and_clean_statement(eoy_file: str) -> pd.DataFrame:
    """Load and process the bank statement from the given file."""
    print_progress("Loading bank statement...")
    bank_df = pd.read_excel(eoy_file)
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

    print_progress(f"Loaded bank statement data shape: {bank_df.shape}")
    return bank_df


if __name__ == "__main__":
    bank_df = process_data()
    print_progress(f"Processed bank statement data shape: {bank_df.shape}")
    #  Write the processed bank_df to an Excel file
    bank_df.to_excel("processed_bank_statement.xlsx", index=False)
    print("Data processing complete. Processed data saved to 'processed_bank_statement.xlsx'.")
else:
    print("This script is intended to be run as a standalone program.")
    # If imported, the process_data function can be called directly.
