# Add the necessary imports
import pandas as pd
from rapidfuzz import fuzz, process
from income_data_loader import IncomeDataLoader

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
DROP_NAMES_LIST = ["NIGHT SAFE", "STWDSHP"]


def print_progress(message):
    """Prints a progress message to the console."""
    print(f"[INFO] {message}")


def process_data() -> pd.DataFrame:
    # Load the bank statement and members list
    print_progress("Loading bank statement and members list...")
    bank_df = load_and_clean_statement("bank_statement.xlsx")
    print_progress(f"Process Data: Loaded bank statement data shape: {bank_df.shape}")
 
    # Load the members list from an Excel file
    members_df = pd.read_excel("members_list.xlsx")
    print_progress(f"Process Data: Loaded members list data shape: {members_df.shape}")

    #  Normalize the First Name and Last Name columns in the members list
    members_df[FIRST_NAME_FIELD] = members_df[FIRST_NAME_FIELD].str.strip().str.title()
    members_df[LAST_NAME_FIELD] = members_df[LAST_NAME_FIELD].str.strip().str.title()
    print_progress(
        "Process Data: Normalized First Name and Last Name columns in the members list."
    )

    # Clean up the members list by dropping duplicate rows with the same First Name and Last Name, keeping the last occurrence
    members_df.drop_duplicates(
        subset=[FIRST_NAME_FIELD, LAST_NAME_FIELD], keep="last", inplace=True
    )
    print_progress(
        f"Process Data: Members list data shape after dropping duplicates: {members_df.shape}"
    )

    # Match the Description_Name with member names in the members list and create a new column Matched Member with the matched name and the ID from the members list
    members_df[FULL_NAME_FIELD] = (
        members_df[FIRST_NAME_FIELD].str.strip()
        + " "
        + members_df[LAST_NAME_FIELD].str.strip()
    )

    tmp_bank_df = match_payee_to_members(bank_df, members_df)
    bank_df[MATCHED_MEMBER_FIELD] = tmp_bank_df[MATCHED_MEMBER_FIELD]
    return bank_df


def manual_match_member(
    description: str, first_names: list, last_names: list, members_df: pd.DataFrame
) -> str:
    bank_names = description.split()
    if len(bank_names) == 2:
        matched_rows = members_df[
            members_df[LAST_NAME_FIELD].str.lower() == bank_names[1].lower()
        ]
        if len(matched_rows) == 1:
            # If we have a single match, check if the first name matches
            first_name = bank_names[0]
            f_name = matched_rows[FIRST_NAME_FIELD].values[0]
            #  Get the first and last name concatenated with space from the matched rows
            full_name = f"{matched_rows[FIRST_NAME_FIELD].values[0]} {matched_rows[LAST_NAME_FIELD].values[0]}"
            if f_name.startswith(first_name[0]):
                print(
                    f"Description1: {description}, Matched first name: {f_name}, Last name: {matched_rows[LAST_NAME_FIELD].values[0]}"
                )
                return full_name
            return full_name + " (Check)"
        elif len(matched_rows) > 1:
            print(
                f"Description: {description}, Matched rows: {matched_rows}, but multiple last names match."
            )
            # If we have multiple matches, we need to check the first name
            first_name = bank_names[0]
            matched_first_name = members_df[
                members_df[FIRST_NAME_FIELD]
                .str.lower()
                .str.startswith(first_name.lower())
            ]
            if matched_first_name.empty:
                print(f"Description: {description}, No first name match found.")
                return None
            if len(matched_first_name) == 1:
                # If we have a single match, return the full name
                full_name = f"{matched_first_name[FIRST_NAME_FIELD].values[0]} {matched_first_name[LAST_NAME_FIELD].values[0]}"
                print(
                    f"Description: {description}, Matched first name: {matched_first_name[FIRST_NAME_FIELD].values[0]}, Last name: {matched_first_name[LAST_NAME_FIELD].values[0]}"
                )
                return full_name
            else:
                print(
                    f"Description: {description}, Multiple matches found, but first name does not match."
                )
                return None
    if len(bank_names) == 2:
        if len(bank_names[0]) == 1:
            #  If the first name is just a single character, we can assume it's a first name
            last_name = bank_names[1]
            # find all last names that match the last name in the bank description
            matched_last_names = [
                name for name in last_names if name.lower() == last_name.lower()
            ]
            if len(matched_last_names) == 1:
                # Check if the first name starts with the same letter as our first name
                first_name = bank_names[0]
                if matched_last_names[0].startswith(first_name[0]):
                    return matched_last_names[0]
            return None

            print(
                f"Description: {description}, Matched last names: {matched_last_names}"
            )

    return (
        process.extractOne(description, first_names, scorer=fuzz.partial_ratio)[0]
        if isinstance(description, str)
        else None
    )


def load_and_clean_statement(eoy_file: str) -> pd.DataFrame:
    """Load and process the bank statement from the given file."""
    print_progress("Loading bank statement...")
    bank_df = pd.read_excel(eoy_file)
    print_progress(f"Loaded bank statement data shape: {bank_df.shape}")

    # mark the description field as a string to avoid issues with mixed types
    bank_df[DESCRIPTION_FIELD] = bank_df[DESCRIPTION_FIELD].astype(str)

    print_progress("Cleaning and processing bank statement data...")
    #  Rename "Business a/c" to "Amount"
    bank_df.rename(columns={"Business a/c": AMOUNT_FIELD}, inplace=True)

    # Drop rows with NaN in Description or Amount
    bank_df.dropna(subset=[DESCRIPTION_FIELD, AMOUNT_FIELD], inplace=True)
    print_progress(
        f"Bank statement data shape after dropping NaN in Description or Amount: {bank_df.shape}"
    )

    #  Drop rows where the Description contains any of the names in the DROP_NAMES_LIST
    for name in DROP_NAMES_LIST:
        bank_df = bank_df[
            ~bank_df[DESCRIPTION_FIELD].str.contains(name, case=False, na=False)
        ]
        print_progress(
            f"Bank statement data shape after dropping rows with '{name}' in Description: {bank_df.shape}"
        )

    # Drop rows if the Purpose colomn does not contain the word 'Tithe' or 'Offering'
    bank_df = bank_df[
        bank_df[PURPOSE_FIELD].str.contains("Tithe", case=False, na=False)
        | bank_df[PURPOSE_FIELD].str.contains("Offering", case=False, na=False)
    ]
    print_progress(
        f"Filtered bank statement data shape after dropping non Tithes and Offering: {bank_df.shape}"
    )
    #  Drop Restricted rows if the Restricted field exists
    if RESTRICTED_FIELD in bank_df.columns.to_list():
        # If the Restricted field exists, filter out rows where it is True
        bank_df = bank_df[
            ~bank_df[RESTRICTED_FIELD].str.contains("Yes", case=False, na=False)
        ]
    print_progress(
        f"Filtered bank statement data shape after removing restricted rows: {bank_df.shape}"
    )

    # Drop all rows where Description is just a number
    bank_df = bank_df[
        ~pd.to_numeric(bank_df[DESCRIPTION_FIELD], errors="coerce").notnull()
    ]
    print_progress(
        f"Bank statement data shape after dropping rows with numeric Description: {bank_df.shape}"
    )

    # # From the bank description extract the first two words into a Description_Name column
    # bank_df[DESC_NAME_FIELD] = (
    #     bank_df[DESCRIPTION_FIELD]
    #     .astype(str)
    #     .apply(lambda x: " ".join(x.split()[:2]) if isinstance(x, str) else x)
    # )

    bank_df = extract_payee_name(bank_df)
    print_progress("Extracted Description_Name from bank statement data.")

    # Sum amounts by Description_Name
    bank_df = bank_df.groupby(DESC_NAME_FIELD, as_index=False)[AMOUNT_FIELD].sum()
    print_progress(
        "Grouped bank statement data by Description_Name and summed amounts."
    )

    print_progress(f"Loaded bank statement data shape: {bank_df.shape}")
    return bank_df


import pandas as pd
from rapidfuzz import process, fuzz


def match_payee_to_members(
    bank_df: pd.DataFrame, members_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Matches the payee_name in the bank transactions DataFrame to the members list DataFrame.

    Args:
        bank_df (pd.DataFrame): DataFrame containing the payee_name column.
        members_df (pd.DataFrame): DataFrame containing ID, first name, and last name columns.

    Returns:
        pd.DataFrame: Updated bank_df with matched full name and ID columns.
    """
    # Create a full name column in the members list
    members_df[FULL_NAME_FIELD] = (
        members_df[FIRST_NAME_FIELD].str.strip()
        + " "
        + members_df[LAST_NAME_FIELD].str.strip()
    )
    members_df[FULL_NAME_FIELD] = members_df[FULL_NAME_FIELD].str.lower()

    # Convert the full name column to a list for matching
    member_names = members_df[FULL_NAME_FIELD].tolist()

    # Define a function to match payee_name to the members list
    def match_member(payee_name: str) -> tuple:
        if pd.isnull(payee_name):
            return None, None
        best_match, score, _ = process.extractOne(
            payee_name.lower(), member_names, scorer=fuzz.partial_token_sort_ratio
        )
        if score >= 80:  # Set a threshold for matching
            # print_progress(
            #     f"Matching payee_name: {payee_name}, Best match: {best_match.title()}, Score: {score}"
            # )

            matched_id = members_df.loc[
                members_df[FULL_NAME_FIELD] == best_match, MEMBERS_ID_FIELD
            ].values[0]
            #  Convert the best match to title case for consistency
            best_match = best_match.title()
            return best_match, matched_id
        else:
            # print_progress(
            #     f"Matching payee_name: {payee_name}, No match found or score below threshold."
            # )
            # If no match is found, try manual matching
            return manual_match_member(payee_name)

    #  Try manual matching for failed fuzzy matches
    def manual_match_member(payee_name: str) -> tuple:
        if pd.isnull(payee_name):
            return None, None
        # Split the payee_name into words
        bank_names = payee_name.split()
        if len(bank_names) == 2:
            #  if the first name is just a single character, we can assume it's a first name
            if len(bank_names[0]) == 1:
                inital = bank_names[0]
                last_name = bank_names[1]
                # find all last names that match the last name in the bank description
                matched_rows = members_df[
                    members_df[LAST_NAME_FIELD].str.lower() == last_name.lower()
                ]
                if len(matched_rows) == 1:
                    # If we have a single match, check if the first name matches
                    f_name = matched_rows[FIRST_NAME_FIELD].values[0]
                    #  Get the first and last name concatenated with space from the matched rows
                    full_name = f"{matched_rows[FIRST_NAME_FIELD].values[0]} {matched_rows[LAST_NAME_FIELD].values[0]}"
                    if f_name.lower().startswith(inital[0].lower()):
                        print(
                            f"Matching payee_name: {payee_name}, Matched first name: {f_name}, Last name: {matched_rows[LAST_NAME_FIELD].values[0]}"
                        )
                        return full_name, matched_rows[MEMBERS_ID_FIELD].values[0]
                    return (
                        full_name + " (Check)",
                        matched_rows[MEMBERS_ID_FIELD].values[0],
                    )
                elif len(matched_rows) > 1:
                    print(
                        f"Matching payee_name: {payee_name}, Matched rows: {matched_rows}, but multiple last names match."
                    )
                    # If we have multiple matches, we need to check the first name
                    matched_first_name = members_df[
                        members_df[FIRST_NAME_FIELD]
                        .str.lower()
                        .str.startswith(inital.lower())
                    ]
                    if matched_first_name.empty:
                        print(
                            f"Matching payee_name: {payee_name}, No first name match found."
                        )
                        return None, None
                    if len(matched_first_name) == 1:
                        # If we have a single match, return the full name
                        full_name = f"{matched_first_name[FIRST_NAME_FIELD].values[0]} {matched_first_name[LAST_NAME_FIELD].values[0]}"
                        print(
                            f"Matching payee_name: {payee_name}, Matched first name: {matched_first_name[FIRST_NAME_FIELD].values[0]}, Last name: {matched_first_name[LAST_NAME_FIELD].values[0]}"
                        )
                        return full_name, matched_first_name[MEMBERS_ID_FIELD].values[0]
                    else:
                        print(
                            f"Matching payee_name: {payee_name}, Multiple matches found, but first name does not match."
                        )
                        # return the first match as a fallback
                        full_name = f"{matched_first_name[FIRST_NAME_FIELD].values[0]} {matched_first_name[LAST_NAME_FIELD].values[0]} (Check)"
                        return full_name, matched_first_name[MEMBERS_ID_FIELD].values[0]

        return None, None

    # Apply the matching function to the payee_name column
    bank_df[MATCHED_MEMBER_FIELD], bank_df[MATCHED_MEMBER_ID_FIELD] = zip(
        *bank_df[DESC_NAME_FIELD].apply(match_member)
    )
    #  Prunt how many matches were made
    matched_count = bank_df[MATCHED_MEMBER_FIELD].notnull().sum()
    print_progress(f"Matched {matched_count} members in the bank statement data.")

    return bank_df



def extract_payee_name(df: pd.DataFrame) -> pd.DataFrame:
    def parse_description(desc: str) -> str:
        parts = desc.split()
        
        # Handle FPI: likely long transaction with structure at the end
        if len(parts) >= 7:
            # Assume last 5 words are transaction details, reference is before that (max 18 chars)
            main = parts[:-5]
            #  Name will always be at least the first two words
            # Try to find the reference (max 18 chars, can be multiple words)
            # We'll assume the reference is the last word before the 5 details if it's <= 18 chars
            # Name is everything before the reference
            for i in range(len(main)-1, 1, -1):
                ref_candidate = ' '.join(main[i:])
                if len(ref_candidate) == 18:
                    return ' '.join(main[:i])
                if len(ref_candidate) > 18:
                    return ' '.join(main[:i +1])
            
            name = ' '.join(main[:2])  # Fallback to first two words if no reference found
            if len(main) > 3 and main[1] in ['&', '+']:
                # If the third word is '&' or 'and', we can include it in the name
                #  Add the third word and fourth word to the name
                name = name + ' ' + main[2] + ' ' + main[3]
                return name
            if len(main) > 3 and main[2] in ['&', '+']: 
                # If the fourth word is '&' or 'and', it is a joint account with the last name first, get the second initial
                name = name + ' ' + main[3]

            return name # Fallback to first two words if we get here
        elif len(parts) >= 2:
            # Handle SO: name followed by a reference
            return " ".join(parts[:2])  # At least the first two words as name
        elif len(parts) == 1:
            # Possibly a cheque deposit or one-word name
            return parts[0]
        
        return desc  # Fallback to full description if pattern not matched

    df[DESC_NAME_FIELD] = df[DESCRIPTION_FIELD].apply(parse_description)
    return df

def load_consolidated_data(file_path: str, income_headers: list) -> pd.DataFrame:
    """
    Load and process consolidated income data from an Excel file.

    Args:
        file_path (str): Path to the Excel file.
        income_headers (list): List of headers that identify the title row of the income data.

    Returns:
        pd.DataFrame: Processed DataFrame containing consolidated income data.
    """
    # main.py
    
    # Define the headers to look for in the title row
    income_headers = ['Date',	'Branch',	'Transaction Description',	'Purpose',	'Description',	'Business a/c']

    # Create an instance of the loader
    loader = IncomeDataLoader(file_path, income_headers)

    # Load the data
    loader.load_income_data()

    # Get the DataFrames
    dataframes = loader.get_dataframes()
    print_progress(f"Loaded {len(dataframes)} income data sheets from {file_path}.")
    # print the names of the loaded DataFrames
    for name in dataframes.keys():
        print(f"Loaded DataFrame: {name}")

    # Example: access the Bedford sheet data
    bedford_df = dataframes.get("Bedford_df")

    if bedford_df is not None:
        print("Bedford DataFrame:")
        print(bedford_df.head())
    else:
        print("No income data found for 'Bedford'.")

if __name__ == "__main__":
    load_consolidated_data("consolidated_income_data_2024.xlsx", income_headers=None)
    print_progress("Consolidated income data loaded and processed.")
    exit(0)
    bank_df = process_data()
    print_progress(f"Processed bank statement data shape: {bank_df.shape}")
    #  Write the processed bank_df to an Excel file
    bank_df.to_excel("processed_bank_statement.xlsx", index=False)
    print(
        "Data processing complete. Processed data saved to 'processed_bank_statement.xlsx'."
    )
else:
    print("This script is intended to be run as a standalone program.")
    # If imported, the process_data function can be called directly.
