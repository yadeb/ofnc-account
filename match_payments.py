import pandas as pd
from rapidfuzz import fuzz, process

# Load data
bank_df = pd.read_excel("bank_statement.xlsx")
members_df = pd.read_excel("members_list.xlsx")
bank_df.dropna(subset=['Description', 'Amount'], inplace=True)

# Print loaded data shapes
print(f"Bank Data Shape: {bank_df.shape}")
print(f"Members Data Shape: {members_df.shape}")

# Create full name column for matching
members_df['Full Name'] = members_df['First Name'].str.strip() + " " + members_df['Last Name'].str.strip()
member_names = members_df['Full Name'].tolist()

def process_data():
    # Load the bank statement and members list
    bank_df = pd.read_excel("bank_statement.xlsx")
    members_df = pd.read_excel("members_list.xlsx")

    # Drop rows with NaN in Description or Amount
    bank_df.dropna(subset=['Description', 'Amount'], inplace=True)

    # Drop all rows where Description is just a number
    bank_df = bank_df[~bank_df['Description'].astype(str).str.isnumeric()]

    # From the bank description extract the first two words into a Description_Name column
    bank_df['Description_Name'] = bank_df['Description'].astype(str).apply(lambda x: ' '.join(x.split()[:2]) if isinstance(x, str) else x)

    # Sum amounts by Description_Name
    bank_df = bank_df.groupby('Description_Name', as_index=False)['Amount'].sum()

    #  Clean up the memebers list drop duplicate rowas with the same First Name and Last Name, keep the last occurrence with the most recent Start Time column
    members_df.drop_duplicates(subset=['First Name', 'Last Name'], keep='last', inplace=True)

    # Match the Description_Name with member names in the members list and create a new column Matched Member with the matched name and the ID from the members list
    members_df['Full Name'] = members_df['First Name'].str.strip() + " " + members_df['Last Name'].str.strip()
    member_names = members_df['Full Name'].tolist()
    bank_df['Matched Member'] = bank_df['Description_Name'].astype(str).apply(lambda x: process.extractOne(x, member_names, scorer=fuzz.partial_ratio)[0] if isinstance(x, str) else None)

    bank_df['Matched Member ID'] = bank_df['Matched Member'].apply(lambda x: members_df[members_df['Full Name'] == x]['ID'].values[0] if x in members_df['Full Name'].values else None)

    #  Write the processed bank_df to an Excel file
    bank_df.to_excel("processed_bank_statement.xlsx", index=False)



def match_member(description):
    # print(f"Matching description: {description}")
    if pd.isnull(description):
        return None
    best_match, score, _ = process.extractOne(description, member_names, scorer=fuzz.partial_ratio)
    # print(f"Best match: {best_match}, Score: {score}")
    return best_match if score >= 80 else None

def get_names(description):
    """Extract last name from description."""
    if pd.isnull(description):
        return None

    parts = str(description).strip()
 
    parts = parts.rsplit('   ', 1)
    if len(parts) == 1:
        return parts[0].strip(), ""
    
    if len(parts) == 2:
        parts = parts[0].strip().split(' ')
    if parts[-1].isdigit():
        #  Drop the last 2 parts if there are more than 6 parts
        parts = parts[:-1]
        # if the last part contains a digit, it is likely a date or amount
        if any(part[-1].isdigit() for part in parts[-1]):
            parts = parts[:-1]

    
    print(f"Parts after split: {parts}")
    if len(parts) == 1:
        return parts[0].strip(), ""
    
    if len(parts[0]) == 1:
        return parts[1].strip(), parts[0].strip()
    
    if len(parts[1]) == 1 and len(parts) > 3:
        if parts[1].strip() in ['&', '+']:
            initial = parts[0].strip() + parts[1].strip() + parts[2].strip()
            return parts[3].strip(), initial

    return parts[1].strip(), parts[0].strip()


# Apply matching function
# Get the first two space separated words from the Description column
bank_df['Description_Name'] = bank_df['Description'].astype(str).apply(lambda x: ' '.join(x.split()[:2]) if isinstance(x, str) else x)

# Extract last name and initial from the description
bank_df['Last Name'], bank_df['Initial'] = zip(*bank_df['Description'].apply(get_names))
# Print the
print(bank_df[['Description_Name', 'Last Name', 'Initial']])

# Match members
bank_df['Matched Member'] = bank_df['Description_Name'].astype(str).apply(match_member)
# Print matching results


# Convert Amount column to numeric if needed
bank_df['Amount'] = pd.to_numeric(bank_df['Amount'], errors='coerce')

# Group and sum
summary_df = bank_df.groupby('Matched Member', dropna=True)['Amount'].sum().reset_index()

# Save output
summary_df.to_excel("member_payment_summary.xlsx", index=False)
print("✅ Payment matching complete. Results saved to 'member_payment_summary.xlsx'")