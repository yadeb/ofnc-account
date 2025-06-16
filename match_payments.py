import pandas as pd
from rapidfuzz import fuzz, process

# Load data
bank_df = pd.read_excel("bank_statement.xlsx")
members_df = pd.read_excel("members_list.xlsx")

# Print loaded data shapes
print(f"Bank Data Shape: {bank_df.shape}")
print(f"Members Data Shape: {members_df.shape}")

# Create full name column for matching
members_df['Full Name'] = members_df['First Name'].str.strip() + " " + members_df['Last Name'].str.strip()
member_names = members_df['Full Name'].tolist()

def match_member(description):
    print(f"Matching description: {description}")
    if pd.isnull(description):
        return None
    best_match, score, _ = process.extractOne(description, member_names, scorer=fuzz.partial_ratio)
    print(f"Best match: {best_match}, Score: {score}")
    return best_match if score >= 80 else None

# Apply matching function
bank_df['Matched Member'] = bank_df['Description'].astype(str).apply(match_member)
# Print matching results


# Convert Amount column to numeric if needed
bank_df['Amount'] = pd.to_numeric(bank_df['Amount'], errors='coerce')

# Group and sum
summary_df = bank_df.groupby('Matched Member', dropna=True)['Amount'].sum().reset_index()

# Save output
summary_df.to_excel("member_payment_summary.xlsx", index=False)
print("✅ Payment matching complete. Results saved to 'member_payment_summary.xlsx'")