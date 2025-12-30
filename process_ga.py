import os
import pandas as pd
from rapidfuzz import fuzz, process
from income_data_loader import IncomeDataLoader
import argparse

class GAProcessor:
    # Constants
    RESTRICTED_FIELD = "Restricted"
    DESCRIPTION_FIELD = "Transaction Description"
    PURPOSE_FIELD = "Purpose"
    AMOUNT_FIELD = "Amount"
    DATE_FIELD = "Date"
    DESC_NAME_FIELD = "TransactionDesc_Name"
    FIRST_NAME_FIELD = "First Name"
    LAST_NAME_FIELD = "Last Name"
    FULL_NAME_FIELD = "Full Name"
    MATCHED_MEMBER_FIELD = "Matched Member"
    MEMBERS_ID_FIELD = "ID"
    DROP_NAMES_LIST = ["NIGHT SAFE", "STWDSHP"]
    ALLOWED_PURPOSES = ["Tithe", "Offering", "Donation to Charity"]
    BRANCH_NAME_FIELD = "OFNC Branch"
    PREVIOUSLY_MATCHED_FIELD = "Previously Matched"
    SURNAME_FIELD = "Surname"
    HOUSE_NUMBER_FIELD = "House Number"
    FULL_HOUSE_ADDRESS_FIELD = "Full House Address"
    POSTCODE_FIELD = "Postcode"
    SPONSORED_EVENT_FIELD = "Sponsored Event"
    TITLE_FIELD = "Title"
    AGGREGATED_DONATION_FIELD = "Aggregated Donation (Leave Blank)"
    SPONSORED_EVENT_FIELD2 = "Sponsored Event (Yes/Blank)"

    def __init__(self, folder=".", account_file="ConsolidatedAccounts2024Final1_GA.xlsx", consent_file="ga_consent_list.xlsx"):
        self.folder = folder
        self.account_file = account_file
        self.consent_file = consent_file
        self.account_file_path = os.path.join(self.folder, self.account_file)
        self.consent_file_path = os.path.join(self.folder, self.consent_file)

    @staticmethod
    def parse_args():
        parser = argparse.ArgumentParser(description="Process GA consolidated income data.")
        parser.add_argument("--folder", type=str, default=".", help="Path to the folder containing the Excel files (default: current directory)")
        parser.add_argument("--account_file", type=str, default="ConsolidatedAccounts2024Final1_GA.xlsx", help="Name of the consolidated account Excel file")
        parser.add_argument("--consent_file", type=str, default="ga_consent_list.xlsx", help="Name of the GA consent Excel file")
        args = parser.parse_args()
        return args.folder, args.account_file, args.consent_file

    def print_progress(self, message):
        print(f"[INFO] {message}")

    def manual_match_member(self, description, first_names, last_names, members_df):
        bank_names = description.split()
        if len(bank_names) == 2:
            matched_rows = members_df[
                members_df[self.LAST_NAME_FIELD].str.lower() == bank_names[1].lower()
            ]
            if len(matched_rows) == 1:
                first_name = bank_names[0]
                f_name = matched_rows[self.FIRST_NAME_FIELD].values[0]
                full_name = f"{matched_rows[self.FIRST_NAME_FIELD].values[0]} {matched_rows[self.LAST_NAME_FIELD].values[0]}"
                if f_name.startswith(first_name[0]):
                    print(f"Description1: {description}, Matched first name: {f_name}, Last name: {matched_rows[self.LAST_NAME_FIELD].values[0]}")
                    return full_name
                return full_name + " (Check)"
            elif len(matched_rows) > 1:
                print(f"Description: {description}, Matched rows: {matched_rows}, but multiple last names match.")
                first_name = bank_names[0]
                matched_first_name = members_df[
                    members_df[self.FIRST_NAME_FIELD]
                    .str.lower()
                    .str.startswith(first_name.lower())
                ]
                if matched_first_name.empty:
                    print(f"Description: {description}, No first name match found.")
                    return None
                if len(matched_first_name) == 1:
                    full_name = f"{matched_first_name[self.FIRST_NAME_FIELD].values[0]} {matched_first_name[self.LAST_NAME_FIELD].values[0]}"
                    print(f"Description: {description}, Matched first name: {matched_first_name[self.FIRST_NAME_FIELD].values[0]}, Last name: {matched_first_name[self.LAST_NAME_FIELD].values[0]}")
                    return full_name
                else:
                    print(f"Description: {description}, Multiple matches found, but first name does not match.")
                    return None
        if len(bank_names) == 2:
            if len(bank_names[0]) == 1:
                last_name = bank_names[1]
                matched_last_names = [
                    name for name in last_names if name.lower() == last_name.lower()
                ]
                if len(matched_last_names) == 1:
                    first_name = bank_names[0]
                    if matched_last_names[0].startswith(first_name[0]):
                        return matched_last_names[0]
                return None
        return (
            process.extractOne(description, first_names, scorer=fuzz.partial_ratio)[0]
            if isinstance(description, str)
            else None
        )

    def load_and_clean_statement(self, bank_df):
        bank_df[self.DESCRIPTION_FIELD] = bank_df[self.DESCRIPTION_FIELD].astype(str)
        self.print_progress("Cleaning and processing bank statement data...")
        bank_df.rename(columns={"Business a/c": self.AMOUNT_FIELD}, inplace=True)
        bank_df.dropna(subset=[self.DESCRIPTION_FIELD, self.AMOUNT_FIELD], inplace=True)
        self.print_progress(f"Bank statement data shape after dropping NaN in Description or Amount: {bank_df.shape}")
        self.print_progress(f"Bank statement columns: {bank_df.columns.tolist()}")
        for name in self.DROP_NAMES_LIST:
            bank_df = bank_df[
                ~bank_df[self.DESCRIPTION_FIELD].str.contains(name, case=False, na=False)
            ]
            self.print_progress(f"Bank statement data shape after dropping rows with '{name}' in Description: {bank_df.shape}")
        if self.PURPOSE_FIELD not in bank_df.columns:
            self.print_progress(f"Purpose field '{self.PURPOSE_FIELD}' not found in bank statement data. Skipping filtering by Purpose.")
        else:
            bank_df = bank_df[
                bank_df[self.PURPOSE_FIELD].str.contains(
                    "|".join(self.ALLOWED_PURPOSES), case=False, na=False
                )
            ]
            self.print_progress(f"Filtered bank statement data shape after keeping only Tithes, Offerings, and Donations: {bank_df.shape}")
        if self.RESTRICTED_FIELD in bank_df.columns.to_list():
            bank_df = bank_df[
                ~bank_df[self.RESTRICTED_FIELD].str.contains("Yes", case=False, na=False)
            ]
            self.print_progress(f"Filtered bank statement data shape after removing restricted rows: {bank_df.shape}")
        bank_df = bank_df[
            ~pd.to_numeric(bank_df[self.DESCRIPTION_FIELD], errors="coerce").notnull()
        ]
        self.print_progress(f"Bank statement data shape after dropping rows with numeric Description: {bank_df.shape}")
        bank_df = self.extract_payee_name(bank_df)
        bank_df[self.DATE_FIELD] = pd.to_datetime(bank_df[self.DATE_FIELD], errors='coerce').dt.strftime('%d/%m/%Y')
        bank_df = bank_df.groupby(self.DESC_NAME_FIELD).agg(
            Amount=(self.AMOUNT_FIELD, 'sum'),
            Date=(self.DATE_FIELD, 'max'),
            Transaction_Description=(self.DESCRIPTION_FIELD, 'last'),
        ).reset_index()
        self.print_progress(f"Loaded bank statement data shape: {bank_df.shape}")
        return bank_df

    def match_payee_to_members(self, bank_df, members_df_in):
        members_df = members_df_in.copy()
        members_df.loc[:, self.FULL_NAME_FIELD] = (
            members_df[self.FIRST_NAME_FIELD].str.strip()
            + " "
            + members_df[self.LAST_NAME_FIELD].str.strip()
        )
        members_df.loc[:, self.FULL_NAME_FIELD] = members_df[self.FULL_NAME_FIELD].str.lower()
        member_names = members_df[self.FULL_NAME_FIELD].tolist()

        def match_member(payee_name):
            if pd.isnull(payee_name):
                return None, None
            best_match, score, _ = process.extractOne(
                payee_name.lower(), member_names, scorer=fuzz.partial_token_sort_ratio
            )
            if score >= 80:
                matched_id = members_df.loc[
                    members_df[self.FULL_NAME_FIELD] == best_match, self.MEMBERS_ID_FIELD
                ].values[0]
                best_match_title = best_match.title()
                return best_match_title, matched_id
            else:
                return self.manual_match_member(payee_name, [], [], members_df), None

        merged_df = bank_df.copy()
        if not members_df.empty:
            bank_df[self.MATCHED_MEMBER_FIELD], bank_df[self.MEMBERS_ID_FIELD] = zip(
                *bank_df[self.DESC_NAME_FIELD].apply(match_member)
            )
            merged_df = bank_df.merge(
                members_df[
                    [self.MEMBERS_ID_FIELD, self.TITLE_FIELD, self.FIRST_NAME_FIELD, self.LAST_NAME_FIELD, self.FULL_HOUSE_ADDRESS_FIELD, self.POSTCODE_FIELD]
                ],
                on=self.MEMBERS_ID_FIELD, how='left'
            )
            merged_df = merged_df[[
                "Transaction_Description", self.DESC_NAME_FIELD, self.MATCHED_MEMBER_FIELD, self.TITLE_FIELD,
                self.FIRST_NAME_FIELD, self.LAST_NAME_FIELD, self.FULL_HOUSE_ADDRESS_FIELD, self.POSTCODE_FIELD,
                self.AMOUNT_FIELD, self.DATE_FIELD
            ]]
            merged_df[self.HOUSE_NUMBER_FIELD] = merged_df[self.FULL_HOUSE_ADDRESS_FIELD].fillna('').str.strip().str.split().str[0]
            merged_df.drop(columns=[self.FULL_HOUSE_ADDRESS_FIELD], inplace=True)
        else:
            merged_df = bank_df.copy()
            ga_fields = [self.MATCHED_MEMBER_FIELD, self.TITLE_FIELD, self.FIRST_NAME_FIELD, self.LAST_NAME_FIELD, self.HOUSE_NUMBER_FIELD, self.POSTCODE_FIELD]
            for field in ga_fields:
                merged_df[field] = ''
            cols = list(merged_df.columns)
            cols.insert(0, cols.pop(cols.index("Transaction_Description")))
            merged_df = merged_df[cols]
        cols = list(merged_df.columns)
        last_name_index = cols.index(self.LAST_NAME_FIELD)
        cols.insert(last_name_index + 1, cols.pop(cols.index(self.HOUSE_NUMBER_FIELD)))
        merged_df = merged_df[cols]
        merged_df.insert(
            merged_df.columns.get_loc(self.POSTCODE_FIELD) + 1,
            self.AGGREGATED_DONATION_FIELD,
            ''
        )
        merged_df.insert(
            merged_df.columns.get_loc(self.AGGREGATED_DONATION_FIELD) + 1,
            self.SPONSORED_EVENT_FIELD2,
            ''
        )
        matched_count = merged_df[self.MATCHED_MEMBER_FIELD].notnull().sum()
        self.print_progress(f"Matched {matched_count} members in the bank statement data.")
        return merged_df

    def extract_payee_name(self, df):
        def parse_description(desc):
            parts = desc.split()
            if len(parts) >= 7:
                main = parts[:-5]
                for i in range(len(main) - 1, 1, -1):
                    ref_candidate = " ".join(main[i:])
                    if len(ref_candidate) == 18:
                        return " ".join(main[:i])
                    if len(ref_candidate) > 18:
                        return " ".join(main[: i + 1])
                name = " ".join(main[:2])
                if len(main) > 3 and main[1] in ["&", "+"]:
                    name = name + " " + main[2] + " " + main[3]
                    return name
                if len(main) > 3 and main[2] in ["&", "+"]:
                    name = name + " " + main[3]
                return name
            elif len(parts) >= 2:
                return " ".join(parts[:2])
            elif len(parts) == 1:
                return parts[0]
            return desc
        df[self.DESC_NAME_FIELD] = df[self.DESCRIPTION_FIELD].apply(parse_description)
        return df

    def load_consolidated_data(self, income_headers=None):
        income_headers = ['Date', 'Branch', 'Transaction Description', 'Purpose', 'Description', 'Business a/c']
        loader = IncomeDataLoader(self.account_file_path, income_headers)
        loader.load_income_data()
        dataframes = loader.get_dataframes()
        self.print_progress(f"Loaded {len(dataframes)} income data sheets from {self.account_file_path}.")
        for branch_name in dataframes.keys():
            print(f"Loaded DataFrame: {branch_name}")
        filename = os.path.basename(self.account_file_path)
        filename, _ = os.path.splitext(filename)
        output_excel = f"processed_{filename}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        members_df = pd.read_excel(self.consent_file_path)
        members_df = self.cleanup_ga_consent_list(members_df)
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            for branch_name, df in dataframes.items():
                branch_name_ga = branch_name.replace("_", " ")
                if branch_name == "London":
                    branch_name_ga = "London Central"
                if branch_name == "Teeside":
                    branch_name_ga = "Teesside"
                self.print_progress(f"Processing DataFrame: {branch_name}, shape: {df.shape}")
                summarize_df = self.load_and_clean_statement(df)
                self.print_progress(f"Processed DataFrame: {branch_name}, shape: {summarize_df.shape}")
                branch_members_df = members_df
                if 'NEC' not in branch_name:
                    branch_members_df = self.get_branch_members(branch_name_ga, self.BRANCH_NAME_FIELD, members_df)
                    self.print_progress(f"Branch members DataFrame shape: {branch_members_df.shape}")
                if branch_name != "Exeter":
                    summarize_df = self.match_payee_to_members(summarize_df, branch_members_df)
                else:
                    self.print_progress(f"No members found for branch '{branch_name}'. Skipping member matching.")
                summarize_df.to_excel(writer, sheet_name=branch_name, index=False)
                self.print_progress(f">>>>>>>>>>>> Saved processed DataFrame: {branch_name} to {output_excel}")

    def get_branch_members(self, branch_name, branch_field, members_df):
        return members_df[members_df[branch_field] == branch_name]

    def cleanup_ga_consent_list(self, ga_consent_list):
        sorted_df = ga_consent_list.sort_values('Completion time', ascending=False).drop_duplicates(subset=[self.FIRST_NAME_FIELD, self.LAST_NAME_FIELD], keep='first')
        filtered_df = ga_consent_list[ga_consent_list.index.isin(sorted_df.index)]
        consent_cols = [col for col in filtered_df.columns if 'gift aid' in col.lower()]
        if len(consent_cols) > 1:
            self.print_progress("**** ga_consent_list seems invalid. Output may be incorrect ")
            consent_cols = consent_cols[:1]
        filtered_df = filtered_df[~filtered_df[consent_cols].eq('No').any(axis=1)]
        filtered_df[self.MEMBERS_ID_FIELD] = filtered_df[self.MEMBERS_ID_FIELD].fillna('').astype(str).str.strip()
        return filtered_df

    def validate_files(self):
        if not os.path.isfile(self.account_file_path):
            print(f"Account file not found: {self.account_file_path}")
            exit(1)
        if not os.path.isfile(self.consent_file_path):
            print(f"GA consent file not found: {self.consent_file_path}")
            exit(1)

    def run(self):
        self.validate_files()
        self.load_consolidated_data()

if __name__ == "__main__":
    folder, account_file, consent_file = GAProcessor.parse_args()
    processor = GAProcessor(folder, account_file, consent_file)
    processor.run()
else:
    print("This script is intended to be run as a standalone program.")
    # If imported, the process_data function can be called directly.