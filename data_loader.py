import pandas as pd

class DataLoader:
    def __init__(self, bank_statement_path: str, members_list_path: str):
        """
        Initialize the DataLoader with file paths.

        Args:
            bank_statement_path (str): Path to the bank statement file.
            members_list_path (str): Path to the members list file.
        """
        self.bank_statement_path = bank_statement_path
        self.members_list_path = members_list_path

    def load_bank_statement(self) -> pd.DataFrame:
        """
        Load and preprocess the bank statement.

        Returns:
            pd.DataFrame: Preprocessed bank statement DataFrame.
        """
        try:
            print("Loading bank statement...")
            bank_df = pd.read_excel(self.bank_statement_path)
            print("Bank statement loaded successfully.")

            # Preprocess the bank statement
            bank_df["Transaction Description"] = bank_df["Transaction Description"].astype(str)
            bank_df.dropna(subset=["Transaction Description", "Amount"], inplace=True)
            bank_df.rename(columns={"Business a/c": "Amount"}, inplace=True)

            print("Bank statement preprocessed successfully.")
            return bank_df
        except FileNotFoundError:
            print(f"Error: File not found at {self.bank_statement_path}")
            return pd.DataFrame()
        except Exception as e:
            print(f"Error while loading bank statement: {e}")
            return pd.DataFrame()

    def load_members_list(self) -> pd.DataFrame:
        """
        Load and preprocess the members list.

        Returns:
            pd.DataFrame: Preprocessed members list DataFrame.
        """
        try:
            print("Loading members list...")
            members_df = pd.read_excel(self.members_list_path)
            print("Members list loaded successfully.")

            # Preprocess the members list
            members_df["First Name"] = members_df["First Name"].str.strip().str.title()
            members_df["Last Name"] = members_df["Last Name"].str.strip().str.title()
            members_df.drop_duplicates(subset=["First Name", "Last Name"], keep="last", inplace=True)

            print("Members list preprocessed successfully.")
            return members_df
        except FileNotFoundError:
            print(f"Error: File not found at {self.members_list_path}")
            return pd.DataFrame()
        except Exception as e:
            print(f"Error while loading members list: {e}")
            return pd.DataFrame()

    def get_dataframes(self) -> tuple:
        """
        Load and return both the bank statement and members list DataFrames.

        Returns:
            tuple: (bank statement DataFrame, members list DataFrame)
        """
        bank_df = self.load_bank_statement()
        members_df = self.load_members_list()
        return bank_df, members_df