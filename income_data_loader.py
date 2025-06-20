import pandas as pd
import logging

logging.basicConfig(level=logging.INFO)

class IncomeDataLoader:
    def __init__(self, file_path, income_headers):
        """
        :param file_path: Path to the Excel file
        :param income_headers: List of headers that identify the title row of the income data
        """
        self.file_path = file_path
        self.income_headers = [header.lower().strip() for header in income_headers]
        self.dataframes = {}

    def _find_title_row(self, df):
        for idx, row in df.iterrows():
            normalized_row = [str(cell).lower().strip() for cell in row]
            if all(header in normalized_row for header in self.income_headers):
                return idx
        return None

    def load_income_data(self):
        xl = pd.ExcelFile(self.file_path)
        sheets = xl.sheet_names

        try:
            start_index = sheets.index("Bedford")
        except ValueError:
            raise ValueError("Sheet named 'Bedford' not found in workbook")

        for sheet in sheets[start_index:]:
            df_raw = xl.parse(sheet, header=None)
            title_row_index = self._find_title_row(df_raw)

            if title_row_index is None:
                logging.warning(f"Income data not found in sheet: {sheet}")
                continue

            header = df_raw.iloc[title_row_index].values
            data_start = title_row_index + 1

            # Collect rows until the first blank row (all NaNs)
            data_rows = []
            for i in range(data_start, len(df_raw)):
                row = df_raw.iloc[i]
                if row.isnull().all():
                    break
                data_rows.append(row.values)

            df = pd.DataFrame(data_rows, columns=header)
            var_name = f"{sheet}_df"
            self.dataframes[var_name] = df
            logging.info(f"Loaded data from sheet: {sheet} into '{var_name}'")

    def get_dataframes(self):
        return self.dataframes