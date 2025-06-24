import pandas as pd
import os

class ExcelHandler:
    def __init__(self, filepath, sheetname):
        self.filepath = filepath
        self.sheetname = sheetname
        # Load sheet with header at second row (index=1)
        self.df = pd.read_excel(filepath, header=0, sheet_name=sheetname)
        self.all_link = self.df['Notebook Link']

        # Initialize output DataFrame with Notebook Link and empty status column
        self.output_df = pd.DataFrame({
            'Notebook Link': self.all_link,
            'status': [''] * len(self.all_link)
        })

    def get_notebook_link(self, index):
        try:
            index = int(index)
            return self.all_link.iloc[index]
        except IndexError:
            return f"Index {index} out of range. Max index: {len(self.all_link) - 1}"
        except Exception as e:
            return f"Error: {str(e)}"

    def get_total_rows(self):
        return len(self.df)

    def update_status(self, index, status_text):
        try:
            index = int(index)
            self.output_df.at[index, 'status'] = status_text
        except Exception as e:
            print(f"Error updating status at index {index}: {e}")

    def add_column_with_value(self, index, param1, value1):
        if param1 not in self.output_df.columns:
            self.output_df[param1] = ''

        try:
            index = int(index)
            self.output_df.at[index, param1] = value1
        except Exception as e:
            print(f"Error setting column '{param1}' at index {index}: {e}")

    def save_to_csv(self, filename):
        self.output_df.to_csv(filename, index=False)

    def export_row_to_csv(self, index, csv_path="output.csv"):
        try:
            index = int(index)
            row = self.output_df.iloc[index]
            # Write header only if the file does not exist
            write_header = not os.path.exists(csv_path)
            row.to_frame().T.to_csv(csv_path, mode='a', header=write_header, index=False)
        except Exception as e:
            print(f"Error exporting row {index}: {e}")
