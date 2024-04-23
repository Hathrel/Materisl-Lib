import pandas as pd
import openpyxl as xl
import os

class MaterialsFile:
    def __init__(self, file_name, directory = os.path.join(os.path.expanduser("~"), "Downloads")):
        self.file_name = file_name
        self.headers = []
        self.directory = directory

    def open_file(self):
        file_path = os.path.join(self.directory, self.file_name)
        if os.path.splitext(file_path)[1] == ".csv":
            try:
                production_df = pd.read_csv(file_path)
                print("File loaded successfully!")
                return production_df
            except FileNotFoundError:
                print(f"That file doesn't exist in {self.directory}. Please try again.")
            except pd.errors.EmptyDataError:
                print("The file is empty. Please check the file content.")
            except pd.errors.ParserError:
                print("There was a parsing error while reading the file. Please check the file content.")
            except Exception as e:
                print(f"An error occurred: {e}. Please try again.")
            return None
        elif os.path.splitext(file_path)[1] == ".xlsx" | os.path.splitext(file_path)[1] == ".xls":
            try:
                    production_df = pd.read_excel(file_path)
                    print("File loaded successfully!")
                    return production_df
            except FileNotFoundError:
                print(f"That file doesn't exist in {self.directory}. Please try again.")
            except pd.errors.EmptyDataError:
                print("The file is empty. Please check the file content.")
            except pd.errors.ParserError:
                print("There was a parsing error while reading the file. Please check the file content.")
            except Exception as e:
                print(f"An error occurred: {e}. Please try again.")
            return None