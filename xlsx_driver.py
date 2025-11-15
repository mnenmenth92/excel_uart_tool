from openpyxl import load_workbook

class ExcelColumnReader:
    def __init__(self, file_path):
        """
        Initialize with the path to the Excel file.
        """
        self.file_path = file_path
        self.workbook = load_workbook(file_path, data_only=True)

    def read_column(self, sheet_name, column="A"):
        """
        Read all values from the specified column of the sheet.
        Returns a list of non-empty cell values.
        """
        if sheet_name not in self.workbook.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in {self.file_path}")

        sheet = self.workbook[sheet_name]
        column_values = []

        for row in sheet[column]:
            if row.value is not None:
                column_values.append(row.value)

        return column_values

# Example usage
if __name__ == "__main__":
    reader = ExcelColumnReader("CalibrationData.xlsx")
    raw_list_values = reader.read_column("raw_list", column="A")

    for val in raw_list_values:
        print(val)
