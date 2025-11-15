import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports
# Import the Excel driver class
from xlsx_driver import ExcelColumnReader

# Function to list available COM ports
def list_serial_ports():
    return [port.device for port in serial.tools.list_ports.comports()]


# GUI application
class CalibrationApp:
    def __init__(self, root, config_file="config.ini"):
        self.root = root
        self.root.title("Calibration Tool")

        # Initialize Excel reader
        self.config_file = config_file
        self.excel_reader = ExcelColumnReader(self.config_file)
        self.param_values = []

        # COM selection
        ttk.Label(root, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.com_var = tk.StringVar()
        self.combobox = ttk.Combobox(root, textvariable=self.com_var, values=list_serial_ports(), width=20)
        self.combobox.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Buttons
        self.write_btn = ttk.Button(root, text="Write", command=self.write)
        self.write_btn.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.write_verify_btn = ttk.Button(root, text="Write & Verify", command=self.write_and_verify)
        self.write_verify_btn.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.verify_btn = ttk.Button(root, text="Verify", command=self.verify)
        self.verify_btn.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

        # Status Field
        status_frame = ttk.Frame(root, relief="groove", padding=(10, 5))
        status_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="w", background="#f0f0f0")
        self.status_label.pack(fill="x")

    def read_parameters(self):
        """
        read parameters from excel
        return: error - for codes go to xlsx_driver.py
        :return:
        """
        # reload file to grant newest values
        self.excel_reader = ExcelColumnReader(self.config_file)
        # read values from excel
        self.param_values = self.excel_reader.read_column("A")
        if isinstance(self.param_values, int) and self.param_values < 0:  # is error
            err_text = self.excel_reader.get_error_text(self.param_values)
            self.status_var.set(f"ERROR: {err_text}")
            return self.param_values
        return 0

    def write(self):
        error = self.read_parameters()
        if not error:
            print(f'write: {self.param_values}')
        else:
            print('error!')

    def write_and_verify(self):
        self.read_parameters()
        print('write and verify')

    def verify(self):
        self.read_parameters()
        print("verify")
        #


if __name__ == "__main__":
    root = tk.Tk()
    app = CalibrationApp(root)
    root.mainloop()
