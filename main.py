import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports
# Import the Excel driver class
from xlsx_driver import ExcelColumnReader
from serial_driver import SerialDriver
import time

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
        self.combobox.bind("<<ComboboxSelected>>", self.connect_to_selected_com)

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

    def connect_to_selected_com(self, v):
        selected_com = self.com_var.get()
        print(f'selected: {selected_com}')
        self.serial_driver = SerialDriver(selected_com, self.config_file)
        self.update_status('Selected: {selected_com}')
        # ToDo cover errors

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
            self.update_status(f"ERROR: {err_text}")
            return self.param_values
        return 0

    def write(self, skip_disconnect=False):
        """
        read excel
        send data
        :return:
        """
        error = self.read_parameters()
        if not error:
            self.update_status(f"Updating...")
            # time.sleep(0.5)  # for user to see a blink
            print(f'write: {self.param_values}')
            self.serial_driver.connect()
            self.serial_driver.send_integers(self.param_values)
            self.serial_driver.disconnect()
            self.update_status(f"Done")
            # Todo add waiting for response
        else:
            print('error!')

    def write_and_verify(self):
        self.write(skip_disconnect=False)  # ToDo put True
        print('write and verify')

    def verify(self):
        """
        read excel
        read data from device
        compare excel input with data on device
        :return:
        """
        error = self.read_parameters()
        if not error:
            self.update_status(f"Verifying...")
            self.serial_driver.connect()
            response = self.serial_driver.wait_for_response()

            if isinstance(response, int) and response < 0:  # is error (timeout)
                err_txt = self.serial_driver.get_error_text(response)
                self.update_status(err_txt)
            else:  # correct response
                if self.verify_response(response):
                    self.update_status('Verification: OK')
                else:
                    self.update_status('Verification: NOK')

            self.serial_driver.disconnect()
        else:
            print('error!')

    def update_status(self, text):
        """
        update status field
        :param text: text to be shown on status field
        :return:
        """
        self.status_var.set(text)
        self.root.update_idletasks()

    def verify_response(self, response):
        """
        Parse response string to list and compares to self.param_values
        :param response: comma divided string finished with ; and EoL
        :return: bool
        """

        try:
            # response_list = response[:-1].split(',')
            clean = response.strip().rstrip(';')
            response_list = clean.split(',')

            response_list = [int(x) for x in response_list]
            print(f"compare:\n {response_list}\n {self.param_values}")
            return self.param_values == response_list
        except Exception as e:
            print(e)
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = CalibrationApp(root)
    root.mainloop()
