import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports

# Function to list available COM ports
def list_serial_ports():
    return [port.device for port in serial.tools.list_ports.comports()]

# GUI application
class CalibrationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Calibration Tool")

        # COM selection
        ttk.Label(root, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.com_var = tk.StringVar()
        self.combobox = ttk.Combobox(root, textvariable=self.com_var, values=list_serial_ports(), width=20)
        self.combobox.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Buttons
        self.write_btn = ttk.Button(root, text="Write")
        self.write_btn.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.write_verify_btn = ttk.Button(root, text="Write & Verify")
        self.write_verify_btn.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.verify_btn = ttk.Button(root, text="Verify")
        self.verify_btn.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

        # Status Field
        status_frame = ttk.Frame(root, relief="groove", padding=(10,5))
        status_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="w", background="#f0f0f0")
        self.status_label.pack(fill="x")


if __name__ == "__main__":
    root = tk.Tk()
    app = CalibrationApp(root)
    root.mainloop()
