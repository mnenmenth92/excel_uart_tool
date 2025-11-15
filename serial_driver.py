import serial
import configparser
import os
import time


class SerialDriver:
    errors = {
        -1005: "Wrong config file",
        -1006: "Wrong config file keys",
        -2000: "COM port connection error",
        -2001: "No device response (timeout)",
    }

    def __init__(self, com_port, config_file="config.ini"):
        """
        Initialize the serial driver using COM port and config file.
        Config file format:
        [connection_data]
        baudrate = 115200
        timeout = 1
        """

        self.ser = None
        self.com_port = com_port
        self.baudrate = 9600  # default
        self.timeout = 1      # default

        # Load config
        if not os.path.exists(config_file):
            self.error = -1005
            return

        self.config = configparser.ConfigParser()
        self.config.read(config_file)

        try:
            self.baudrate = self.config.getint("connection_data", "baudrate")
            self.timeout = self.config.getfloat("connection_data", "timeout")
        except (configparser.NoSectionError, configparser.NoOptionError):
            self.error = -1006
            return

        self.error = 0

    def connect(self):
        """
        Open the serial connection.
        """
        try:
            self.ser = serial.Serial(self.com_port, self.baudrate, timeout=self.timeout)
            if self.ser.is_open:
                print(f"Connected to {self.com_port}")
            else:
                self.ser.open()
        except serial.SerialException as e:
            print(f"Error connecting to {self.com_port}: {e}")
            self.ser = None
            self.error = -2000

    def send_integers(self, int_list):
        """
        Send a list of integers as a comma-separated string ending with a semicolon.
        """
        if self.ser is None or not self.ser.is_open:
            print("Serial port not connected.")
            return False

        try:
            message = ",".join(str(i) for i in int_list) + ";"
            self.ser.write(message.encode())
            print(f"Sent: {message}")
            return True
        except Exception as e:
            print(f"Error sending data: {e}")
            return False

    def wait_for_response(self, timeout=20):
        """
        Wait up to <timeout> seconds for device UART response.
        Returns decoded response string or error code -2001.
        """
        if self.ser is None or not self.ser.is_open:
            print("Serial port not connected.")
            return None

        print(f"Waiting for device response (timeout {timeout}s)...")
        start = time.time()
        buffer = b""

        while time.time() - start < timeout:
            try:
                if self.ser.in_waiting > 0:
                    buffer += self.ser.read(self.ser.in_waiting)

                    if buffer:
                        decoded = buffer.decode(errors="ignore")
                        print(f"Device response: {decoded}")
                        return decoded
            except Exception as e:
                print(f"Error reading response: {e}")
                return None

            time.sleep(0.1)

        print("No response (timeout)")
        return -2001  # <-- Timeout error code

    def disconnect(self):
        """
        Close the serial connection.
        """
        if self.ser and self.ser.is_open:
            self.ser.close()
            print(f"Disconnected from {self.com_port}")

    def get_error_text(self, error_num):
        return SerialDriver.errors[error_num]


# driver = SerialDriver("COM7")
# driver.connect()
# driver.send_integers([1,2,3])
# response = driver.wait_for_response(20)
# driver.disconnect()