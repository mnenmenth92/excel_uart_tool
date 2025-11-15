# README

## Overview
This application simplifies the calibration of a device by sending parameters directly from a predefined **.xlsx** file over a UART connection. It also allows verification of parameters already stored on the device. The app provides a compact Tkinter interface with essential controls for writing and verifying calibration data.

---

## Excel File Structure
The application uses a fixed-name **.xlsx** file located in the app’s directory.

The file contains **two sheets**:

### 1. User Sheet
- Designed for user convenience.  
- Can be arranged freely for readability and usability.  
- *Not used directly for UART transfer.*

### 2. Transfer Sheet
- Contains the actual list of parameters in a structured format.  
- Used for UART communication with the device.  
- Must follow the fixed parameter list expected by the app and the device.
- All parameters must be listed in column A without any empty cells or additional text.

---

## Application Interface
The Tkinter interface includes:

- **COM port selector**  
- **Write** button  
- **Write and Verify** button  
- **Verify** button  
- **Status field** displaying progress and results

---

## Functionality

### Write
1. Establishes a UART connection.  
2. Reads the parameter list from the transfer sheet of the Excel file.  
3. Sends parameters to the device.  
4. Status updates:
   - “Uploading…”  
   - “Done” when finished
5. Close UART connection

---

### Write and Verify
1. Establishes a UART connection.  
2. Reads parameters from the transfer sheet.  
3. Sends the parameters to the device.  
4. Reads back parameters from the device.  
5. Compares them with values from the Excel file.  
6. Status updates:
   - “Uploading…”  
   - “Reading…”  
   - **OK** if all parameters match  
   - **NOK** if any value differs
7. Close UART connection

---

### Verify
1. Establishes a UART connection.  
2. Reads parameters from the transfer sheet.  
3. Reads the current parameters from the device.  
4. Compares the two sets.  
5. Status updates:
   - “Reading…”  
   - **OK** if all parameters match  
   - **NOK** if any value differs
6. Close UART connection

---

## Requirements
- Python 3.x  
- Required packages listed in `requirements.txt`  
- Predefined `.xlsx` file in the application directory  
- UART driver installed for the selected COM port  

---

## Usage
1. Place the fixed-name Excel file in the application directory.  
2. Run the application.  
3. Select the COM port.  
4. Click one of the available actions:
   - **Write**  
   - **Write and Verify**  
   - **Verify**  
5. Monitor the status field for progress and results.

---

## Notes
- The user sheet may be arranged freely.  
- The transfer sheet must remain in the correct struc
