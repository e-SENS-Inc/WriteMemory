# -*- coding: utf-8 -*-
"""
Created on Fri Apr 26 10:41:21 2024

@author: Jason

V4: 7/8/2024: Use the sensorConfigMenuChange function to:
        1. Hide the Rinse and T1 frames when not needed
        2. Change from Cal 5 to Cal 7 for pH_Cl_Cart
    Pulls clean solutions numbers to save in memory for rinse values (because of the checks in the firmware there needs to be something saved in rinse)
    update_values function no longer updates rinse or T1 if it isn't being used in the configurtaion
    
V9: 11/12/2024: Removing the pH only Clean and Cal 7 code
11/18/2024: Adding check on SN and sensor config
"""

# import tkinter
import customtkinter
import struct
import serial
import time
import serial.tools.list_ports
# import sys
# import subprocess
import tkinter as tk
from tkinter import messagebox
import binascii
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
from datetime import datetime, timedelta


from typing import Union, Tuple, Optional

""" Ari's code that will break this entire script"""
# Shared dictionary to store data
shared_data = {}

# Path to the service account JSON key file
service_account_file = r'circular-shield-424018-c4-b6df851a2d3b.json'  

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(service_account_file, scope)
client = gspread.authorize(credentials)

# Open the Google Sheet
sheet_name = 'Cartridges to place into Fishbowl'  # Replace with the actual name of your Google Sheet
sheet = client.open(sheet_name)

# Open the "Cartridges" sheet
cartridge_sheet = sheet.worksheet('Cartridges')

# List of sheets to fetch data from
sheets_to_fetch = ['Rinse', 'T1', 'Cal 5', 'Cal 6', 'Clean']

def get_today_date():
    """Returns today's date in MM/DD/YY format."""
    return datetime.now().strftime("%m/%d/%y")

def get_future_date(days=45):
    """Returns the date `days` days from today in MM/DD/YY format."""
    future_date = datetime.now() + timedelta(days=days)
    return future_date.strftime("%m/%d/%y")

# Function to fetch data based on batch serial number, ignoring empty cells
def fetch_batch_data(batch_serial_number, sheet_name):
    worksheet = sheet.worksheet(sheet_name)
    try:
        print(f"Fetching data for {batch_serial_number} from {sheet_name}")  # Debug print
        cell = worksheet.find(batch_serial_number)
        if cell:
            row_data = worksheet.row_values(cell.row)

            # Dynamically find the indices for relevant data columns
            if sheet_name == 'T1':
                required_data = [value for value in row_data[4:5] if value]  # Only one item for T1
            else:
                required_data = [value for value in row_data[4:12] if value]  # Default data from columns E to L, ignoring empty cells

            print(f"Fetched data from {sheet_name} for batch {batch_serial_number}: {required_data}")  # Debug print
            return required_data
    except gspread.exceptions.APIError:
        print(f"Batch serial number {batch_serial_number} not found in {sheet_name}.")  # Debug print
    return None


def fetch_data_for_cartridge(serial_number):
    try:
        cell = cartridge_sheet.find(serial_number)
        if cell:
            row_data = cartridge_sheet.row_values(cell.row)
            # Assuming the order of batch serial numbers in the row matches the expected order
            batch_serial_numbers = {}

            if len(row_data) > 2:
                batch_serial_numbers["Rinse"] = row_data[2]
            if len(row_data) > 4:
                batch_serial_numbers["Cal 5"] = row_data[4]
            if len(row_data) > 5:
                batch_serial_numbers["Cal 6"] = row_data[5]
            if len(row_data) > 6:
                batch_serial_numbers["Clean"] = row_data[6]
            if len(row_data) > 3:
                batch_serial_numbers["T1"] = row_data[3]

            print(f"Batch serial numbers for cartridge {serial_number}: {batch_serial_numbers}")  # Debug print

            data = {}
            for sheet_name, batch_serial_number in batch_serial_numbers.items():
                batch_data = fetch_batch_data(batch_serial_number, sheet_name)
                if batch_data:
                    data[sheet_name] = batch_data
                    print(f"Added data for {sheet_name}: {batch_data}")  # Debug print
                else:
                    print(f"No data found for {sheet_name}.")  # Debug print
            print(f"Final data dictionary: {data}")  # Debug print
            return data
    except gspread.exceptions.APIError:
        print(f"Serial number {serial_number} not found in Cartridges sheet.")  # Debug print
    return None



# Function to register data into the GUI
def register_memory_values():
    global shared_data
    cartridge_serial_number = app.cartSN.get()
    if not cartridge_serial_number:
        messagebox.showerror("Error", "Please enter a cartridge serial number")
        return

    data = fetch_data_for_cartridge(cartridge_serial_number)
    if not data:
        messagebox.showerror("Error", "Serial number not found")
        return

    shared_data['cartridge_serial_number'] = cartridge_serial_number
    shared_data['data'] = data  # Store the fetched data in the shared dictionary
    print(f"Shared data: {shared_data}")  # Debug print

    app.update_values()
    print("Data fetched and displayed.")
    print(f"Stored cartridge serial number: {cartridge_serial_number}")
    
    
# End section of Ari's code that will break this entire script

# System Settings
customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("blue")


# # Way to redirect sys.stdout to text box in gui
# class ConsoleRedirector:
#     def __init__(self, widget):
#         self.widget = widget

#     def write(self, text):
#         self.widget.insert(tk.END, text)
#         self.widget.see(tk.END)  # Auto-scroll to the bottom

# Class to open a confirmation dialog asking the user to accept
class CTkConfirmDialog(customtkinter.CTkToplevel):
    """
    Dialog with extra window, cancel and confirm button.
    For detailed information check out the documentation.
    """

    def __init__(self,
                 fg_color: Optional[Union[str, Tuple[str, str]]] = None,
                 text_color: Optional[Union[str, Tuple[str, str]]] = None,
                 button_fg_color: Optional[Union[str, Tuple[str, str]]] = None,
                 button_hover_color: Optional[Union[str, Tuple[str, str]]] = None,
                 button_text_color: Optional[Union[str, Tuple[str, str]]] = None,
                 # entry_fg_color: Optional[Union[str, Tuple[str, str]]] = None,
                 # entry_border_color: Optional[Union[str, Tuple[str, str]]] = None,
                 # entry_text_color: Optional[Union[str, Tuple[str, str]]] = None,

                 title: str = "CTkDialog",
                 font: Optional[Union[tuple, customtkinter.CTkFont]] = None,
                 text: str = "CTkDialog"):

        super().__init__(fg_color=fg_color)

        self._fg_color = customtkinter.ThemeManager.theme["CTkToplevel"]["fg_color"] if fg_color is None else self._check_color_type(fg_color)
        self._text_color = customtkinter.ThemeManager.theme["CTkLabel"]["text_color"] if text_color is None else self._check_color_type(button_hover_color)
        self._button_fg_color = customtkinter.ThemeManager.theme["CTkButton"]["fg_color"] if button_fg_color is None else self._check_color_type(button_fg_color)
        self._button_hover_color = customtkinter.ThemeManager.theme["CTkButton"]["hover_color"] if button_hover_color is None else self._check_color_type(button_hover_color)
        self._button_text_color = customtkinter.ThemeManager.theme["CTkButton"]["text_color"] if button_text_color is None else self._check_color_type(button_text_color)
        # self._entry_fg_color = ThemeManager.theme["CTkEntry"]["fg_color"] if entry_fg_color is None else self._check_color_type(entry_fg_color)
        # self._entry_border_color = ThemeManager.theme["CTkEntry"]["border_color"] if entry_border_color is None else self._check_color_type(entry_border_color)
        # self._entry_text_color = ThemeManager.theme["CTkEntry"]["text_color"] if entry_text_color is None else self._check_color_type(entry_text_color)

        self._user_input: Union[str, None] = None
        self._running: bool = False
        self._title = title
        self._text = text
        self._font = font

        self.title(self._title)
        self.lift()  # lift window on top
        self.attributes("-topmost", True)  # stay on top
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.after(10, self._create_widgets)  # create widgets with slight delay, to avoid white flickering of background
        self.resizable(False, False)
        self.grab_set()  # make other windows not clickable

    def _create_widgets(self):
        self.grid_columnconfigure((0, 1), weight=1)
        self.rowconfigure(0, weight=1)

        self._label = customtkinter.CTkLabel(master=self,
                               width=300,
                               wraplength=300,
                               fg_color="transparent",
                               text_color=self._text_color,
                               text=self._text,
                               font=self._font)
        self._label.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky="ew")

        # self._entry = CTkEntry(master=self,
        #                        width=230,
        #                        fg_color=self._entry_fg_color,
        #                        border_color=self._entry_border_color,
        #                        text_color=self._entry_text_color,
        #                        font=self._font)
        # self._entry.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 20), sticky="ew")

        self._ok_button = customtkinter.CTkButton(master=self,
                                    width=100,
                                    border_width=0,
                                    fg_color=self._button_fg_color,
                                    hover_color=self._button_hover_color,
                                    text_color=self._button_text_color,
                                    text='Confirm',
                                    font=self._font,
                                    command=self._ok_event)
        self._ok_button.grid(row=2, column=0, columnspan=1, padx=(20, 10), pady=(0, 20), sticky="ew")

        self._cancel_button = customtkinter.CTkButton(master=self,
                                        width=100,
                                        border_width=0,
                                        fg_color=self._button_fg_color,
                                        hover_color=self._button_hover_color,
                                        text_color=self._button_text_color,
                                        text='Cancel',
                                        font=self._font,
                                        command=self._cancel_event)
        self._cancel_button.grid(row=2, column=1, columnspan=1, padx=(10, 20), pady=(0, 20), sticky="ew")

        self.after(150, lambda: self._ok_button.focus())  # set focus to entry with slight delay, otherwise it won't work
        # self._entry.bind("<Return>", self._ok_event)

    def _ok_event(self, event=None):
        self._user_input = True
        self.grab_release()
        self.destroy()

    def _on_closing(self):
        self.grab_release()
        self.destroy()

    def _cancel_event(self):
        self._user_input = False
        self.grab_release()
        self.destroy()

    def get_input(self):
        self.master.wait_window(self)
        return self._user_input

class App(customtkinter.CTk):
    def _on_closing(self):
        self.grab_release()
        self.destroy()
        sys.exit()

    def __init__(self):
        super().__init__()
        
        # Configure window
        self.geometry(f"{self.winfo_screenwidth()/2}x{self.winfo_screenheight()/2}")
        self.title("Write Cartridge Memory")
        self.grid_columnconfigure(0, weight=0)
        self.grid_rowconfigure(0, weight=0)

        # Create all the frames first
        self.sensor_config_frame = customtkinter.CTkFrame(self, width=180, height=110)
        self.sensor_config_frame.grid(row=0, column=0, padx=(20, 0), pady=10, sticky="nsew")

        self.cartinfo_frame = customtkinter.CTkFrame(self)
        self.cartinfo_frame.grid(row=1, column=0, padx=(20, 0), pady=10, sticky="nsew")
        
        self.dates_frame = customtkinter.CTkFrame(self)
        self.dates_frame.grid(row=2, column=0, padx=(20, 0), pady=10, sticky="nsew")
        
        self.clcal_frame = customtkinter.CTkFrame(self)
        self.clcal_frame.grid(row=3, column=0, padx=(20, 0), pady=10, sticky="nsew")
        self.clcal_frame.grid_columnconfigure(1, weight=1)
        
        self.misc_frame = customtkinter.CTkFrame(self)
        self.misc_frame.grid(row=3, column=1, padx=(20, 20), pady=10, sticky="nsew")
        self.misc_frame.grid_columnconfigure(0, weight=1)
        
        self.solinfo_frame = customtkinter.CTkFrame(self)
        self.solinfo_frame.grid(row=0, column=1, rowspan=3, columnspan=2, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.solinfo_frame.grid_columnconfigure(0, weight=1)
        self.solinfo_frame.grid_columnconfigure(1, weight=1)
        self.solinfo_frame.grid_columnconfigure(2, weight=1)
        self.solinfo_frame.grid_columnconfigure(3, weight=1)
        self.solinfo_frame.grid_columnconfigure(4, weight=1)
        
        self.button_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=3, column=2, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.button_frame.grid_columnconfigure((0, 1), weight=1)
        
        # Sensor config frame
        self.sensor_config_check = customtkinter.CTkCheckBox(master=self.sensor_config_frame, text="Write Sensor Config")
        self.sensor_config_check.grid(row=0,column=0, padx=10, pady=10, sticky="w")
        
        # self.label_config = customtkinter.CTkLabel(master=self.sensor_config_frame, text="Sensor Config:")
        # self.label_config.grid(row=1, column=0, padx=5, pady=5)
        self.config_menu = customtkinter.CTkOptionMenu(self.sensor_config_frame, values=["CR2300 (V7)", "CR800", "CR1300", "V11", "V12"], command=self.sensorConfigMenuChange)
        self.config_menu.grid(row=1, column=0, padx=5, pady=5)

        # Cartridge info frame
        self.cart_info_check = customtkinter.CTkCheckBox(master=self.cartinfo_frame, text="Write Cartridge Info")
        self.cart_info_check.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        
        self.label_cartSN = customtkinter.CTkLabel(master=self.cartinfo_frame, text="Cartridge SN:")
        self.label_cartSN.grid(row=3, column=0, padx=5, pady=0, sticky="e")
        self.cartSN = customtkinter.CTkEntry(master=self.cartinfo_frame, width=100, height=30)
        self.cartSN.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.cartSN.insert(0, shared_data.get('cartridge_serial_number', ''))
        
        self.label_sensorSN = customtkinter.CTkLabel(master=self.cartinfo_frame, text="Sensor SN:")
        self.label_sensorSN.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.sensorSN = customtkinter.CTkEntry(master=self.cartinfo_frame, width=100, height=30)
        self.sensorSN.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

        # self.label_maxdays = customtkinter.CTkLabel(master=self.cartinfo_frame, text="Max Days:")
        # self.label_maxdays.grid(row=8, column=0, padx=5, pady=5, sticky="e")
        
        # Max days is no longer visible because it is not important for the app
        # But it is important to keep the byte array transferred correct, so DO NOT REMOVE unless you want to rework the writing function
        self.maxdays = customtkinter.CTkEntry(master=self.cartinfo_frame, width=100, height=30)
        # self.maxdays.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.maxdays.insert(0, 45)

        self.label_maxtests = customtkinter.CTkLabel(master=self.cartinfo_frame, text="Max Tests:")
        self.label_maxtests.grid(row=9, column=0, padx=5, pady=5, sticky="e")
        self.maxtests = customtkinter.CTkEntry(master=self.cartinfo_frame, width=100, height=30)
        self.maxtests.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
        self.maxtests.insert(0, 100)

        self.label_maxcals = customtkinter.CTkLabel(master=self.cartinfo_frame, text="Max Cals:")
        self.label_maxcals.grid(row=10, column=0, padx=5, pady=5, sticky="e")
        self.maxcals = customtkinter.CTkEntry(master=self.cartinfo_frame, width=100, height=30)
        self.maxcals.grid(row=10, column=1, padx=5, pady=5, sticky="ew")
        self.maxcals.insert(0, 30)

        # Dates frame
        self.dates_check = customtkinter.CTkCheckBox(master=self.dates_frame, text="Write Dates")
        self.dates_check.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        
        self.label_expdate = customtkinter.CTkLabel(master=self.dates_frame, text="Expiration Date:")
        self.label_expdate.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.expdate = customtkinter.CTkEntry(master=self.dates_frame, width=100, height=30, placeholder_text=("MM/DD/YY"))
        self.expdate.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.expdate.insert(0, get_future_date(46))

        self.label_hyddate = customtkinter.CTkLabel(master=self.dates_frame, text="Hyrdration Date:")
        self.label_hyddate.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.hyddate = customtkinter.CTkEntry(master=self.dates_frame, width=100, height=30, placeholder_text=("MM/DD/YY"))
        self.hyddate.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.hyddate.insert(0, get_today_date())

        # Cl Calibration frame
        self.clcal_check = customtkinter.CTkCheckBox(master=self.clcal_frame, text="Write Cl Calibration")
        self.clcal_check.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")
        
        self.label_slope = customtkinter.CTkLabel(master=self.clcal_frame, text="Slope:")
        self.label_slope.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.label_int = customtkinter.CTkLabel(master=self.clcal_frame, text="Int:")
        self.label_int.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        
        self.label_fcl_low = customtkinter.CTkLabel(master=self.clcal_frame, text="FCl Low:")
        self.label_fcl_low.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.fcl_low_slope = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.fcl_low_slope.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.fcl_low_int = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.fcl_low_int.grid(row=2, column=2, padx=5, pady=5, sticky="ew")
        self.fcl_low_slope.insert(0, -40.592)
        self.fcl_low_int.insert(0, -1.625)
        
        self.label_fcl_high = customtkinter.CTkLabel(master=self.clcal_frame, text="FCl High:")
        self.label_fcl_high.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.fcl_high_slope = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.fcl_high_slope.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.fcl_high_int = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.fcl_high_int.grid(row=3, column=2, padx=5, pady=5, sticky="ew")
        self.fcl_high_slope.insert(0, -70.932)
        self.fcl_high_int.insert(0, 8.388)
        
        self.label_tcl_low = customtkinter.CTkLabel(master=self.clcal_frame, text="TCl Low:")
        self.label_tcl_low.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.tcl_low_slope = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.tcl_low_slope.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.tcl_low_int = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.tcl_low_int.grid(row=4, column=2, padx=5, pady=5, sticky="ew")
        self.tcl_low_slope.insert(0, -65.987)
        self.tcl_low_int.insert(0, -7.547)
        
        self.label_tcl_high = customtkinter.CTkLabel(master=self.clcal_frame, text="TCl High:")
        self.label_tcl_high.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.tcl_high_slope = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.tcl_high_slope.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.tcl_high_int = customtkinter.CTkEntry(master=self.clcal_frame, width=75, height=30)
        self.tcl_high_int.grid(row=5, column=2, padx=5, pady=5, sticky="ew")
        self.tcl_high_slope.insert(0, -74.374)
        self.tcl_high_int.insert(0, -4.779)
        
        # Misc frame        
        self.therm_frame = customtkinter.CTkFrame(self.misc_frame)
        self.therm_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.therm_check = customtkinter.CTkCheckBox(master=self.therm_frame, text="Write Thermistor Slope")
        self.therm_check.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.label_therm = customtkinter.CTkLabel(master=self.therm_frame, text="Therm Slope:")
        self.label_therm.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.therm = customtkinter.CTkEntry(master=self.therm_frame, width=75, height=30)
        self.therm.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        self.valve_frame = customtkinter.CTkFrame(self.misc_frame)
        self.valve_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.valve_check = customtkinter.CTkCheckBox(master=self.valve_frame, text="Write Valve Setup")
        self.valve_check.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.label_valve = customtkinter.CTkLabel(master=self.valve_frame, text="Valve Setup:")
        self.label_valve.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.valve = customtkinter.CTkOptionMenu(self.valve_frame, values=["V1 Normal", "V2 Alternate"])
        self.valve.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # Solution info frame
        self.sol_info_check = customtkinter.CTkCheckBox(master=self.solinfo_frame, text="Write Solution Info")
        self.sol_info_check.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")
        
        # Generate button
        self.generate_button = customtkinter.CTkButton(master=self.solinfo_frame, text="Generate Solution Values", command=register_memory_values)
        self.generate_button.grid(row=0, column=2, columnspan=4, padx=5, pady=10, sticky="ew")
        
        # Rinse
        self.rinse_frame = customtkinter.CTkFrame(self.solinfo_frame)
        self.rinse_frame.grid(row=1, column=0, padx=(5, 5), pady=(5, 5), sticky="nsew")
        self.rinse_frame.grid_columnconfigure(1, weight=1)
        self.label_rinse = customtkinter.CTkLabel(master=self.rinse_frame, text="Rinse:")
        self.label_rinse.grid(row=1, column=0, padx=5, pady=5)

        self.label_rinse_pH = customtkinter.CTkLabel(master=self.rinse_frame, text="pH:")
        self.label_rinse_pH.grid(row=2, column=0, padx=5, pady=5)
        self.rinse_pH = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_pH.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_pH.insert(0, 7.5)

        self.label_rinse_Ca = customtkinter.CTkLabel(master=self.rinse_frame, text="Ca:")
        self.label_rinse_Ca.grid(row=3, column=0, padx=5, pady=5)
        self.rinse_Ca = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_Ca.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_Ca.insert(0, 150)

        self.label_rinse_TH = customtkinter.CTkLabel(master=self.rinse_frame, text="TH:")
        self.label_rinse_TH.grid(row=4, column=0, padx=5, pady=5)
        self.rinse_TH = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_TH.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_TH.insert(0, 300)

        self.label_rinse_NH4 = customtkinter.CTkLabel(master=self.rinse_frame, text="NH4:")
        self.label_rinse_NH4.grid(row=5, column=0, padx=5, pady=5)
        self.rinse_NH4 = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_NH4.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_NH4.insert(0, 0.5)

        self.label_rinse_cond = customtkinter.CTkLabel(master=self.rinse_frame, text="Cond:")
        self.label_rinse_cond.grid(row=6, column=0, padx=5, pady=5)
        self.rinse_cond = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_cond.grid(row=6, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_cond.insert(0, 966)

        self.label_rinse_IS = customtkinter.CTkLabel(master=self.rinse_frame, text="IS:")
        self.label_rinse_IS.grid(row=7, column=0, padx=5, pady=5)
        self.rinse_IS = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_IS.grid(row=7, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_IS.insert(0, 0.0132)

        self.label_rinse_KT = customtkinter.CTkLabel(master=self.rinse_frame, text="KT:")
        self.label_rinse_KT.grid(row=8, column=0, padx=5, pady=5)
        self.rinse_KT = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_KT.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_KT.insert(0, -0.0132)

        self.label_rinse_CondTComp = customtkinter.CTkLabel(master=self.rinse_frame, text="Cond T Comp:")
        self.label_rinse_CondTComp.grid(row=9, column=0, padx=5, pady=5)
        self.rinse_CondTComp = customtkinter.CTkEntry(master=self.rinse_frame, width=75, height=30)
        self.rinse_CondTComp.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
        self.rinse_CondTComp.insert(0, 0.0213)
        
        # Clean
        self.clean_frame = customtkinter.CTkFrame(self.solinfo_frame)
        self.clean_frame.grid(row=1, column=1, padx=(5, 5), pady=(5, 5), sticky="nsew")
        self.clean_frame.grid_columnconfigure(1, weight=1)
        self.label_clean = customtkinter.CTkLabel(master=self.clean_frame, text="Clean:")
        self.label_clean.grid(row=1, column=0, padx=5, pady=5)

        self.label_clean_pH = customtkinter.CTkLabel(master=self.clean_frame, text="pH:")
        self.label_clean_pH.grid(row=2, column=0, padx=5, pady=5)
        self.clean_pH = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_pH.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.clean_pH.insert(0, 8.97)

        self.label_clean_Ca = customtkinter.CTkLabel(master=self.clean_frame, text="Ca:")
        self.label_clean_Ca.grid(row=3, column=0, padx=5, pady=5)
        self.clean_Ca = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_Ca.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.clean_Ca.insert(0, 20)

        self.label_clean_TH = customtkinter.CTkLabel(master=self.clean_frame, text="TH:")
        self.label_clean_TH.grid(row=4, column=0, padx=5, pady=5)
        self.clean_TH = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_TH.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.clean_TH.insert(0, 40)

        self.label_clean_NH4 = customtkinter.CTkLabel(master=self.clean_frame, text="NH4:")
        self.label_clean_NH4.grid(row=5, column=0, padx=5, pady=5)
        self.clean_NH4 = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_NH4.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.clean_NH4.insert(0, 0)

        self.label_clean_cond = customtkinter.CTkLabel(master=self.clean_frame, text="Cond:")
        self.label_clean_cond.grid(row=6, column=0, padx=5, pady=5)
        self.clean_cond = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_cond.grid(row=6, column=1, padx=5, pady=5, sticky="ew")
        self.clean_cond.insert(0, 1063)

        self.label_clean_IS = customtkinter.CTkLabel(master=self.clean_frame, text="IS:")
        self.label_clean_IS.grid(row=7, column=0, padx=5, pady=5)
        self.clean_IS = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_IS.grid(row=7, column=1, padx=5, pady=5, sticky="ew")
        self.clean_IS.insert(0, 0.00954)

        self.label_clean_KT = customtkinter.CTkLabel(master=self.clean_frame, text="KT:")
        self.label_clean_KT.grid(row=8, column=0, padx=5, pady=5)
        self.clean_KT = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_KT.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.clean_KT.insert(0, -0.0098)

        self.label_clean_CondTComp = customtkinter.CTkLabel(master=self.clean_frame, text="Cond T Comp:")
        self.label_clean_CondTComp.grid(row=9, column=0, padx=5, pady=5)
        self.clean_CondTComp = customtkinter.CTkEntry(master=self.clean_frame, width=75, height=30)
        self.clean_CondTComp.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
        self.clean_CondTComp.insert(0, 0.021)

        # Cal 5
        self.cal_5_frame = customtkinter.CTkFrame(self.solinfo_frame)
        self.cal_5_frame.grid(row=1, column=2, padx=(5, 5), pady=(5, 5), sticky="nsew")
        self.cal_5_frame.grid_columnconfigure(1, weight=1)
        self.label_cal_5 = customtkinter.CTkLabel(master=self.cal_5_frame, text="Cal 5:")
        self.label_cal_5.grid(row=1, column=0, padx=5, pady=5)

        self.label_cal_5_pH = customtkinter.CTkLabel(master=self.cal_5_frame, text="pH:")
        self.label_cal_5_pH.grid(row=2, column=0, padx=5, pady=5)
        self.cal_5_pH = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_pH.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_pH.insert(0, 6.03)

        self.label_cal_5_Ca = customtkinter.CTkLabel(master=self.cal_5_frame, text="Ca:")
        self.label_cal_5_Ca.grid(row=3, column=0, padx=5, pady=5)
        self.cal_5_Ca = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_Ca.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_Ca.insert(0, 300)

        self.label_cal_5_TH = customtkinter.CTkLabel(master=self.cal_5_frame, text="TH:")
        self.label_cal_5_TH.grid(row=4, column=0, padx=5, pady=5)
        self.cal_5_TH = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_TH.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_TH.insert(0, 600)

        self.label_cal_5_NH4 = customtkinter.CTkLabel(master=self.cal_5_frame, text="NH4:")
        self.label_cal_5_NH4.grid(row=5, column=0, padx=5, pady=5)
        self.cal_5_NH4 = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_NH4.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_NH4.insert(0, 2.7)

        self.label_cal_5_cond = customtkinter.CTkLabel(master=self.cal_5_frame, text="Cond:")
        self.label_cal_5_cond.grid(row=6, column=0, padx=5, pady=5)
        self.cal_5_cond = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_cond.grid(row=6, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_cond.insert(0, 2169)

        self.label_cal_5_IS = customtkinter.CTkLabel(master=self.cal_5_frame, text="IS:")
        self.label_cal_5_IS.grid(row=7, column=0, padx=5, pady=5)
        self.cal_5_IS = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_IS.grid(row=7, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_IS.insert(0, 0.0263)

        self.label_cal_5_KT = customtkinter.CTkLabel(master=self.cal_5_frame, text="KT:")
        self.label_cal_5_KT.grid(row=8, column=0, padx=5, pady=5)
        self.cal_5_KT = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_KT.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_KT.insert(0, -0.0097)

        self.label_cal_5_CondTComp = customtkinter.CTkLabel(master=self.cal_5_frame, text="Cond T Comp:")
        self.label_cal_5_CondTComp.grid(row=9, column=0, padx=5, pady=5)
        self.cal_5_CondTComp = customtkinter.CTkEntry(master=self.cal_5_frame, width=75, height=30)
        self.cal_5_CondTComp.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
        self.cal_5_CondTComp.insert(0, 0.0204)
        
        # Cal 6
        self.cal_6_frame = customtkinter.CTkFrame(self.solinfo_frame)
        self.cal_6_frame.grid(row=1, column=3, padx=(5, 5), pady=(5, 5), sticky="nsew")
        self.cal_6_frame.grid_columnconfigure(1, weight=1)
        self.label_cal_6 = customtkinter.CTkLabel(master=self.cal_6_frame, text="Cal 6:")
        self.label_cal_6.grid(row=1, column=0, padx=5, pady=5)

        self.label_cal_6_pH = customtkinter.CTkLabel(master=self.cal_6_frame, text="pH:")
        self.label_cal_6_pH.grid(row=2, column=0, padx=5, pady=5)
        self.cal_6_pH = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_pH.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_pH.insert(0, 4.21)

        self.label_cal_6_Ca = customtkinter.CTkLabel(master=self.cal_6_frame, text="Ca:")
        self.label_cal_6_Ca.grid(row=3, column=0, padx=5, pady=5)
        self.cal_6_Ca = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_Ca.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_Ca.insert(0, 0)

        self.label_cal_6_TH = customtkinter.CTkLabel(master=self.cal_6_frame, text="TH:")
        self.label_cal_6_TH.grid(row=4, column=0, padx=5, pady=5)
        self.cal_6_TH = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_TH.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_TH.insert(0, 0)

        self.label_cal_6_NH4 = customtkinter.CTkLabel(master=self.cal_6_frame, text="NH4:")
        self.label_cal_6_NH4.grid(row=5, column=0, padx=5, pady=5)
        self.cal_6_NH4 = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_NH4.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_NH4.insert(0, 1.1)

        self.label_cal_6_cond = customtkinter.CTkLabel(master=self.cal_6_frame, text="Cond:")
        self.label_cal_6_cond.grid(row=6, column=0, padx=5, pady=5)
        self.cal_6_cond = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_cond.grid(row=6, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_cond.insert(0, 335)

        self.label_cal_6_IS = customtkinter.CTkLabel(master=self.cal_6_frame, text="IS:")
        self.label_cal_6_IS.grid(row=7, column=0, padx=5, pady=5)
        self.cal_6_IS = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_IS.grid(row=7, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_IS.insert(0, 0.00432)

        self.label_cal_6_KT = customtkinter.CTkLabel(master=self.cal_6_frame, text="KT:")
        self.label_cal_6_KT.grid(row=8, column=0, padx=5, pady=5)
        self.cal_6_KT = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_KT.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_KT.insert(0, -0.0025)

        self.label_cal_6_CondTComp = customtkinter.CTkLabel(master=self.cal_6_frame, text="Cond T Comp:")
        self.label_cal_6_CondTComp.grid(row=9, column=0, padx=5, pady=5)
        self.cal_6_CondTComp = customtkinter.CTkEntry(master=self.cal_6_frame, width=75, height=30)
        self.cal_6_CondTComp.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
        self.cal_6_CondTComp.insert(0, 0.0219)
        
        # T1
        self.T1_frame = customtkinter.CTkFrame(self.solinfo_frame)
        self.T1_frame.grid(row=1, column=4, padx=(5, 5), pady=(5, 5), sticky="nsew")
        self.T1_frame.grid_columnconfigure(1, weight=1)
        self.label_T1 = customtkinter.CTkLabel(master=self.T1_frame, text="T1:")
        self.label_T1.grid(row=1, column=0, padx=5, pady=5)

        self.label_T1_HCl_N = customtkinter.CTkLabel(master=self.T1_frame, text="HCl N:")
        self.label_T1_HCl_N.grid(row=2, column=0, padx=5, pady=5)
        self.T1_HCl_N = customtkinter.CTkEntry(master=self.T1_frame, width=75, height=30)
        self.T1_HCl_N.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.T1_HCl_N.insert(0, 0.060)

        # Buttons
        self.label_device = customtkinter.CTkLabel(master=self.button_frame, text="Select Device:", font=("Arial", 18))
        self.label_device.grid(row=0, column=0, padx=10, pady=0, columnspan=2)
        self.device_dropdown = customtkinter.CTkOptionMenu(master=self.button_frame, values=[], state="disabled")
        self.device_dropdown.grid(row=1, column=0, padx=5, pady=10, sticky="ew")
        self.device_dropdown.set("Search for devices ->")
        self.device_button = customtkinter.CTkButton(master=self.button_frame, text="Refresh", command=self.PopulateDevices)
        self.device_button.grid(row=1, column=1, padx=0, pady=10, sticky="ew")

        self.write_button = customtkinter.CTkButton(self.button_frame, text="Write Memory", command=self.writeMemory, height=50, width=150, state="disabled")
        self.write_button.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        self.clear_button = customtkinter.CTkButton(self.button_frame, text="Clear Memory", command=self.clearMemory, fg_color=("red"), height=50, width=150, hover_color=("darkred"), state="disabled")
        self.clear_button.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        self.text_box = customtkinter.CTkTextbox(master=self.button_frame, width=300, height=100)
        self.text_box.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # More Ari code
        # Populate initial values
        self.update_values()

    def update_values(self):
        self.cartSN.delete(0, tk.END)
        self.cartSN.insert(0, shared_data.get('cartridge_serial_number', ''))

        config = self.config_menu.get()

        # Updating Fields based on cartridge SN
        data = shared_data.get('data', {})
        if 'Rinse' in data and config != "CR800":
            self.rinse_pH.delete(0, tk.END)
            self.rinse_pH.insert(0, data['Rinse'][0])
            self.rinse_cond.delete(0, tk.END)
            self.rinse_cond.insert(0, data['Rinse'][1])
            self.rinse_Ca.delete(0, tk.END)
            self.rinse_Ca.insert(0, data['Rinse'][2])
            self.rinse_TH.delete(0, tk.END)
            self.rinse_TH.insert(0, data['Rinse'][3])
            self.rinse_NH4.delete(0, tk.END)
            self.rinse_NH4.insert(0, data['Rinse'][4])
            self.rinse_IS.delete(0, tk.END)
            self.rinse_IS.insert(0, data['Rinse'][5])
            self.rinse_KT.delete(0, tk.END)
            self.rinse_KT.insert(0, data['Rinse'][6])
            self.rinse_CondTComp.delete(0, tk.END)
            self.rinse_CondTComp.insert(0, data['Rinse'][7])
        if 'Clean' in data:
            self.clean_pH.delete(0, tk.END)
            self.clean_pH.insert(0, data['Clean'][0])
            self.clean_cond.delete(0, tk.END)
            self.clean_cond.insert(0, data['Clean'][1])
            self.clean_Ca.delete(0, tk.END)
            self.clean_Ca.insert(0, data['Clean'][2])
            self.clean_TH.delete(0, tk.END)
            self.clean_TH.insert(0, data['Clean'][3])
            self.clean_NH4.delete(0, tk.END)
            self.clean_NH4.insert(0, data['Clean'][4])
            self.clean_IS.delete(0, tk.END)
            self.clean_IS.insert(0, data['Clean'][5])
            self.clean_KT.delete(0, tk.END)
            self.clean_KT.insert(0, data['Clean'][6])
            self.clean_CondTComp.delete(0, tk.END)
            self.clean_CondTComp.insert(0, data['Clean'][7])
        if 'Cal 5' in data:
            self.cal_5_pH.delete(0, tk.END)
            self.cal_5_pH.insert(0, data['Cal 5'][0])
            self.cal_5_cond.delete(0, tk.END)
            self.cal_5_cond.insert(0, data['Cal 5'][1])
            self.cal_5_Ca.delete(0, tk.END)
            self.cal_5_Ca.insert(0, data['Cal 5'][2])
            self.cal_5_TH.delete(0, tk.END)
            self.cal_5_TH.insert(0, data['Cal 5'][3])
            self.cal_5_NH4.delete(0, tk.END)
            self.cal_5_NH4.insert(0, data['Cal 5'][4])
            self.cal_5_IS.delete(0, tk.END)
            self.cal_5_IS.insert(0, data['Cal 5'][5])
            self.cal_5_KT.delete(0, tk.END)
            self.cal_5_KT.insert(0, data['Cal 5'][6])
            self.cal_5_CondTComp.delete(0, tk.END)
            self.cal_5_CondTComp.insert(0, data['Cal 5'][7])
        if 'Cal 6' in data:
            self.cal_6_pH.delete(0, tk.END)
            self.cal_6_pH.insert(0, data['Cal 6'][0])
            self.cal_6_cond.delete(0, tk.END)
            self.cal_6_cond.insert(0, data['Cal 6'][1])
            self.cal_6_Ca.delete(0, tk.END)
            self.cal_6_Ca.insert(0, data['Cal 6'][2])
            self.cal_6_TH.delete(0, tk.END)
            self.cal_6_TH.insert(0, data['Cal 6'][3])
            self.cal_6_NH4.delete(0, tk.END)
            self.cal_6_NH4.insert(0, data['Cal 6'][4])
            self.cal_6_IS.delete(0, tk.END)
            self.cal_6_IS.insert(0, data['Cal 6'][5])
            self.cal_6_KT.delete(0, tk.END)
            self.cal_6_KT.insert(0, data['Cal 6'][6])
            self.cal_6_CondTComp.delete(0, tk.END)
            self.cal_6_CondTComp.insert(0, data['Cal 6'][7])
        if 'T1' in data and config != "CR1300" and config != "CR800":
            self.T1_HCl_N.delete(0, tk.END)
            self.T1_HCl_N.insert(0, data['T1'][0])
            # End more Ari code
            
    def writeMemory(self):
        dialog = CTkConfirmDialog(text="Are you sure?", title="Write Memory?")
        # try:
        if dialog.get_input():
            self.sensor_config_check.configure(text_color="white")
            self.cart_info_check.configure(text_color="white")
            self.dates_check.configure(text_color="white")
            self.clcal_check.configure(text_color="white")
            self.sol_info_check.configure(text_color="white")
            self.therm_check.configure(text_color="white")
            self.valve_check.configure(text_color="white")
            
            # self.label_status.configure(text="Working...", text_color="white")
            # print("Writing Memory...")
            count = 0
            if self.sensor_config_check.get():
                self.write("Writing Sensor configuration\n")
                count += 1
                if not self.WriteSensorConfig():
                    self.sensor_config_check.configure(text_color="red")
                else:
                    self.sensor_config_check.configure(text_color="green")
            
            if self.cart_info_check.get():
                self.write("Writing Cartridge Info\n")
                count += 1
                if self.ValidateCartridgeInfo():
                    if not self.WriteCartridgeInfo():
                        self.cart_info_check.configure(text_color="red")
                    else:
                        self.cart_info_check.configure(text_color="green")
                        
            if self.dates_check.get():
                self.write("Writing Dates\n")
                count += 1
                if not self.WriteDates():
                    self.dates_check.configure(text_color="red")
                else:
                    self.dates_check.configure(text_color="green")
                        
            if self.sol_info_check.get():
                self.write("Writing Solution Info\n")
                count += 1
                if not self.WriteSolutions():
                    self.sol_info_check.configure(text_color="red")
                else:
                    self.sol_info_check.configure(text_color="green")
            if self.clcal_check.get():
                self.write("Writing Cl Calibration\n")
                count += 1
                if not self.WriteClCal():
                    self.clcal_check.configure(text_color="red")
                else:
                    self.clcal_check.configure(text_color="green")
            if self.therm_check.get():
                self.write("Writing Thermistor Slope\n")
                count += 1
                if not self.WriteTherm():
                    self.therm_check.configure(text_color="red")
                else:
                    self.therm_check.configure(text_color="green")
            if self.valve_check.get():
                self.write("Writing Valve Setup\n")
                count += 1
                if not self.WriteValve():
                    self.valve_check.configure(text_color="red")
                else:
                    self.valve_check.configure(text_color="green")
            
            # self.label_status.configure(text="Memory Written!", text_color="green")
            
            # Check if there was at least one item was checked
            if count == 0:
                self.write(text="No items selected!\n")
        else:
            # self.label_status.configure(text="Cancelled", text_color="white")
            self.write("Cancelled\n")
            
        # except:
        #     print("Something went wrong!")
            
    def clearMemory(self):
        dialog = CTkConfirmDialog(text="Are you sure?", title="Clear Memory?")
        # try:
        if dialog.get_input():
            self.write("Clearing Memory...\n")
            com = self.device_dropdown.get().split(":")[0]
            
            if "COM" in com:
                # Open the serial port device identified by 
                SerialObj = serial.Serial(timeout=1)
                SerialObj.baudrate = 115200  # set Baud rate to 9600
                SerialObj.bytesize = 8   # Number of data bits = 8
                SerialObj.parity  ='N'   # No parity
                SerialObj.stopbits = 1   # Number of Stop bits = 1
                SerialObj.port = com
                SerialObj.open()
                
                SerialObj.write(b'M')    #transmit 'A' (8bit) to microcontroller
                # time.sleep(1)
                
                attempt = 0
                msg = ""
                while attempt < 5 and "clear memory" not in msg:
                    msg = str(SerialObj.readline().decode("iso-8859-1"))
                    print(msg)
                    if "Plug in sensor" in msg:
                        # self.label_status.configure(text_color="red", text="Failed, plug in sensor chip!")
                        self.write("Plug in sensor chip!\n")
                        return False
                    # else:
                    #     return False
                
                if attempt == 5:
                    # self.label_status.configure(text_color="red", text="Timed out!")
                    return False
                
                SerialObj.write(b'C')    #transmit 'A' (8bit) to microcontroller
                time.sleep(1)
                SerialObj.write(b'y')    #transmit 'A' (8bit) to microcontroller
                
                attempt = 0
                msg = ""
                while attempt < 520 and "Done!" not in msg:
                    msg = str(SerialObj.readline().decode("iso-8859-1"))
                    print(msg, end="")

                # if "Done!" in msg:
                    # self.label_status.configure(text_color="green", text="Memory Cleared!")
                    
                # UART_Rx = ""
                # while "!" not in UART_Rx:
                #     UART_Rx = str(SerialObj.read(1).decode("ascii"))
                #     print(UART_Rx, end="")
                
                SerialObj.close()      # Close the port
            else:
                self.write("Not an expected COM number\n")
            
        else:
            self.write("Cancelled\n")
        # except:
        #     print("Something went wrong!")
            
    def sensorConfigMenuChange(self, new_config: str):
        config = self.config_menu.get()
        self.maxdays.delete(0,len(self.maxdays.get()))
        self.maxtests.delete(0,len(self.maxtests.get()))
        self.maxcals.delete(0,len(self.maxcals.get()))

        if config == "CR800": # config = self.config_menu.get()
            self.maxdays.insert(0, 60)
            self.maxtests.insert(0, 150)
            self.maxcals.insert(0, 45)
            self.T1_frame.grid_remove()
            self.rinse_frame.grid_remove()
            self.expdate.delete(0, 'end')
            self.expdate.insert(0, get_future_date(60))
        else:
            self.maxdays.insert(0, 45)
            self.maxtests.insert(0, 100)
            self.maxcals.insert(0, 30)
            self.T1_frame.grid(row=1, column=4, padx=(5, 5), pady=(5, 5), sticky="nsew")
            self.rinse_frame.grid(row=1, column=0, padx=(5, 5), pady=(5, 5), sticky="nsew")
            self.expdate.delete(0, 'end')
            self.expdate.insert(0, get_future_date(45))
            
            
    def ValidateCartridgeInfo(self):
        passed = True
        # Validate Cartridge Info
        if len(self.cartSN.get()) != 7:
            self.label_cartSN.configure(text_color="red")
            passed = False
        else:
            self.label_cartSN.configure(text_color="white")
            if "H00" in self.cartSN.get() and (self.config_menu.get() == "CR800" or self.config_menu.get() == "CR1300"):
                dialog = CTkConfirmDialog(text="Serial Number doesn't match up with sensor configuration, please confirm", title="Sensor Config Check")
                if not dialog.get_input():
                    passed = False                
            if "H01" in self.cartSN.get() and self.config_menu.get() != "CR800":
                dialog = CTkConfirmDialog(text="Serial Number doesn't match up with sensor configuration, please confirm", title="Sensor Config Check")
                if not dialog.get_input():
                    passed = False
            if "H02" in self.cartSN.get() and self.config_menu.get() != "CR1300":
                dialog = CTkConfirmDialog(text="Serial Number doesn't match up with sensor configuration, please confirm", title="Sensor Config Check")
                if not dialog.get_input():
                    passed = False
            
        if len(self.sensorSN.get()) != 7:
            self.label_sensorSN.configure(text_color="red")
            passed = False
        else:
            self.label_sensorSN.configure(text_color="white")
            
        if int(self.maxdays.get()) < 0 or int(self.maxdays.get()) > 255:
            # self.label_maxdays.configure(text_color="red")
            passed = False
        # else:
        #     self.label_maxdays.configure(text_color="white")
        
        expdate_list = self.expdate.get().split("/")
        if len(expdate_list) != 3 or int(expdate_list[0]) < 1 or int(expdate_list[0]) > 12 or int(expdate_list[1]) < 1 or int(expdate_list[1]) > 31 or int(expdate_list[2]) < 20 or int(expdate_list[2]) > 38:
            self.label_expdate.configure(text_color="red")
            passed = False
        else:
            self.label_expdate.configure(text_color="white")
            
        # Check that expiration date is in the future
        if datetime.strptime(self.expdate.get(), "%m/%d/%y") < datetime.now():
            expiration_dialog = CTkConfirmDialog(text="Do you want an expiration date in the past?", title="Past Expiration Date")
            if not expiration_dialog.get_input():
                passed = False
            
        hyddate_list = self.hyddate.get().split("/")
        if len(hyddate_list) != 3 or int(hyddate_list[0]) < 1 or int(hyddate_list[0]) > 12 or int(hyddate_list[1]) < 1 or int(hyddate_list[1]) > 31 or int(hyddate_list[2]) < 20 or int(hyddate_list[2]) > 38:
            self.label_hyddate.configure(text_color="red")
            passed = False
        else:
            self.label_hyddate.configure(text_color="white")
            
        if datetime.strptime(self.expdate.get(), "%m/%d/%y") < datetime.strptime(self.hyddate.get(), "%m/%d/%y"):
            self.write("Hydration is after expiration!\n")
            print("Hydration is after expiration")
            passed = False
            
        if int(self.maxtests.get()) < 0 or int(self.maxtests.get()) > 255:
            self.label_maxtests.configure(text_color="red")
            passed = False
        else:
            self.label_maxtests.configure(text_color="white")
            
        if int(self.maxcals.get()) < 0 or int(self.maxcals.get()) > 255:
            self.label_maxcals.configure(text_color="red")
            passed = False
        else:
            self.label_maxcals.configure(text_color="white")
            
        if passed:
            self.cart_info_check.configure(text_color="white")
        else:
            self.cart_info_check.configure(text_color="red")
            
        return passed
    
    def WriteSensorConfig(self):
        try:
            self.write("Writing Sensor Config...\n")
            
            sensor_config = self.config_menu.get()
            # ["Across V7", "pH Cl Cart", "Disinfection: 2 Cr, 6 NH4, 2 Cr"]
            if sensor_config == "CR2300 (V7)":
                sensor_config_byte = bytearray([15])
            elif sensor_config == "CR800":
                sensor_config_byte = bytearray([1])
            elif sensor_config == "CR1300":
                sensor_config_byte = bytearray([24])
            elif sensor_config == "V11":
                sensor_config_byte = bytearray([25])
            elif sensor_config == "V12":
                sensor_config_byte = bytearray([26])
                
            success = self.WriteMemoryUART(sensor_config_byte, 1, 39)
            
            return success
        except:
            self.write("Write Sensor Config Failed\n")
            return False
        
    def WriteCartridgeInfo(self):
        try:
            self.write("Writing Cartridge Info...\n")
            ba = bytearray(self.cartSN.get().encode('iso-8859-1'))
            ba.extend(bytearray(self.sensorSN.get().encode('iso-8859-1')))
            ba.append((int(self.maxdays.get())))
            
            max_run_bytes = bytearray([int(self.maxtests.get()), 0, int(self.maxcals.get())])
            
            print([ "0x%02x" % b for b in ba ])
            print([ "0x%02x" % b for b in max_run_bytes ])
            
            success = True
            success = success and self.WriteMemoryUART(ba, 1, 0)
            success = success and self.WriteMemoryUART(max_run_bytes, 1, 33)
            
            return success
        except:
            self.write("Write Cartridge Info Failed\n")
            return False

    def WriteDates(self):
        try:
            self.write("Writing Dates...\n")
            expdate_list = self.expdate.get().split("/")
            date_bytes = bytearray([int(expdate_list[0]), int(expdate_list[1]), 20, int(expdate_list[2])])
            hyddate_list = self.hyddate.get().split("/")
            date_bytes.extend(bytearray([int(hyddate_list[0]), int(hyddate_list[1]), 20, int(hyddate_list[2])])) 
            
            success = self.WriteMemoryUART(date_bytes, 1, 24)
            
            return success
        except:
            self.write("Write Dates Failed\n")
            return False

    def WriteSolutions(self):
        config = self.config_menu.get()
        
        try:
            ba = bytearray()
            
            # ba = bytearray(struct.pack("f",float(self.T1_HCl_N.get())))
            
            if self.T1_frame.winfo_ismapped():
                ba.extend(bytearray(struct.pack("f",float(self.T1_HCl_N.get()))))
            else:
                ba.extend(bytearray(struct.pack("f",float(0))))
            
            if config == "CR800":
                ba.extend(bytearray(struct.pack("f",float(self.clean_pH.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.clean_NH4.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.clean_Ca.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.clean_TH.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.clean_cond.get()))))
            else:
                ba.extend(bytearray(struct.pack("f",float(self.rinse_pH.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.rinse_NH4.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.rinse_Ca.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.rinse_TH.get()))))
                ba.extend(bytearray(struct.pack("f",float(self.rinse_cond.get()))))
                
            ba.extend(bytearray(struct.pack("f",float(self.cal_6_pH.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_6_NH4.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_6_Ca.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_6_TH.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_6_cond.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.clean_Ca.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_5_pH.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_5_NH4.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_5_Ca.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_5_TH.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_5_cond.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.clean_pH.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.clean_NH4.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.clean_cond.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.clean_TH.get()))))
            if config == "CR800":
                ba.extend(bytearray(struct.pack("f",float(self.clean_IS.get()))))
            else:
                ba.extend(bytearray(struct.pack("f",float(self.rinse_IS.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.clean_IS.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_6_IS.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_5_IS.get()))))
            if config == "CR800":
                ba.extend(bytearray(struct.pack("f",float(self.clean_KT.get()))))
            else:
                ba.extend(bytearray(struct.pack("f",float(self.rinse_KT.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_6_KT.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.cal_5_KT.get()))))
            ba.extend(bytearray(struct.pack("f",float(0))))
            ba.extend(bytearray(struct.pack("f",float(self.clean_KT.get()))))
            
            if config == "CR800":
                page2 = bytearray(struct.pack("f",float(self.clean_CondTComp.get())))
            else:
                page2 = bytearray(struct.pack("f",float(self.rinse_CondTComp.get())))
            page2.extend(bytearray(struct.pack("f",float(self.cal_6_CondTComp.get()))))
            page2.extend(bytearray(struct.pack("f",float(self.cal_5_CondTComp.get()))))
            page2.extend(bytearray(struct.pack("f",float(self.clean_CondTComp.get()))))
            print([ "0x%02x" % b for b in ba ])
            print([ "0x%02x" % b for b in page2 ])
            self.sol_info_check.configure(text_color="white")
            
            success = True
            success = success and self.WriteMemoryUART(ba, 3, 0)
            success = success and self.WriteMemoryUART(page2, 3, 128)
            return success
            
        except:
            self.sol_info_check.configure(text_color="red")
            self.write("Can't convert solution value\n")
            return False
            

        
    def WriteClCal(self):
        try:
            ba = bytearray(struct.pack("f",float(self.tcl_low_slope.get())))
            ba.extend(bytearray(struct.pack("f",float(self.tcl_low_int.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.tcl_high_slope.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.tcl_high_int.get()))))
            
            ba.extend(bytearray(struct.pack("f",float(self.fcl_low_slope.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.fcl_low_int.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.fcl_high_slope.get()))))
            ba.extend(bytearray(struct.pack("f",float(self.fcl_high_int.get()))))

            print([ "0x%02x" % b for b in ba ])
            self.clcal_check.configure(text_color="white")
            
            success = self.WriteMemoryUART(ba, 2, 20)
            return success
            
        except:
            self.clcal_check.configure(text_color="red")
            self.write("Can't write Cl calibration\n")
            return False
                
        

    def WriteTherm(self):
        try:
            ba = bytearray(struct.pack("f",float(self.therm.get())))

            print([ "0x%02x" % b for b in ba ])
            self.therm_check.configure(text_color="white")
            
            success = self.WriteMemoryUART(ba, 2, 56)
            return success
            
        except:
            self.therm_check.configure(text_color="red")
            self.write("Can't write therm slope\n")
            return False
        
            
    def WriteValve(self):
        try:
            if self.valve.get() == "V1 Normal":
                ba = bytearray([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
            elif self.valve.get() == "V2 Alternate":
                ba = bytearray([10, 1, 3, 4, 5, 8, 9, 6, 2, 7])
            else:
                raise Exception("Valve not recognized")
    
            print([ "0x%02x" % b for b in ba ])
            self.valve_check.configure(text_color="white")
            
            success = self.WriteMemoryUART(ba, 1, 128)
            return success
            
        except:
            self.valve_check.configure(text_color="red")
            self.write("Can't write valve setup\n")
            return False
                
        

    def WriteMemoryUART(self, ba, page, offset):
        crc_calc = binascii.crc32(ba)
        print("CRC: {}".format(crc_calc))
        com = self.device_dropdown.get().split(":")[0]
        
        if "COM" in com:
            # Open the serial port device identified by 
            SerialObj = serial.Serial(timeout=0.5)
            SerialObj.baudrate = 115200  # set Baud rate to 115200
            SerialObj.bytesize = 8   # Number of data bits = 8
            SerialObj.parity  ='N'   # No parity
            SerialObj.stopbits = 1   # Number of Stop bits = 1
            SerialObj.port = com
            SerialObj.open()
            
            try:    # Write to the device, if it doesn't work still close the serial port
                #               
                # Send the first M command to tell the ROAM it is writing the memory
                #
                SerialObj.write(b'M')    #transmit 'A' (8bit) to microcontroller
                time.sleep(0.025)
                attempt = 0
                msg = ""
                while attempt < 5 and "clear memory" not in msg:
                    msg = str(SerialObj.readline().decode("iso-8859-1"))
                    print(msg)
                    attempt += 1
                    if "Plug in sensor" in msg:
                        # self.label_status.configure(text_color="red", text="Failed, plug in sensor chip!")
                        self.write("Plug in sensor chip!\n")
                        return False
                
                if attempt == 5:
                    # self.label_status.configure(text_color="red", text="Timed out!")
                    self.write("Timed out!\n")
                    SerialObj.close()      # Close the port
                    return False
            
                #
                # Send the second M command to tell the ROAM it is in memory streaming mode
                #
                SerialObj.write(b'M')    #transmit 'A' (8bit) to microcontroller
                time.sleep(0.025)
                command = bytearray([len(ba), page, 0, offset]) # Length of SNs and Max days, page 1, offset 0
                command.extend(crc_calc.to_bytes(4, byteorder="little", signed=False))
                
                SerialObj.write(command)    # Send inital command (length, page, offset, CRC)
                
                UART_Rx = ""
                
                timeout = time.time() + 3
                while "Ready!" not in UART_Rx:
                    try:
                        UART_Rx += str(SerialObj.read(1).decode("iso-8859-1"))
                        print(UART_Rx[-1], end="")
                    except UnicodeDecodeError:
                        pass
                    if time.time() > timeout:
                        print("Python timeout")
                        raise TimeoutError
                    
                
                    
                SerialObj.write(ba)    # Send the data to be saved to the memory
                
                UART_Rx = ""
                timeout = time.time() + 15
                while "Done!" not in UART_Rx:
                    try:
                        UART_Rx += str(SerialObj.read(1).decode("iso-8859-1"))
                        print(UART_Rx[-1], end="")
                    except UnicodeDecodeError:
                        pass
                    if time.time() > timeout:
                        print("Python timeout")
                        raise TimeoutError
                    
                msg = ""
                attempt = 0
                while "Pass" not in msg and "Fail" not in msg and attempt < 5:
                    msg = str(SerialObj.readline().decode("iso-8859-1"))
                    attempt += 1
                
                print(msg)
                    
                SerialObj.close()      # Close the port
                if "Pass" in msg:
                    return True
                else:
                    return False
            except:
                SerialObj.close()      # Close the port
                print("Error writing UART")
                return False
        else:
            self.write("Not an expected COM number\n")
            return False

    def PopulateDevices(self):
        self.write("Looking for devices...\n")
        # self.label_status.configure(text="Looking for devices...", text_color="white")
        ports = serial.tools.list_ports.comports()
        ROAM_ports = []
        dropdown_options = []

        SerialObj = serial.Serial(write_timeout=1,timeout=1)
        SerialObj.baudrate = 115200  # set Baud rate to 9600
        SerialObj.bytesize = 8   # Number of data bits = 8
        SerialObj.parity  ='N'   # No parity
        SerialObj.stopbits = 1   # Number of Stop bits = 1

        for port in ports:
            
            if "UART" in port.description or "Stellaris" in port.description or "Prolific" in port.description or "USB Serial Device" in port.description or "com0com" in port.description:
                # SerialObj = serial.Serial('COM4') # COMxx  format on Windows
                # SerialObj = serial.Serial(port.device) # COMxx  format on Windows

                SerialObj.port = port.device
                
                try:
                    SerialObj.open() # Write timeout 1 second
                    # time.sleep(3)
                    SerialObj.write(b'0')    #transmit '0' (8bit) to microcontroller
                    time.sleep(0.25)
                    bytesToRead = SerialObj.inWaiting()
                    if bytesToRead == 0:
                        # print("{} didn't respond".format(port.device))
                        self.write("{} didn't respond\n".format(port.device))
                    else:
                        # ROAM_SN = (str(SerialObj.read(bytesToRead).decode("ascii"))[6:10])
                        ROAM_SN = (str(SerialObj.read(bytesToRead).decode("iso-8859-1"))[-6:-1])
                        self.write(port.device + ": ROAM " + ROAM_SN + "\n")
                        ROAM_ports.append([port.device, ROAM_SN])
                        dropdown_options.append(str(port.device + ": ROAM " + ROAM_SN))
                    SerialObj.close()      # Close the port
                except serial.serialutil.SerialException:
                    dropdown_options.append(str(port.device + ": Busy"))
                    self.write(port.device + ": couldn't open serial port")
                except UnicodeDecodeError:
                    dropdown_options.append(str(port.device + ": Busy"))
                    self.write(port.device + ": received non-unicode characters")
                
                if SerialObj.is_open:
                    SerialObj.close()      # Close the port
                
        # for port, SN in ROAM_ports:
        #     print(port)
        #     print(SN)
        
        self.device_dropdown.configure(values = dropdown_options)
        if len(dropdown_options) > 0:
            self.device_dropdown.configure(values = dropdown_options, state="normal")
            self.device_dropdown.set(dropdown_options[-1])
            self.clear_button.configure(state="normal")
            self.write_button.configure(state="normal")
            # self.label_status.configure(text="Found devices!", text_color="white")
            # self.write("Found devices!\n")
        else:
            self.device_dropdown.configure(values = dropdown_options, state="disabled")
            self.device_dropdown.set("No devices found")
            self.clear_button.configure(state="disabled")
            self.write_button.configure(state="disabled")
            # self.label_status.configure(text="No devices found!", text_color="red")
            self.write("No devices found!\n")
        
    def write(self, text):
        self.text_box.insert(tk.END, text)
        self.text_box.see(tk.END)  # Auto-scroll to the bottom

# Run app
if __name__ == "__main__":
    app = App()
    app.mainloop()

#C:\Users\Ari>pyinstaller --onefile --add-data "C:\Users\Ari\anaconda3\envs\pyinstaller_test\Lib\site-packages/customtkinter;customtkinter/" "WriteMemoryExe_V5.py"
# pyinstaller --onefile --add-data "C:\Users\Jason\anaconda3\envs\pyinstaller_test\Lib\site-packages/customtkinter;customtkinter/" "WriteMemoryExe.py"
# To build the executable use the anaconda prompt, activate the conda environment with pyinstaller (conda activate pyinstaller_test)
# cd to where this script is saved then run the following command
# pyinstaller --onefile --add-data "C:\Users\ari\anaconda3\envs\pyinstaller\Lib\site-packages/customtkinter;customtkinter/" "WriteMemoryExe_V5.py"
# Note to ari: you have to navigate to your "pyinstaller" environment before running the code to make the executables
