import os
import sys
import re
import json
import serial
import threading
import time
import struct
import random
import pyautogui
import xlsxwriter
import openpyxl
import binascii
import webbrowser
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from tkinter import filedialog
import serial.tools.list_ports
from datetime import datetime
from tkinter import messagebox
from PIL import Image, ImageTk
from tkinter import simpledialog, messagebox


# get the path to the script directory
script_dir = sys.path[0]

class MainGui:
    def __init__(self, window):
        self.root = root
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        data = 'full_config.json'
        filename = self.get_json_file_path(data)
        self.uartState = False
        self.is_running = False
        self.is_action_in_progress = False
        self.start_time = time.time()
        self.data = None
        self.entry_lists1 = []
        self.flag = 0
        self.flag2 = 0
        self.open = 0
        self.ser = None
        self.le = 0
        self.ser = serial.Serial()
        self.selected_port = tk.StringVar()
        
        def resource_path(relative_path):
            """ Get the absolute path to the resource, works for dev and for PyInstaller """
            base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            return os.path.join(base_path, relative_path)

        icon_path = resource_path('aptech.ico')  # Path to your .ico file
        root.iconbitmap(icon_path)

        # Set modern theme and appearance
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        self.setup_gui()
        
        # Serial object
        self.ser = serial.Serial()
        self.root.mainloop()

    def setup_gui(self):
        # Main window configuration
        self.root.title("Tuner Software V6.0")
        self.root.resizable(True, True)
        self.root.configure(fg_color="white")
        
        # Set window size and position
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.85)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Create main container with modern styling
        self.main_container = ctk.CTkFrame(self.root, fg_color="white")
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        #create header
        header_frame = ctk.CTkFrame(self.main_container, fg_color="white")
        header_frame.pack(fill=tk.X, padx=5, pady=(0, 10))
        
        # Logo
        img_path = self.get_json_file_path('APT_LOGO.png')
        try:
            image = Image.open(img_path)
            width, height = 500, 45 
            resized_image = image.resize((width, height), Image.LANCZOS)
            img = ctk.CTkImage(dark_image=resized_image, size=(width, height))
            logo_label = ctk.CTkLabel(header_frame, image=img, text="")
            logo_label.pack(side=tk.LEFT, padx=10)
        except Exception as e:
            print(f"Error loading logo: {e}")
        
        # Title and version
        title_frame = ctk.CTkFrame(header_frame, fg_color="white")
        title_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        title_label = ctk.CTkLabel(title_frame, text="⚡TUNER SOFTWARE", font=('Helvetica', 14, 'bold'), text_color="#2a5885")
        title_label.pack(side=tk.TOP, anchor='w', padx=15)
        
        version_label = ctk.CTkLabel(title_frame, text="  Version 6.0", font=('Helvetica', 13, 'bold'),text_color="#666666")
        version_label.pack(side=tk.TOP, anchor='w', padx=15)
        
        # Help button
        help_btn = ctk.CTkButton(header_frame, text="?", width=30, height=30,
                               font=('Helvetica', 14, 'bold'), fg_color="#2a5885", hover_color="#3b7cb1", command=self.show_help)
        help_btn.pack(side=tk.RIGHT, padx=10)

        #Control panel
        control_frame = ctk.CTkFrame(self.main_container, border_width=1, 
                                   border_color="#dddddd", corner_radius=8, fg_color="#f8f9fa")
        control_frame.pack(fill=tk.X, padx=5, pady=5, ipady=10)
        
        # Left section - COM Port controls
        port_frame = ctk.CTkFrame(control_frame, fg_color="transparent", border_color='black', border_width=2)
        port_frame.pack(side=tk.LEFT, padx=(15, 2), pady=5, ipadx=20, ipady=8)
        
        ctk.CTkLabel(port_frame, text="SERIAL PORT", font=('Helvetica', 14),
                   text_color="#333333").pack(padx=(10, 1), pady=2, anchor='w')
        
        # COM port selection
        serial_ports = self.detect_serial_ports()
        self.selected_port = tk.StringVar(value="Select Port")
        
        port_combo = ctk.CTkComboBox(port_frame, variable=self.selected_port, values=serial_ports, width=190, dropdown_fg_color="#f8f9fa", 
                                   dropdown_text_color="#333333", button_color="#2a5885", border_color="#cccccc")
        port_combo.pack(pady=(5, 0))
        self.selected_port.trace_add("write", self.on_port_selected)
        
        # Connection buttons
        btn_frame = ctk.CTkFrame(port_frame, fg_color="transparent")
        btn_frame.pack(pady=(10, 0))
        
        self.buttonSS = ctk.CTkButton(btn_frame, text="CONNECT", width=95, font=('Helvetica', 13, 'bold'),
                                    fg_color="#28a745", hover_color="#218838", state="disabled", command=self.processButtonSS)
        self.buttonSS.pack(side=tk.LEFT, padx=5)
        
        self.buttonStop = ctk.CTkButton(btn_frame, text="DISCONNECT", width=85, font=('Helvetica', 13, 'bold'),
                                      fg_color="#dc3545", hover_color="#c82333", state="disabled", command=self.stopButtonSS)
        self.buttonStop.pack(side=tk.LEFT)
        
        # Middle section - PIC ID
        pic_frame = ctk.CTkFrame(control_frame, fg_color="transparent", border_color='black', border_width=2)
        pic_frame.pack(side=tk.LEFT, padx=20, pady=5, ipadx=10, ipady=7)
        
        ctk.CTkLabel(pic_frame, text="PIC ID", font=('Helvetica', 15, 'bold'), text_color="#333333").pack(padx=20, pady=2)
        
        id_entry_frame = ctk.CTkFrame(pic_frame, fg_color="transparent")
        id_entry_frame.pack(pady=(5, 0))
        
        ctk.CTkLabel(id_entry_frame, text="High", font=('Helvetica', 11, 'bold')).grid(row=0, column=1, padx=(0, 5))
        self.label2 = ctk.CTkEntry(id_entry_frame, width=100, height=35, border_color="black", font=('Helvetica', 11))
        self.label2.grid(row=1, column=1, padx=(0, 15))
        
        ctk.CTkLabel(id_entry_frame, text="Low", font=('Helvetica', 11, 'bold')).grid(row=0, column=2, padx=(0, 5))
        self.label3 = ctk.CTkEntry(id_entry_frame, width=100, height=35, border_color="black", font=('Helvetica', 11))
        self.label3.grid(row=1, column=2)
        
        # Right section - Status
        status_frame = ctk.CTkFrame(control_frame, fg_color="transparent", border_color='black', border_width=2)
        status_frame.pack(side=tk.LEFT, padx=25, pady=5, ipadx=60, ipady=4)
        
        ctk.CTkLabel(status_frame, text="STATUS", font=('Helvetica', 15, 'bold'), text_color="#333333").pack(padx=20, ipadx=60, pady=15)
        
        self.dialog = ctk.CTkLabel(status_frame, text="", font=('Helvetica', 11), anchor="center", justify="center", width=200, wraplength=400)
        self.dialog.pack(pady=(5, 20))
        
        # Firmware version section
        firmware_frame = ctk.CTkFrame(control_frame, fg_color="transparent", border_color='black', border_width=2)
        firmware_frame.pack(side=tk.LEFT, padx=20, pady=5, ipadx=10, ipady=7)
        
        ctk.CTkLabel(firmware_frame, text="FIRMWARE VERSION", font=('Helvetica', 14, 'bold'), text_color="#333333").pack(pady=2)
        
        fw_entry_frame = ctk.CTkFrame(firmware_frame, fg_color="transparent")
        fw_entry_frame.pack(pady=(5, 0))
        
        ctk.CTkLabel(fw_entry_frame, text="High", font=('Helvetica', 11, 'bold')).grid(row=0, column=1, padx=(0, 5))
        self.firmware2 = ctk.CTkEntry(fw_entry_frame, width=100, height=35, border_color="black", font=('Helvetica', 11))
        self.firmware2.grid(row=2, column=1, padx=(0, 15))
        
        ctk.CTkLabel(fw_entry_frame, text="Low", font=('Helvetica', 11, 'bold')).grid(row=0, column=3, padx=(0, 5))
        self.firmware3 = ctk.CTkEntry(fw_entry_frame, width=100, height=35, border_color="black", font=('Helvetica', 11))
        self.firmware3.grid(row=2, column=3)

        # Color code indicators
        color_frame = ctk.CTkFrame(control_frame, fg_color="transparent", border_color='black', border_width=2)
        color_frame.pack(side=tk.RIGHT, padx=10, pady=(30, 1), ipadx=3, ipady=5)

        ctk.CTkLabel(color_frame, text="Color Code", font=('Helvetica', 14, 'bold'), text_color="#333333").pack(pady=2)
        
        ctk.CTkLabel(color_frame, text="Read", font=('Helvetica', 12, 'bold'), 
                   fg_color="#90EE90", corner_radius=7, width=50).pack(side=tk.LEFT, padx=4, ipadx=10)
        ctk.CTkLabel(color_frame, text="Write", font=('Helvetica', 12, 'bold'), 
                   fg_color="#ADD8E6", corner_radius=7, width=50).pack(side=tk.LEFT, padx=4, ipadx=10)

    
        # Create a frame for the configuration area
        config_frame = ctk.CTkFrame(self.main_container, border_width=2, border_color="#dddddd", corner_radius=8, fg_color="#ffffff")
        config_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
        
        # Create a canvas and scrollbars
        self.config_canvas = tk.Canvas(config_frame, bg="white", highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(config_frame, orient="vertical", command=self.config_canvas.yview)
        h_scrollbar = ttk.Scrollbar(config_frame, orient="horizontal", command=self.config_canvas.xview)
        
        # Create scrollable frame
        self.scrollable_frame = ctk.CTkFrame(self.config_canvas, fg_color="white")
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.config_canvas.configure(scrollregion=self.config_canvas.bbox("all")))
        
        # Configure canvas
        self.config_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.config_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout for scrollable area
        self.config_canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        config_frame.grid_rowconfigure(0, weight=1)
        config_frame.grid_columnconfigure(0, weight=1)
        
        # Mouse wheel bindings
        self.config_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.config_canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)
        
        # Load configuration fields
        self.entry_lists1 = self.commonRead(self.scrollable_frame, 'full_config.json', "Full Configuration")

        #create_action_buttons
        action_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        action_frame.pack(padx=10, pady=(5, 10))
        
        # Button styles
        btn_style = {'font': ('Helvetica', 14, 'bold'), 'width': 160, 'height': 42, 'corner_radius': 6, 'fg_color': "royalblue", 'hover_color': "indigo", 'text_color': "white"}
        
        # Create buttons
        self.button_write = ctk.CTkButton(action_frame, text="Write", command=self.processButtonSend, **btn_style)
        self.button_write.pack(side=tk.LEFT, padx=10)
        
        self.button_read = ctk.CTkButton(action_frame, text="Read", command=self.processButtonReceive, **btn_style)
        self.button_read.pack(side=tk.LEFT, padx=10)
        
        self.button_upload = ctk.CTkButton(action_frame, text="Upload", command=self.open_file, **btn_style)
        self.button_upload.pack(side=tk.LEFT, padx=10)
        
        self.button_print = ctk.CTkButton(action_frame, text="Print", command=self.write_excelbutton, **btn_style)
        self.button_print.pack(side=tk.LEFT, padx=10)
        
        self.button_zero = ctk.CTkButton(action_frame, text="Zero Angle Estimate", command=self.start_continuous, **btn_style)
        self.button_zero.pack(side=tk.LEFT, padx=10)
        
        self.button_screenshot = ctk.CTkButton(action_frame, text="Screenshot", command=self.capture_screenshot, **btn_style)
        self.button_screenshot.pack(side=tk.LEFT, padx=10)


    #star continuous
    def start_continuous(self):
        def task():
            self.button_zero.configure(state='disabled', fg_color="#6c757d")
            try:
                self.receiveRandom()  # Assuming this is a blocking call
            finally:
                self.button_zero.configure(state='normal', fg_color="#2a5885")

        thread = threading.Thread(target=task)
        thread.start()

    #capture screenshot
    def capture_screenshot(self):
        # Ask user where to save with default filename
        default_filename = f"tuner_screenshot_{time.strftime('%Y_%m_%d_%H-%M-%S')}.png"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png", 
            filetypes=[("PNG files", "*.png")], 
            title="Save Screenshot As",
            initialfile=default_filename)
        
        if file_path:
            # Get window coordinates
            x = self.root.winfo_rootx()
            y = self.root.winfo_rooty()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
                
            # Capture and save
            screenshot = pyautogui.screenshot(region=(x, y, width, height))
            screenshot.save(file_path)


        # Setup button hover animations
        def animate_button(button, color_from, color_to):
            current_color = button.cget("fg_color")
            if current_color == color_from:
                button.configure(fg_color=color_to)
            else:
                button.configure(fg_color=color_from)
        
        # Bind animation to buttons
        buttons = [self.button_write, self.button_read, self.button_upload, 
                  self.button_print, self.button_zero, self.button_screenshot]
        
        for btn in buttons:
            btn.bind("<Enter>", lambda e, b=btn: animate_button(b, "#2a5885", "#3b7cb1"))
            btn.bind("<Leave>", lambda e, b=btn: animate_button(b, "#3b7cb1", "#2a5885"))


    def show_help(self):
        self.help_window = ctk.CTkToplevel(self.root)
        self.help_window.title("Help")
        self.help_window.geometry("500x400")
        self.help_window.resizable(False, False)
        self.help_window.grab_set()
        self.help_window.focus()
        self.help_window.transient(self.root)
    
        help_text = """
        Tuner Software V3.9 Help
    
        1. COM Port: Select the serial port your device is connected to
        2. Connect/Disconnect: Establish or terminate connection
        3. Read: Read current configuration from device
        4. Write: Send configuration to device
        5. Upload: Load configuration from file
        6. Print: Save current configuration to file
        7. Zero Angle: Perform zero angle estimation
        8. Screenshot: Capture application window
    
        For more help, visit our documentation.
        """

        text_frame = ctk.CTkFrame(self.help_window, fg_color="white")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
        help_label = ctk.CTkLabel(text_frame, text=help_text,
                                  font=('Helvetica', 12), justify='left', anchor='w', wraplength=480)
        help_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
        doc_link = ctk.CTkLabel(text_frame, text="Open Documentation", text_color="blue",
                                cursor="hand2", font=('Helvetica', 12, 'underline'))
        doc_link.pack(pady=(0, 10))
        doc_link.bind("<Button-1>", lambda e: webbrowser.open("https://docs.example.com"))

                

    def detect_serial_ports(self):
        self.ports = serial.tools.list_ports.comports()
        if self.ports:
            return [port.device for port in serial.tools.list_ports.comports()]
        else:
            return ['No ports available']


    # reading the file inside the program
    def get_json_file_path(self, filename):
        return os.path.join(getattr(sys, '_MEIPASS', script_dir), filename)

    # reading the xlsx side to entry list
    def open_file(self):
        filename = filedialog.askopenfilename(title="Open a File", filetype=(("xlxs files", ".*xlsx"),
                                                                             ("All Files", "*.")))
        if filename:
            try:
                filename = r"{}".format(filename)
                df = openpyxl.load_workbook(filename)
                sheet = df.active
                max_row = sheet.max_row - 2
                rE = 0
                le = 0
            
                # Iterate over rows up to 10
                for row in sheet.iter_rows(min_row=1, max_row=max_row, max_col=8):
                    row_data = []
                
                    # Iterate over cells in the row
                    for cell in row: 
                        if cell.column_letter == "B" or cell.column_letter == "D" or cell.column_letter == "F" or cell.column_letter == "H":
                            row_data.append(cell.value)

                    # reading the data and values
                    if le < len(self.entry_lists1):
                        self.AlgoToSetValue(self.entry_lists1[le], row_data)
                    else:
                        print(f"Warning: entry_lists1 index {le} is out of range.")
                    rE += 1
                    le += 1

                    # If the file is successfully processed, show "Upload successful"
                self.dialog.configure(text="Upload successful", font=('Helvetica', 17, 'bold'), fg_color="#ADD8E6")

            except ValueError:
                self.dialog.configure(text="File could not be opened")
        
            except FileNotFoundError:
                self.dialog.configure(text="File Not Found")

    # showing the random data that has to be received from the code
    def receiveRandom(self):
        rs = None
        self.is_running = True
        start_time = time.time()
        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)
        if self.uartState:
            while self.is_running:
                rs = self.ser.read(20)
                if time.time() - start_time > 12:
                    break
                if rs:
                    self.randomVari(rs, self.entry_lists1[2])
                    break

            if self.is_running:
                self.is_running = False
                self.button3['text'] = "Zero Angle Estimate"
            else:
                self.button3['text'] = "Zero Angle Estimate"

        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)

    # reading the picId
    def readFrame(self):
        if self.uartState:
                # Read PIC ID
                self.ser.write(bytes.fromhex('20012F'))
                bval = self.ser.read(10)
                
                if len(bval) < 10:
                    self.dialog.configure(text="Incomplete PIC ID response", font=('Helvetica', 14, 'bold'))
                    return

                # Check if the expected response length is reached
                if len(bval) >= 10:
                    # Validate the first and last bytes
                    first_byte = bval[0]
                    last_byte = bval[-1]

                    if bval.hex().startswith("20") and len(bval) == 10:
                        if first_byte == 0x20 and last_byte == 0x2f:
                            self.flag = 1
                        else:
                            print(f"Received PIC ID response is Invalid: {bval.hex()}")

                unpacked_data = bval[1:-1]
                h1 = (unpacked_data[:4]).hex()
                h1 = h1[::-1]
                st1 = h1[1] + h1[0] + h1[3] + h1[2] + h1[5] + h1[4] + h1[7] + h1[6]
                h2 = (unpacked_data[4:]).hex()
                h2 = h2[::-1]
                st2 = h2[1] + h2[0] + h2[3] + h2[2] + h2[5] + h2[4] + h2[7] + h2[6]

                formatted_pic_id = st2.zfill(4) + " " + st1.zfill(4)

                # update (self.label2) with h1 value
                self.label2.configure(state='normal')
                self.label2.delete(0, tk.END)
                self.label2.insert(0, st1)
                self.label2.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')

                # update (self.label3) with h2 value
                self.label3.configure(state='normal')
                self.label3.delete(0, tk.END)
                self.label3.insert(0, st2)
                self.label3.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')

                # Small delay before sending the next command
                time.sleep(0.1)

                # Read firmware version
                self.ser.timeout = 1
                self.ser.write(bytes.fromhex('30013F'))
                bval_fw = self.ser.read(10)
                
                if len(bval_fw) < 10:
                    self.dialog.configure(text="Incomplete Firmware", font=('Helvetica', 14, 'bold'))
                    return

                # Check if the expected response length is reached
                if len(bval_fw) >= 10:
                    # Validate the first and last bytes
                    first_byte = bval_fw[0]
                    last_byte = bval_fw[-1]

                    if bval_fw.hex().startswith("30") and len(bval_fw) == 10:
                        if first_byte == 0x30 and last_byte == 0x3f:
                            self.flag = 1
                        else:
                            print(f"Received Firmware ID response is Invalid: {bval_fw.hex()}")

                unpacked_data_fw = bval_fw[1:-1]
                #print(f"Unpacked firmware version data: {unpacked_data_fw.hex()}")
                fw1 = (unpacked_data_fw[:4]).hex()
                fw1 = fw1[::-1]
                st_fw1 = fw1[1] + fw1[0] + fw1[3] + fw1[2] + fw1[5] + fw1[4] + fw1[7] + fw1[6]
                fw2 = (unpacked_data_fw[4:]).hex()
                fw2 = fw2[::-1]
                st_fw2 = fw2[1] + fw2[0] + fw2[3] + fw2[2] + fw2[5] + fw2[4] + fw2[7] + fw2[6]

                formatted_fw_version = st_fw2.zfill(4) + " " + st_fw1.zfill(4)

                # update firmware labels
                self.firmware2.configure(state='normal')
                self.firmware2.delete(0, tk.END)
                self.firmware2.insert(0, st_fw1)
                self.firmware2.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')

                self.firmware3.configure(state='normal')
                self.firmware3.delete(0, tk.END)
                self.firmware3.insert(0, st_fw2)
                self.firmware3.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')                
        else:
            self.dialog.configure(text="Not in connect", font=('Helvetica', 14, 'bold'), text_color='red')


    # for connecting the process button with port
    def processButtonSS(self):
        if not self.uartState:
            self.ser.port = self.selected_port.get()
            self.ser.baudrate = 9600
            self.ser.timeout = 1
            try:
                self.ser.open()
                #print(f"Serial port {self.ser.port} opened")
            except:
                self.dialog.configure(
                    text="Can't open the Port", font=('Helvetica', 17, 'bold'), text_color="red")
        
            if self.ser.is_open:
                self.buttonSS.configure(fg_color="lightsteelblue")  # Set clicked color
                self.buttonSS.update_idletasks()
                self.dump()
                print("Dump command sent")
            
                if self.open == 1:  # open success
                    self.uartState = True

                    if self.flag2 == 0:
                        self.dialog.configure(text="Busy", font=('Helvetica', 17, 'bold'), text_color="blue")
                
                    #reading from the serial port
                    self.serial_operations()

                    #disabled the start button
                    self.buttonSS.update_idletasks()
                    self.disable_button(self.buttonSS)
                    self.enable_button(self.buttonStop)

    def stopButtonSS(self):
        if self.uartState:
            self.ser.close()
            self.buttonSS["text"] = "START"
            self.uartState = False
            print("Serial port closed")
            self.dialog.configure(text="Port Closed", font=('Helvetica', 17, 'bold'), bg_color='red', text_color='white')

            #disabled the STOP button after its click
            self.disable_button(self.buttonStop)
            self.enable_button(self.buttonSS)

    def serial_operations(self):
        self.processReceived()
        self.readFrame()
    
        self.flag2 = 1
        if self.flag2 == 1:
            self.dialog.configure(
                text="Read Successful", font=('Helvetica', 17, 'bold'), text_color='black')
            self.flag2 = 0
            self.open = 0

    # Assuming you have these functions defined elsewhere
    def on_enter(self, button, color):
        button.configure(fg_color=color)

    def on_leave(self, button, color):
        button.configure(fg_color=color)

    # Disable button function 
    def disable_button(self, button):
        button.configure(state="disabled", fg_color="grey", text_color="darkgrey")
        button.unbind("<Enter>")  # disable hover effect
        button.unbind("<Leave>")

    # Enable button function 
    def enable_button(self, button):
        if button == self.buttonSS:
            button.configure(state="normal", fg_color="royalblue", text_color="white")
            self.buttonSS.bind("<Enter>", lambda event, button=self.buttonSS: self.on_enter(button, "blue"))
            self.buttonSS.bind("<Leave>", lambda event, button=self.buttonSS: self.on_leave(button, "royalblue"))
        elif button == self.buttonStop:
            button.configure(state="normal", fg_color="red", text_color="white")
            self.buttonStop.bind("<Enter>", lambda event, button=self.buttonStop: self.on_enter(button, 'red2'))
            self.buttonStop.bind("<Leave>", lambda event, button=self.buttonStop: self.on_leave(button, 'red'))

    def on_port_selected(self, *args):
        selected_port = self.selected_port.get()
        if selected_port != "None":  
            self.buttonSS.configure(state="normal", text_color="white")
            self.dialog.configure(text="  Serial Port Connected  ", font=('Helvetica', 17, 'bold'), fg_color="#ADD8E6", text_color='black')
        else:  
            self.buttonSS.configure(fg_color="grey", text_color="white")


    def dump(self):
        start_time = time.time()
        self.ser.write(bytes.fromhex('10011f'))
    
        response = None
        while True:
            if self.ser.in_waiting:
                response = self.ser.read(3)
                break
            if time.time() - start_time > self.ser.timeout:  # Timeout of 5 seconds
                # print("no response")
                break

        # process the response
        if response:
            self.open = 1
        else:
            self.dialog.configure(text="Connection Error", font=('Helvetica', 17, 'bold'), fg_color='red', text_color='white')

    # stopping the function
    def stop_receiveRandom(self):
        self.is_running = False

    # for receiving the values from the uart
    def processButtonReceive(self):
        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)
        
        if self.uartState:
            if self.flag2 == 0:
                self.dialog.configure(text="Busy", font=('Helvetica', 17, 'bold'), fg_color='orange')

            self.is_running = False
            self.dialog.configure(text="Reading...", font=('Helvetica', 17, 'bold'), text_color='white', fg_color='indigo')

            value = []
            for data in self.entry_lists1:
                for entry in data:
                    if isinstance(entry, str):
                        value.append(entry[2:]+"00")

            self.received_data(value, self.entry_lists1)

            if self.flag2 == 1:
                self.dialog.configure(text="  Read Successful  ", font=('Helvetica', 17, 'bold'), fg_color="#90EE90", text_color='black')
                self.flag2 = 0
            else:
                self.dialog.configure(text="  Connection Error  ", font=('Helvetica', 17, 'bold'), fg_color='red', text_color='white')

        else:
            self.dialog.configure(text="  Read-> Not in Connect  ", font=('Helvetica', 17, 'bold'))

        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)

    # receving the data if the button is pressed
    def processReceived(self):
        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)
        #
        if (self.uartState):
            self.dialog.configure(text="  Busy  ", font=('Helvetica', 17, 'bold'))
            value = []
            for data in self.entry_lists1:
                for entry in data:
                    if isinstance(entry, str):
                        value.append(entry[2:]+"00")

            self.received_data(value, self.entry_lists1)

            if(self.flag2 == 1):
                self.dialog.configure(text="  Read Successful  ", font=('Helvetica', 17, 'bold'), fg_color="#90EE90", text_color='black')
                self.flag2 = 0
        else:
            self.dialog.configure(text="  Received-> Not In Connect  ")

        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)

    # sending the data to the uart
    def processButtonSend(self):
        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)
            
        if (self.uartState):
            self.dialog.configure(text="  Busy  ", font=('Helvetica', 17, 'bold'))

            self.is_running = False
            self.readEntry(self.entry_lists1)

            if self.flag == 0:
                self.dialog.configure(text="  Connection Error  ",  font=('Helvetica', 17, 'bold'), fg_color="red", text_color='white')
            else:
                self.dialog.configure(text="  Write Successful  ", font=('Helvetica', 17, 'bold'), fg_color="#ADD8E6", text_color='black')
                self.flag = 0
        else:
            self.dialog.configure(text="  Sent-> Not in Connect  ",  font=('Helvetica', 17, 'bold'))

        #self.switchButtonState(self.button_write)
        #self.switchButtonState(self.button_read)
        #self.switchButtonState(self.button_upload)

    def readEntry(self, data):
        # get the values from the entry widgets
        values = []
        for entry_list in data:
            frame_values = []
            for entry in entry_list:
                if isinstance(entry, str):
                    frame_values.append(entry)
                else:
                    frame_values.append(float(entry.get()))
            values.append(frame_values)
            

        # converting the string and float values to hexvalues and sent to bytes
        value = b''
        for i in values:
            value = b''
            bv = b''
            for j in i:
                # print(j)
                if isinstance(j, str):
                    bv = bytes.fromhex(j[2:])
                    value = value + bytes.fromhex(j[2:])
                else:
                    value = value + struct.pack('<f', j)
            # bytes values
            value = value + bv
            
            self.ser.write(bytes(value))
            start_time = time.time()
            res = None
            # sending and checking whether the values has been sent or not if
            # is then sent the next
            while True:
                if self.ser.in_waiting > 0:
                    res = self.ser.read(20)
                    #
                    # if bytes(value) == res:
                    #     print('True')
                    # else:
                    #     print('False')
                    break
                if time.time() - start_time > self.ser.timeout:
                    break
            if res:
                self.open = 1
                # break
            else:
                break
        # entries that has been sent
        if self.open == 1:
            for entry_list in data:
                for entry in entry_list:
                    if isinstance(entry, str):
                        continue
                    else:
                        entry.configure(fg_color="#ADD8E6")
            self.flag = 1
            self.open = 0
        else:
            self.flag = 0
            self.open = 0
    

    def commonRead(self, parent_frame, data, title):
        filename = self.get_json_file_path(data)
        with open(filename, 'r') as f:
            data = json.load(f)
            value = data['instructions']
            
            # Create a collapsible frame for this configuration section
            collapsible_frame = CollapsibleFrame(parent_frame, text=title)
            collapsible_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=1)
            
            entry_lists = []
            
            # Create a grid layout for the fields
            grid_frame = ctk.CTkFrame(collapsible_frame.sub_frame, fg_color="white")
            grid_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=0)
            
            # Configure grid columns (8 columns total)
            for i in range(8):
                grid_frame.grid_columnconfigure(i, weight=1, uniform="col")
            
            row = 0
            for i in range(len(value)):
                val = []
                ind = value[i]['fields']
                val.append(value[i]['id'])
                
                # Create a label for the parameter group
                #group_label = ctk.CTkLabel(grid_frame, text=value[i]['id'], font=('Helvetica', 12, 'bold'), text_color="#2a5885")
                #group_label.grid(row=row, column=0, columnspan=8, sticky='w', pady=(10, 5), padx=5)
                #row += 1
                
                for j in range(len(ind)):
                    # Field name label
                    label = ctk.CTkLabel(grid_frame, text=ind[j]['field1_name'], font=('Helvetica', 13, 'bold'), fg_color='darkslategrey', text_color='white', anchor='w', corner_radius=7)
                    label.grid(row=row, column=j*2, sticky='nsew', padx=(5, 1), pady=2, ipady=4)
                    
                    # Entry field
                    var = tk.StringVar()
                    entry = ctk.CTkEntry(grid_frame, textvariable=var, font=('Helvetica', 13, 'bold'), border_color="black", corner_radius=7)
                    entry.custom_name = ind[j]["field1_name"]
                    entry.grid(row=row, column=j*2+1, sticky='nsew', padx=(1, 5), pady=2, ipady=4)
                    
                    val.append(entry)
                
                entry_lists.append(val)
                row += 1
                
            return entry_lists

    # # Receiving the values
    def received_data(self, vbyte, entry_list):
        le = 0
        flag = 0
        
        for i in vbyte:
            start_time = time.time()
            byte_to_send = bytes.fromhex(i)
            check = True
            
            while check:
                self.ser.write(byte_to_send)
                
                while True:
                    if self.ser.in_waiting:
                        res = self.ser.read(20)
                        print(res)

                        # Check if the expected response length is reached
                        if len(res) >= 20:
                            first_byte = res[0]      # Validate the first and last 
                            last_byte = res[-1]

                            if res.hex().startswith("ff") and len(res) == 20:
                                if first_byte == 0xff and last_byte != 0x00:
                                    self.flag = 1
                                else:
                                    print(f"Received Invalid response: {res()}")
                        
                        val = binascii.hexlify(struct.unpack("<2s", res[0:2])[0]).decode() + "00"
                        val = val.upper()
                        
                        if val == i:
                            check = False
                        break
                    
                    if time.time() - start_time > self.ser.timeout:
                        flag = 1
                        check = False
                        break
                    
            if flag == 1:
                self.dialog.configure(text="Connection Error", font=('Helvetica', 17, 'bold'), fg_color='red', text_color='white')
                self.flag2 = 0
                break
            
            else:
                actual = self.AlgoToRead(res)
                self.AlgoToSetValue(entry_list[le], actual)
                le += 1
                self.open = 0
                self.flag2 = 1

    # packing the values for the floats to hex then bytes
    def AlgoToRead(self, data):
        actual = ["{:.4f}".format(struct.unpack('<f', data[i:i + 4])[0])
                  for i in range(2, len(data) - 4, 4)]
        print(actual)
        return actual

    # setting the value according to the entries that are givens
    def AlgoToSetValue(self, entry_list, data):
        
        for i in range(1, len(entry_list)):
            #print("Entry list name ===> ",entry_list[i].custom_name)
    
            if entry_list[i].custom_name in ["kmph/rpm", "rpm kp", "rpm ki"]:
                formatted_value = "{:.4f}".format(float(data.pop(0)))
            else:
                formatted_value = "{:.2f}".format(float(data.pop(0)))
            entry_list[i].delete(0, tk.END)
            entry_list[i].insert(0, formatted_value)
            entry_list[i].configure(fg_color="#90EE90", text_color="black")


##    # changing the state of buttons
##    def switchButtonState(self, button):
##        if self.button_state == "NORMAL":
##            button['state'] = tk.DISABLED
##        else:
##            button['state'] = tk.NORMAL

    def randomVari(self, rs, entry_list):
        actual = self.AlgoToRead(rs)
        for i in range(1, len(entry_list)):
            entry_list[i].delete(0, tk.END)
            entry_list[i].insert(0, actual.pop(0))
            entry_list[i].configure(fg_color="#ADD8E6", text_color="black")

    def common_write(self, data, list, col, worksheet, entry_list):
        filename = self.get_json_file_path(data)
        with open(filename, 'r') as f:
            data = json.load(f)
            value = data['instructions']
            le = 0
            for i in range(len(value)):
                ind = value[i]['fields']
                row = 0
                # parameters names
                for j in range(len(ind)):
                    worksheet.write(list[row] + str(col), ind[j]['field1_name'])
                    row += 2
                row = 1
                # values
                for j in entry_list[le]:
                    if isinstance(j, str):
                        continue
                    else:
                        worksheet.write(list[row] + str(col), float(j.get()))
                        row += 2
                le += 1
                col += 1
        return col, worksheet

    def write_excelbutton(self):
        now = datetime.now()

        # Default file name with timestamp
        default_filename = 'MCU_Parameters' + now.strftime("D%d_%m_%Y_T%H-%M") + '.xlsx'

        # Open a file save dialog
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             title="Save the Excel file", initialfile=default_filename)

        # If a file path is selected, create and save the Excel file
        if save_path:
            workbook = xlsxwriter.Workbook(save_path)
            worksheet = workbook.add_worksheet()

        # define the list of column letters
        column_letters = ["A", "B", "C", "D", "E", "F", "G", "H"]

        # call the common_write method to populate the worksheet with data
        col, worksheet = self.common_write('full_config.json', column_letters, 1, worksheet, self.entry_lists1)

        # determine the last row index
        last_row_index = col + 1  # +1 to account for 0-based indexing

        # write Date and PicID at the end of the rows
        worksheet.write(f"A{last_row_index + 1}", "Date")
        worksheet.write(f"B{last_row_index + 1}", now.strftime("%d/%m/%Y %H:%M:%S"))
        worksheet.write(f"A{last_row_index + 2}", "PicID")
        worksheet.write(f'B{last_row_index + 2}', self.label2.get())
        worksheet.write(f'C{last_row_index + 2}', self.label3.get())
        worksheet.write(f'A{last_row_index + 3}', "Firmware Version")
        worksheet.write(f'B{last_row_index + 3}', self.firmware2.get())
        worksheet.write(f'C{last_row_index + 3}', self.firmware3.get())

        workbook.close()

        # Update the dialogue box with the message "Values printed"
        self.dialog.configure(text="Values printed", font=('Helvetica', 17, 'bold'), fg_color="#ADD8E6", text_color='black')


    def _on_mousewheel(self, event):
        self.config_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        self.config_canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def update_status(self, message, is_error=False):
        self.status_text.set(message)
        if is_error:
            self.dialog.configure(text_color="red")
        else:
            self.dialog.configure(text_color="black")
        
        # Also update status bar
        self.status_msg.set(message)

class CollapsibleFrame(ctk.CTkFrame):
    def __init__(self, parent, text="", *args, **kwargs):
        ctk.CTkFrame.__init__(self, parent, *args, **kwargs)
        self.configure(fg_color="white")
        
        self.show = tk.BooleanVar(value=True)
        self.title_frame = ctk.CTkFrame(self, fg_color="#f8f9fa", corner_radius=4)
        self.title_frame.pack(fill=tk.BOTH, expand=True)
        
        self.title_label = ctk.CTkLabel(self.title_frame, text=text, font=('Helvetica', 13, 'bold'), text_color="#2a5885")
        self.title_label.pack(side=tk.LEFT, padx=10)
        
        self.toggle_button = ctk.CTkButton(self.title_frame, width=30, height=30,
                                         text="-", font=('Helvetica', 15), fg_color="transparent", text_color="black", hover=False, command=self.toggle)
        self.toggle_button.pack(side=tk.RIGHT)
        
        self.sub_frame = ctk.CTkFrame(self, fg_color="white")
        # Ensure sub_frame is packed on startup if show is True
        if self.show.get():
            self.sub_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        
    def toggle(self):
        if self.show.get():
            self.sub_frame.pack_forget()
            self.toggle_button.configure(text="+")
            self.show.set(False)
        else:
            self.sub_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
            self.toggle_button.configure(text="-")
            self.show.set(True)

# main for running the project
if __name__ == "__main__":
    root = ctk.CTk()
    root.update()
    root.state('zoomed')
    app = MainGui(root)
    app.ser.close()
    root.mainloop()
