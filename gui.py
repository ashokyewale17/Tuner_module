import os
import json
import serial
import threading
import time
import tkinter as tk
from PIL import Image
import customtkinter as ctk
from tkinter import messagebox
from utils import get_json_file_path
from data_handler import DataHandler
from file_operations import FileOperations
from credential_dialog import CredentialDialog



class MainGui:
    def __init__(self, root):
        self.root = root
        self.button = tk.Button(root, text="Button")
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        data = 'full_config.json'
        filename = get_json_file_path(data)
        self.data_handler = DataHandler(self, CredentialDialog)
        self.file_operations = FileOperations(self)
        self.setup_gui()

    
    def setup_gui(self):
        self.root.title("Tuner Software V3.9")
        self.root.resizable(True, True)
        self.root.configure(fg_color="white")

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Set appearance mode
        ctk.set_appearance_mode("light")
        self.button_state = "normal"
        
        # Create a frame for the scrollbar and content
        content_frame = ctk.CTkFrame(self.root, fg_color='white')
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Adjust the canvas width to leave space for the vertical scrollbar
        canvas_width = window_width - 20

        # Create a canvas and a scrollbar
        self.canvas = tk.Canvas(content_frame, borderwidth=0)
        self.v_scrollbar = tk.Scrollbar(content_frame, orient="vertical", command=self.canvas.yview)
        self.h_scrollbar = tk.Scrollbar(content_frame, orient="horizontal", command=self.canvas.xview)

        # Create a scrollable frame inside the canvas
        self.scrollable_frame = ctk.CTkFrame(self.canvas, fg_color='white')
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        # Pack the canvas and scrollbar
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        # Add the horizontal scrollbar without overlapping the right-side scrollbar
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")

        # Make the canvas frame expandable
        content_frame.grid_rowconfigure(0, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)

        # Add mouse scrolling for both vertical and horizontal directions
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)  # For Windows and Linux
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)  # For horizontal scrolling

        # Load and display the image
        framePic = ctk.CTkFrame(self.scrollable_frame, fg_color="white")
        img_path = get_json_file_path('APT_LOGO.png')
        image = Image.open(img_path)
        width, height = 500, 45 
        resized_image = image.resize((width, height), Image.LANCZOS)
        img = ctk.CTkImage(dark_image=resized_image, size=(width, height))
        image_label = ctk.CTkLabel(framePic, image=img)
        image_label.configure(text='')
        image_label.grid(row=0, column=0, padx=50, pady=8, columnspan=4)

        # creating the 1 frame
        frameO = ctk.CTkFrame(self.scrollable_frame, width=200, height=50, fg_color="white")
        frame1 = ctk.CTkFrame(frameO, border_width=2, border_color='black', fg_color='white')

        # on click button
        def on_button_click(button):
            if button == self.buttonSS:  
                button.configure(fg_color="blue", text_color="black")
            elif button == self.buttonStop:  
                button.configure(fg_color="red", text_color="black")
        
        # creating the serial port to display in combo-box
        serial_ports = self.detect_serial_ports()
        self.selected_port = tk.StringVar(value="None")
        self.selected_port.trace_add("write", self.on_port_selected)
        
        label = ctk.CTkLabel(frame1, text="COM PORT", font=('GoudyOLStBT', 14, 'bold'))
        label.grid(row=1, column=1, ipadx=3, ipady=3, padx=(4, 0), pady=(2, 12))
        
        serial_box = ctk.CTkComboBox(frame1, variable=self.selected_port, values=serial_ports, border_color='black', width=130)
        serial_box.grid(row=1, column=2, padx=(1, 1), pady=(2, 12), ipadx=10, ipady=4, sticky=tk.N)

        self.buttonSS = ctk.CTkButton(frame1, text="START", font=('GoudyOLStBT', 15, 'bold'), fg_color="grey", state="disabled", 
                                      text_color='white', command=lambda: [on_button_click(self.buttonSS), self.processButtonSS()])
        self.buttonSS.grid(row=7, column=1, columnspan=1, padx=3, pady=(10, 1), ipadx=3, ipady=7)

        self.buttonStop = ctk.CTkButton(frame1, text="STOP", font=('GoudyOLStBT', 15, 'bold'), fg_color="grey", state="disabled",
                                        text_color='white', command=lambda: [on_button_click(self.buttonStop), self.stopButtonSS()])
        self.buttonStop.grid(row=7, column=2, columnspan=1, padx=1, pady=(10, 1), ipadx=3, ipady=7)
        
        frame1.pack(side='left', ipadx=1, ipady=7, padx=15, pady=2)


        # frame 2 dialog and pic ID
        frame2 = ctk.CTkFrame(frameO, border_width=2, border_color='black', fg_color="white")

        labelPicId = ctk.CTkLabel(frame2, text="PIC ID", font=('GoudyOLStBT', 15, 'bold'))
        labelPicId.grid(row=0, column=0, ipadx=90, ipady=5, padx=(14, 0), pady=2)

        labelHigh = ctk.CTkLabel(frame2, text="High", font=('GoudyOLStBT', 14))
        labelHigh.grid(row=1, column=0, ipadx=15, pady=2, padx=(0, 85))

        self.label2 = ctk.CTkEntry(frame2, width=80, border_color='black')
        self.label2.grid(row=2, column=0, ipadx=15, ipady=7, padx=(0, 90), columnspan=1)

        labelLow = ctk.CTkLabel(frame2, text="Low", font=('GoudyOLStBT', 14))
        labelLow.grid(row=1, column=0, ipadx=15, pady=2, padx=(110, 0), columnspan=4)

        self.label3 = ctk.CTkEntry(frame2, width=80, border_color='black')
        self.label3.grid(row=2, column=0, ipadx=15, ipady=7, padx=(114, 0), columnspan=4)
        frame2.pack(side=tk.LEFT, ipadx=4, ipady=3, padx=15)

        # dialog box
        frameF2 = ctk.CTkFrame(frameO,  border_width=2, border_color='black', fg_color='white')
        frameF2.pack(side=tk.LEFT, ipadx=3, ipady=20, padx=5)

        labelDialogbox = ctk.CTkLabel(frameF2, text="STATUS", font=('GoudyOLStBT', 15, 'bold'))
        labelDialogbox.grid(row=0, column=0, ipadx=130, ipady=6, padx=2, pady=4)
        
        self.dialog = ctk.CTkLabel(frameF2, text="",  font=('Helvetica', 14, 'bold'))
        self.dialog.grid(row=1, column=0, ipadx=50, ipady=5, padx=0)

        # creating the firmware frame
        firmware_frame = ctk.CTkFrame(frameO, border_width=2, border_color='black', fg_color='white')
        firmware_frame.pack(side=tk.LEFT, ipadx=2, ipady=1, padx=20)

        # creating the firmware label
        labelFirmware = ctk.CTkLabel(firmware_frame, text="FIRMWARE V", font=('GoudyOLStBT', 15, 'bold'))
        labelFirmware.grid(row=0, column=0, padx=6, pady=2, ipadx=70, ipady=6)

        # creating the "High" label for firmware version entry
        labelHighFirmware = ctk.CTkLabel(firmware_frame, text="High", font=('GoudyOLStBT', 14))
        labelHighFirmware.grid(row=1, column=0, padx=(0, 105))

        # creating the firmware version entry
        self.firmware2 = ctk.CTkEntry(firmware_frame, width=80, border_color='black')
        self.firmware2.grid(row=2, column=0, columnspan=1, ipadx=15, ipady=6, padx=(0, 108), pady=3)

        # creating the "Low" label for firmware version entry
        labelLowFirmware = ctk.CTkLabel(firmware_frame, text="Low", font=('GoudyOLStBT', 14))
        labelLowFirmware.grid(row=1, column=0, padx=(115, 0), columnspan=3)

        # creating the firmware version entry
        self.firmware3 = ctk.CTkEntry(firmware_frame, width=80, border_color='black')
        self.firmware3.grid(row=2, column=0, columnspan=3, ipadx=15, ipady=6, padx=(115, 0), pady=3)

        # Color code
        colLabel = ctk.CTkFrame(frameO, border_width=2, border_color='black', fg_color='white')
        colLabel.pack(side=tk.RIGHT, padx=(100, 100), pady=(30, 10), anchor='e')
        labelColorID = ctk.CTkLabel(colLabel, text="Color Code", font=('GoudyOLStBT', 14, 'bold'))
        labelColorID.grid(row=3, column=0, columnspan=7, ipadx=25, ipady=2, pady=5, padx=0)
        labS = ctk.CTkLabel(colLabel, text="Read", font=('GoudyOLStBT', 13, 'bold'), fg_color="#90EE90", corner_radius=7)
        labR = ctk.CTkLabel(colLabel, text="Write", font=('GoudyOLStBT', 13, 'bold'), fg_color="#ADD8E6", corner_radius=7)
        labS.grid(row=4, column=0, ipadx=15, ipady=2, padx=3, pady=3)
        labR.grid(row=4, column=1, ipadx=15, ipady=2, padx=3, pady=3)
        framePic.pack(anchor='nw', ipadx=10, ipady=5, padx=10)
        frameO.pack(anchor='nw', ipadx=20, ipady=2, pady=1, padx=10)

        # frame 3 for motor config, vehcile and powertrain
        box1 = ctk.CTkFrame(self.scrollable_frame, height=50, border_width=1, fg_color='white')

        self.entry_lists1 = self.commonRead(
            box1, 'full_config.json', "Full Configuration")
        box1.pack(fill=tk.BOTH, expand=True, pady=4)

        frameB = ctk.CTkFrame(self.scrollable_frame, fg_color="white", width=1100, height=50)
        frameB.pack_propagate(False)

        # Function to change button color on hover
        def on_enter(button, color):
            if button.cget('state') == 'normal':
                button.configure(fg_color=color)

        # Function to change button color back to default
        def on_leave(button, color):
            if button.cget('state') == 'normal':
                button.configure(fg_color=color)

        # Function to change button color on click with effect
        def on_button_click(button):
            if button.cget('state') == 'normal':
                button.configure(fg_color='lightsteelblue', text_color='black')
                button.after(100, lambda: button.configure(text_color='white'))

        # Function to change button color back after release
        def on_release(button):
            if button.cget('state') == 'normal':
                button.configure(fg_color='royalblue', text_color='white')

        # Function to handle the click effect specifically for the Print button
        def on_click_print(button):
            if button.cget('state') == 'normal':
                button.configure(fg_color='blue', text_color='black')  

        # Function to handle the button release specifically for the Print button
        def on_release_print(button):
            if button.cget('state') == 'normal':
                button.configure(fg_color='royalblue', text_color='white')

        def configure_button(button):
            button.configure(width=160, height=50)  # Set fixed button size
            button.pack(side=tk.LEFT, padx=20, pady=10)  # Pack the button with padding
            button.pack_propagate(False)

        def disable_all_buttons():
            # Disable all buttons
            self.button.configure(state='disabled', fg_color='grey', text_color='darkgrey')
            self.button2.configure(state='disabled', fg_color='grey', text_color='darkgrey')
            self.button3.configure(state='disabled', fg_color='grey', text_color='darkgrey')
            button4.configure(state='disabled', fg_color='grey', text_color='darkgrey')
            openB.configure(state='disabled', fg_color='grey', text_color='darkgrey')

        def enable_all_buttons():
            # Enable all buttons
            self.button.configure(state='normal', fg_color='royalblue', text_color='white', hover=True)
            self.button2.configure(state='normal', fg_color='royalblue', text_color='white', hover=True)
            self.button3.configure(state='normal', fg_color='royalblue', text_color='white', hover=True)
            button4.configure(state='normal', fg_color='royalblue', text_color='white', hover=True)
            openB.configure(state='normal', fg_color='royalblue', text_color='white', hover=True)

        def process_read_button():
            disable_all_buttons()
            try:
                time.sleep(0.1)
                self.data_handler.processButtonReceive()
            finally:
                enable_all_buttons()

        # Function to start the reading process in a new thread
        def start_read_process():
            threading.Thread(target=process_read_button, daemon=True).start()

        # Ensure fixed padding and size for buttons to prevent shrinking
        def configure_button(button):
            button.pack(side=tk.LEFT, ipadx=26, padx=20, ipady=10)

        self.button = ctk.CTkButton(frameB, text='Write', font=('Helvetica', 15, 'bold'), fg_color='royalblue',
                                text_color='white', command=lambda: [on_button_click(self.button), self.data_handler.processButtonSend()])
        self.button.bind("<Enter>", lambda event, button=self.button: on_enter(button, 'indigo'))
        self.button.bind("<Leave>", lambda event, button=self.button: on_leave(button, 'royalblue'))
        self.button.bind("<ButtonPress>", lambda event, button=self.button: on_click(button))
        self.button.bind("<ButtonRelease>", lambda event, button=self.button: on_release(button))
        configure_button(self.button)

        self.button2 = ctk.CTkButton(frameB, text='Read', font=('Helvetica', 15, 'bold'), fg_color='royalblue',
                                 text_color='white', command=lambda: [on_button_click(self.button2), start_read_process()])
        self.button2.bind("<Enter>", lambda event, button=self.button2: on_enter(button, 'indigo'))
        self.button2.bind("<Leave>", lambda event, button=self.button2: on_leave(button, 'royalblue'))
        self.button2.bind("<ButtonPress>", lambda event, button=self.button2: on_click(button))
        configure_button(self.button2)

        openB = ctk.CTkButton(frameB, text="Upload", font=('Helvetica', 15, 'bold'), fg_color='royalblue',
                         text_color='white', command=lambda: [on_button_click(openB), self.file_operations.open_file()])
        openB.bind("<Enter>", lambda event, button=openB: on_enter(button, 'indigo'))
        openB.bind("<Leave>", lambda event, button=openB: on_leave(button, 'royalblue'))
        openB.bind("<ButtonPress>", lambda event, button=openB: on_click(button))
        configure_button(openB)

        button4 = ctk.CTkButton(frameB, text="Print", font=('Helvetica', 15, 'bold'), fg_color='royalblue', 
                             text_color='white', command=lambda: [on_button_click(button4), self.file_operations.write_excelbutton()])
        button4.bind("<Enter>", lambda event, button=button4: on_enter(button, 'indigo'))
        button4.bind("<Leave>", lambda event, button=button4: on_leave(button, 'royalblue'))
        button4.bind("<ButtonPress>", lambda event, button=button4: on_click(button))
        button4.bind("<ButtonRelease>", lambda event, button=button4: on_release(button))
        configure_button(button4)

        self.button3 = ctk.CTkButton(frameB, text='Zero Angle Estimate', font=('Helvetica', 14, 'bold'), fg_color='royalblue',
                                 text_color='white', command=lambda: [on_button_click(self.button3), self.start_continuous()])
        self.button3.bind("<Enter>", lambda event, button=self.button3: on_enter(button, 'indigo'))
        self.button3.bind("<Leave>", lambda event, button=self.button3: on_leave(button, 'royalblue'))
        self.button3.bind("<ButtonPress>", lambda event, button=self.button3: on_click(button))
        self.button3.bind("<ButtonRelease>", lambda event, button=self.button3: on_release(button))
        configure_button(self.button3)

        frameB.pack(pady=7)
        frameB.pack_propagate(False) 

        # serial object
        self.ser = serial.Serial()
        self.root.mainloop()

    def start_continuous(self):
        thread = threading.Thread(target=self.data_handler.receiveRandom())
        thread.start()

    def detect_serial_ports(self):
        self.ports = serial.tools.list_ports.comports()
        if self.ports:
            return [port.device for port in serial.tools.list_ports.comports()]
        else:
            return ('No ports available',)

    def on_enter(self, button, color):
        button.configure(fg_color=color)

    def on_leave(self, button, color):
        button.configure(fg_color=color)

    def disable_button(self, button):
        button.configure(state="disabled", fg_color="grey", text_color="darkgrey")
        button.unbind("<Enter>")
        button.unbind("<Leave>")

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
            self.buttonSS.configure(fg_color="royalblue", state="normal", text_color="white")
        else:  
            self.buttonSS.configure(fg_color="grey", text_color="white")

    # common for all three entries for json file
    def commonRead(self, box, data, Text):
        filename = get_json_file_path(data)
        with open(filename, 'r') as f:
            data = json.load(f)
            value = data['instructions']
            frame1 = tk.LabelFrame(box, bg='white')
            frame1_label = tk.Label(frame1, text=Text, font=('GoudyOLStBT', 15, 'bold'), fg='white', bg='teal')
            frame1_label.pack(fill='x', pady=5, ipady=4)
            frame1.pack(fill=tk.BOTH, expand=True, padx=15, pady=3, ipadx=5, ipady=1)
            row = 1
            col = 1
            entry_lists = []
            frame = tk.Frame(frame1, bg='white')
            frame.pack(fill=tk.BOTH, expand=True)
            for col in range(8):
                frame.columnconfigure(col, weight=1)
            row = 0
            for i in range(len(value)):
                col = 1
                val = []
                ind = value[i]['fields']
                val.append(value[i]['id'])
                for j in range(len(ind)):
                    label = ctk.CTkLabel(frame, text=ind[j]['field1_name'], font=('GoudyOLStBT', 13, 'bold'),
                                     fg_color='darkslategrey', text_color='white', width=180, corner_radius=7, anchor="w")
                    var = tk.StringVar()
                    entry = ctk.CTkEntry(frame, textvariable=var, font=('GoudyOLStBT', 13, 'bold'), border_color='black',
                                     width=24)
                    entry.custom_name=ind[j]["field1_name"]
                    label.grid(row=row, column=2 * j,
                                ipadx=2, ipady=6, padx=1, pady=2)

                    entry.grid(row=row, column=2 * j + 1, ipadx=70,
                               ipady=6, padx=12, pady=2)
                    val.append(entry)
                    col += 1
                entry_lists.append(val)
                row += 1
            return entry_lists

    # changing the state of buttons
    def switchButtonState(self, button):
        if self.button_state == "NORMAL":
            button['state'] = tk.DISABLED
        else:
            button['state'] = tk.NORMAL

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
