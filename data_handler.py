import os
import sys
import time
import json
import struct
import tkinter as tk
import serial.tools.list_ports
import serial
import threading
from tkinter import messagebox
from tkinter import simpledialog
from credential_dialog import CredentialDialog

class DataHandler:
    def __init__(self, gui, CredentialDialog):
        self.gui = gui
        self.CredentialDialog = CredentialDialog
        self.uartState = False
        self.is_running = False
        self.is_action_in_progress = False
        self.start_time = time.time()
        self.data = None
        self.entry_lists1 = []
        self.flag = 0
        self.flag2 = 0
        self.open = 0
        self.ser = serial.Serial()
        self.selected_port = tk.StringVar()
        self.credentials_provided = False

    # showing the random data that has to be received from the code
    def receiveRandom(self):
        rs = None
        self.is_running = True
        start_time = time.time()
        self.gui.switchButtonState(self.gui.button)
        self.gui.switchButtonState(self.gui.button2)
        self.gui.switchButtonState(self.gui.button3)
        if self.uartState:
            while self.is_running:
                rs = self.ser.read(20)
                if time.time() - start_time > 12:
                    break
                if rs:
                    # print('received->',rs)
                    self.randomVari(rs, self.entry_lists1[2])
                    break

            if self.is_running:
                self.is_running = False
                self.button3['text'] = "Zero Angle Estimate"
            else:
                self.button3['text'] = "Zero Angle Estimate"

        self.gui.switchButtonState(self.gui.button)
        self.gui.switchButtonState(self.gui.button2)
        self.gui.switchButtonState(self.gui.button3)

    # reading the picId
    def readFrame(self):
        if self.uartState:
            self.ser.write(bytes.fromhex('20012F'))
            bval = self.ser.read(10)
            if len(bval) < 10:
                self.gui.dialog.configure(text="Incomplete PIC ID response", font=('Helvetica', 14, 'bold'))
                return

            unpacked_data = bval[1:-1]
            h1 = (unpacked_data[:4]).hex()
            h1 = h1[::-1]
            st1 = h1[1] + h1[0] + h1[3] + h1[2] + h1[5] + h1[4] + h1[7] + h1[6]
            h2 = (unpacked_data[4:]).hex()
            h2 = h2[::-1]
            st2 = h2[1] + h2[0] + h2[3] + h2[2] + h2[5] + h2[4] + h2[7] + h2[6]

            # update (self.label2) with h1 value
            self.gui.label2.configure(state='normal')
            self.gui.label2.delete(0, tk.END)
            self.gui.label2.insert(0, st1)
            self.gui.label2.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')

            # update (self.label3) with h2 value
            self.gui.label3.configure(state='normal')
            self.gui.label3.delete(0, tk.END)
            self.gui.label3.insert(0, st2)
            self.gui.label3.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')

            # Small delay before sending the next command
            time.sleep(0.1)

            # Read firmware version
            self.ser.timeout = 1
            self.ser.write(bytes.fromhex('30013F'))
            bval_fw = self.ser.read(10)
            if len(bval_fw) < 10:
                self.gui.dialog.configure(text="Please Wait", font=('Helvetica', 14, 'bold'))
                return

            unpacked_data_fw = bval_fw[1:-1]
            fw1 = (unpacked_data_fw[:4]).hex()
            fw1 = fw1[::-1]
            st_fw1 = fw1[1] + fw1[0] + fw1[3] + fw1[2] + fw1[5] + fw1[4] + fw1[7] + fw1[6]
            fw2 = (unpacked_data_fw[4:]).hex()
            fw2 = fw2[::-1]
            st_fw2 = fw2[1] + fw2[0] + fw2[3] + fw2[2] + fw2[5] + fw2[4] + fw2[7] + fw2[6]

            # update firmware labels
            self.gui.firmware2.configure(state='normal')
            self.gui.firmware2.delete(0, tk.END)
            self.gui.firmware2.insert(0, st_fw1)
            self.gui.firmware2.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')

            self.gui.firmware3.configure(state='normal')
            self.gui.firmware3.delete(0, tk.END)
            self.gui.firmware3.insert(0, st_fw2)
            self.gui.firmware3.configure(font=('GoudyOLStBT', 14, 'bold'), justify='center', state='readonly')                
        else:
            self.gui.dialog.configure(text="Not in connect", font=('Helvetica', 14, 'bold'), text_color='red')

    # for connecting the process button with port
    def processButtonSS(self):
        if not self.uartState:
            self.ser.port = self.selected_port.get()
            self.ser.baudrate = 9600
            self.ser.timeout = 1
            try:
                self.ser.open()
                print(f"Serial port {self.ser.port} opened")
            except:
                self.gui.dialog.configure(text="Can't open the Port", font=('Helvetica', 17, 'bold'), text_color="red")
                print(f"Failed to open serial port:")   
        
            if self.ser.is_open:
                self.gui.buttonSS.configure(fg_color="lightsteelblue")
                self.gui.buttonSS.update_idletasks()
                self.dump()
                print("Dump command sent")
            
                if self.open == 1:
                    self.gui.buttonSS["text"] = "START"  
                    self.uartState = True

                    if self.flag2 == 0:
                        self.gui.dialog['text'] = "Busy"
                        self.gui.dialog.configure()
                
                    self.serial_operations()

                    self.gui.buttonSS.update_idletasks()
                    self.gui.disable_button(self.gui.buttonSS)
                    self.gui.enable_button(self.gui.buttonStop)

    def stopButtonSS(self):
        if self.uartState:
            self.ser.close()
            self.gui.buttonSS["text"] = "START"
            self.uartState = False
            print("Serial port closed")
            self.gui.dialog.configure(text="Port Closed", font=('Helvetica', 17, 'bold'))

            self.gui.disable_button(self.gui.buttonStop)
            self.gui.enable_button(self.gui.buttonSS)

    def serial_operations(self):
        self.processReceived()
        print("Open-----> Port")
        self.readFrame()
        print('Started Reading------>')
    
        self.flag2 = 1
        if self.flag2 == 1:
            self.gui.dialog.configure(text="Read Successful", font=('Helvetica', 17, 'bold'), text_color='black')
            self.flag2 = 0
            self.open = 0

    #send dump command
    def dump(self):
        start_time = time.time()
        self.ser.write(bytes.fromhex('10011f'))
        response = None
        while True:
            if self.ser.in_waiting:
                response = self.ser.read(3)
                print(f"Received Dump response: {response.hex()}")
                break
            if time.time() - start_time > self.ser.timeout:
                break

        if response:
            self.open = 1
        else:
            self.gui.dialog.configure(text="Connection Error", font=('Helvetica', 17, 'bold'), fg_color='red', text_color='white')

    # stopping the function
    def stop_receiveRandom(self):
        self.is_running = False

    # for receiving the values from the uart
    def processButtonReceive(self):
        self.gui.switchButtonState(self.gui.button)
        self.gui.switchButtonState(self.gui.button2)
        self.gui.switchButtonState(self.gui.button3)
        
        if self.uartState:
            if self.flag2 == 0:
                self.gui.dialog.configure(text="Busy", font=('Helvetica', 17, 'bold'), fg_color='orange')

            self.is_running = False
            self.readFrame()
            self.gui.dialog.configure(text="Reading...", font=('Helvetica', 17, 'bold'), text_color='white', fg_color='indigo')

            value = []
            for data in self.entry_lists1:
                for entry in data:
                    if isinstance(entry, str):
                        value.append(entry[2:]+"00")

            self.received_data(value, self.entry_lists1)

            if self.flag2 == 1:
                self.gui.dialog.configure(text="Read Successful", font=('Helvetica', 17, 'bold'), fg_color="#90EE90", text_color='black')
                self.flag2 = 0
            else:
                self.gui.dialog.configure(text="Connection Error", font=('Helvetica', 17, 'bold'), fg_color='red', text_color='white')

        else:
            self.gui.dialog.configure(text="Read-> Not in Connect", font=('Helvetica', 17, 'bold'))

    # receving the data if the button is pressed
    def processReceived(self):
        self.gui.switchButtonState(self.gui.button)
        self.gui.switchButtonState(self.gui.button2)
        self.gui.switchButtonState(self.gui.button3)
        #
        if (self.uartState):
            self.gui.dialog.configure(text="Busy", font=('Helvetica', 17, 'bold'))
            value = []
            for data in self.entry_lists1:
                for entry in data:
                    if isinstance(entry, str):
                        value.append(entry[2:]+"00")

            self.received_data(value, self.entry_lists1)

            if(self.flag2 == 1):
                self.gui.dialog.configure(text="Read Successful", font=('Helvetica', 17, 'bold'), fg_color="#90EE90", text_color='black')
                self.flag2 = 0
        else:
            self.gui.dialog.configure(text="Received-> Not In Connect")

    #define promt for credential
    def prompt_for_credentials(self):
            dialog = CredentialDialog(self.gui.root)
            return dialog.username, dialog.password

    # sending the data to the uart
    def processButtonSend(self):
##            if not self.credentials_provided:
##                username, password = self.prompt_for_credentials()
##                if username is None or password is None:
##                    return
##                
##                if username == 'admin' and password == 'apt123':
##                    self.credentials_provided = True
##                else:
##                    messagebox.showerror("Authentication Failed", "Invalid username or password!")
##                    return
                
            self.gui.switchButtonState(self.gui.button)
            self.gui.switchButtonState(self.gui.button2)
            self.gui.switchButtonState(self.gui.button3)
            
            if (self.uartState):
                self.gui.dialog.configure(text="Busy", font=('Helvetica', 17, 'bold'))
                self.is_running = False
                self.readEntry(self.entry_lists1)

                if self.flag == 0:
                    self.gui.dialog.configure(text="Connection Error",  font=('Helvetica', 17, 'bold'), fg_color="red", text_color='white')
                else:
                    self.gui.dialog.configure(text="Write Successful", font=('Helvetica', 17, 'bold'), fg_color="#ADD8E6", text_color='black')
                    self.flag = 0
            else:
                self.gui.dialog.configure(text="Sent-> Not in Connect",  font=('Helvetica', 17, 'bold'))

            self.gui.switchButtonState(self.gui.button)
            self.gui.switchButtonState(self.gui.button2)
            self.gui.switchButtonState(self.gui.button3)

    def readEntry(self, data):
        values = []
        for entry_list in data:
            frame_values = []
            for entry in entry_list:
                if isinstance(entry, str):
                    frame_values.append(entry)
                else:
                    frame_values.append(float(entry.get()))
            values.append(frame_values)
            

        value = b''
        for i in values:
            value = b''
            bv = b''
            for j in i:
                if isinstance(j, str):
                    bv = bytes.fromhex(j[2:])
                    value = value + bytes.fromhex(j[2:])
                else:
                    value = value + struct.pack('<f', j)
            value = value + bv
            self.ser.write(bytes(value))
            start_time = time.time()
            res = None
            while True:
                if self.ser.in_waiting > 0:
                    res = self.ser.read(20)
                    break
                if time.time() - start_time > self.ser.timeout:
                    break
            if res:
                self.open = 1
            else:
                break
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

    # Receiving the values
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
                self.gui.dialog.configure(text="Connection Error", font=('Helvetica', 17, 'bold'), fg_color='red', text_color='white')
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
        actual = ["{:.4f}".format(struct.unpack('<f', data[i:i + 4])[0]) for i in range(2, len(data) - 4, 4)]
        print(actual)
        return actual

     # setting the value according to the entries that are givens
    def AlgoToSetValue(self, entry_list, data):
        for i in range(1, len(entry_list)):
            print("Entry list name ===> ", entry_list[i].custom_name)
            if entry_list[i].custom_name in ["kmph/rpm", "rpm kp", "rpm ki"]:
                formatted_value = "{:.4f}".format(float(data.pop(0)))
            else:
                formatted_value = "{:.2f}".format(float(data.pop(0)))
            entry_list[i].delete(0, tk.END)
            entry_list[i].insert(0, formatted_value)
            entry_list[i].configure(fg_color="#90EE90", text_color="black")


    def randomVari(self, rs, entry_list):
        actual = self.AlgoToRead(rs)
        for i in range(1, len(entry_list)):
            entry_list[i].delete(0, tk.END)
            entry_list[i].insert(0, actual.pop(0))
            entry_list[i].configure(fg_color="#ADD8E6", text_color="black")
