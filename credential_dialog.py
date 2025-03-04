import tkinter as tk
import customtkinter as ctk
from tkinter import simpledialog

class CredentialDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Authentication")
        self.geometry("300x250")

        self.transient(parent)
        self.grab_set()

        self.username_label = ctk.CTkLabel(self, text="Username:", font=('GoudyOLStBT', 13, 'bold'))
        self.username_label.pack(pady=(10, 0))
        self.username_entry = ctk.CTkEntry(self, font=('GoudyOLStBT', 13, 'bold'))
        self.username_entry.pack(pady=5, ipadx=4, ipady=5)

        self.password_label = ctk.CTkLabel(self, text="Password:", font=('GoudyOLStBT', 13, 'bold'))
        self.password_label.pack(pady=(10, 0))
        self.password_entry = ctk.CTkEntry(self, font=('GoudyOLStBT', 14, 'bold'), show='*')
        self.password_entry.pack(pady=5, ipadx=4, ipady=5)

        self.submit_button = ctk.CTkButton(self, text="Submit", font=('GoudyOLStBT', 13, 'bold'), command=self.on_submit, text_color='white')
        self.submit_button.pack(pady=(10, 0), ipadx=4, ipady=5)

        self.username = None
        self.password = None

        self.bind("<Return>", lambda event: self.on_submit())

    def on_submit(self):
        self.username = self.username_entry.get()
        self.password = self.password_entry.get()
        self.destroy()  

    def get_credentials(self):
        self.wait_window()  
        return self.username, self.password
