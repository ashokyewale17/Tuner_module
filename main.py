import customtkinter as ctk
from gui import MainGui

if __name__ == "__main__":
    root = ctk.CTk()
    root.update()
    root.state('zoomed')
    app = MainGui(root)
    app.ser.close()
    root.mainloop()