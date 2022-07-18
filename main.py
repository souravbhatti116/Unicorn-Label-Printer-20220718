import ctypes
import sys
import tkinter as tk
import inventory
import printLabel
from PIL import *
from PIL import Image, ImageTk

def main():
    root = tk.Tk()
    root.title("Welcome to Shipping/Inventory App")
    root.geometry("450x400+450+200")
    root.configure(background="honeydew3")
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    image = Image.open("TRACKONOMY-Logo 200 x 50.png")
    photo = ImageTk.PhotoImage(image)
    #tk.Label(root, image=photo).pack(anchor='n', ipadx=50, ipady=15)
    tk.Label(root, image=photo, width=250).place(x=100, y=50)
    # Lbl1=tk.Label(root, text="Please select the application.", font=("Times new roman", 20, 'bold'),
    #           bg='honeydew3').place(x=65, y=100)
    root.resizable(False, False)

    def print_label():
        root.destroy()
        printLabel.printing()

    def inventoryOut():
        root.destroy()
        inventory.inventoryCheckOut()

    def inventoryIn():
        root.destroy()
        inventory.inventoryCheckIN()

    printbtn = tk.Button(root, text="Print Labels", command=lambda: print_label(), width=20, bg="aquamarine",
                        font=("Times new roman", 15), relief="ridge")
    printbtn.place(x=110, y=150)

    shipbtn = tk.Button(root, text="Inventory Check-out", command=lambda: inventoryOut(), width=20, bg="Light blue",
                        font=("Times new roman", 15), relief="ridge")
    shipbtn.place(x=110, y=250)

    invbtn = tk.Button(root, text="Inventory Check-in", command=lambda: inventoryIn(), width=20, bg="Orange",
                       font=("Times new roman", 15), relief='ridge')
    invbtn.place(x=110, y=200)

    exitbtn = tk.Button(root, text="Exit", command=lambda: sys.exit(), width=20, bg="orange red",
                        font=("Times new roman", 15), relief="ridge")
    exitbtn.place(x=110, y=300)

    root.bind('<Escape>', lambda event: sys.exit())

    root.mainloop()


main()
