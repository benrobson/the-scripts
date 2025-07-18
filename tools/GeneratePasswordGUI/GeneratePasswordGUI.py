# -*- coding: utf-8 -*-
"""
Password Generator GUI

This script generates a random password and displays it in a simple GUI.
The password consists of two random words, a random number, and a random symbol.
The user can copy the password to the clipboard or generate a new one.
"""
import tkinter as tk
from tkinter import messagebox
import random
import string

class PasswordGenerator(tk.Tk):
    """
    Password Generator GUI application.
    """
    def __init__(self):
        """
        Initializes the application.
        """
        super().__init__()
        self.title("Password Generator")
        self.geometry("300x200")
        self.resizable(False, False)

        self.word_list = [
            "computer", "school", "teacher", "student", "pen",
            "pencil", "desk", "chair", "paper", "eraser",
            "ruler", "math", "science", "art", "music",
            "play", "friend", "happy", "sad", "fun",
            "game", "park", "color", "red", "blue",
            "green", "yellow", "purple", "orange", "pink",
            "black", "white", "brown", "gray", "shoes",
            "socks", "shirt", "pants", "hat", "jacket",
            "sweater", "dress", "shorts", "skirt", "glasses",
            "hat", "gloves", "scarf", "boots", "backpack",
            "lunchbox", "bedroom", "kitchen", "bathroom", "livingroom",
            "bed", "table", "chair", "sofa", "TV",
            "computer", "phone", "door", "window", "floor",
            "fruit", "vegetable", "pizza", "cake", "ice cream",
            "candy", "cookie", "sandwich", "juice", "milk",
            "water", "bread", "cheese", "chicken", "pasta",
            "rice", "soup", "salad", "burger", "fries",
            "pizza", "spaghetti", "pancake", "waffle", "grapes",
            "melon", "strawberry", "carrot", "broccoli", "potato",
            "tomato", "onion", "lettuce", "banana", "apple",
            "orange", "pear", "peach", "grapefruit", "lemon",
            "watermelon", "pineapple", "cherry", "blueberry", "raspberry",
            "peas", "corn", "beans", "pumpkin", "cucumber"
        ]
        self.symbols = "!@#$%^&*"

        self.password = self.generate_password()

        self.password_var = tk.StringVar()
        self.password_var.set(self.password)

        self.create_widgets()

    def generate_password(self):
        """
        Generates a new password.
        """
        word1 = random.choice(self.word_list)
        word2 = random.choice(self.word_list)
        number = random.randint(10, 100)
        symbol = random.choice(self.symbols)
        return f"{word1}{number}{symbol}{word2}"

    def create_widgets(self):
        """
        Creates the widgets for the GUI.
        """
        self.password_entry = tk.Entry(self, textvariable=self.password_var, state="readonly", width=40)
        self.password_entry.pack(pady=20)

        self.copy_button = tk.Button(self, text="Copy to Clipboard", command=self.copy_to_clipboard)
        self.copy_button.pack(pady=5)

        self.generate_button = tk.Button(self, text="Generate New Password", command=self.regenerate_password)
        self.generate_button.pack(pady=5)

    def copy_to_clipboard(self):
        """
        Copies the current password to the clipboard.
        """
        self.clipboard_clear()
        self.clipboard_append(self.password)
        messagebox.showinfo("Success", "Password copied to clipboard!")

    def regenerate_password(self):
        """
        Generates a new password and updates the GUI.
        """
        self.password = self.generate_password()
        self.password_var.set(self.password)

if __name__ == "__main__":
    app = PasswordGenerator()
    app.mainloop()
