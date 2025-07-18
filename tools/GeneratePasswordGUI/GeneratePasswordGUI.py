# -*- coding: utf-8 -*-
"""
Password Generator GUI

This script generates a random password and displays it in a simple GUI.
The password consists of two random words, a random number, and a random symbol.
The user can copy the password to the clipboard or generate a new one.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import random

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
        self.geometry("400x150")
        self.resizable(False, False)

        # Style configuration
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("TButton", padding=6, relief="flat", background="#cccccc")
        style.configure("TEntry", padding=6, relief="flat")

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

        self.password_var = tk.StringVar()
        self.create_widgets()
        self.regenerate_password()

    def get_password_strength(self, password):
        """
        Determines the strength of the password.
        """
        length = len(password)
        if length < 8:
            return "Weak"
        elif length < 12:
            return "Medium"
        else:
            return "Strong"

    def generate_password(self):
        """
        Generates a new password.
        """
        word1 = random.choice(self.word_list)
        word2 = random.choice(self.word_list)
        number = random.randint(10, 100)
        symbol = random.choice(self.symbols)
        password = f"{word1}{number}{symbol}{word2}"
        return password.capitalize()

    def create_widgets(self):
        """
        Creates the widgets for the GUI.
        """
        self.configure(bg="#2d2d2d")  # Dark background color

        main_frame = ttk.Frame(self, padding="20", style="App.TFrame")
        main_frame.pack(expand=True, fill="both")

        style = ttk.Style(self)
        style.configure("App.TFrame", background="#2d2d2d")
        style.configure("TButton", padding=6, relief="flat",
                        background="#4a4a4a", foreground="white",
                        font=("Arial", 10, "bold"))
        style.map("TButton",
                  background=[("active", "#6a6a6a")],
                  foreground=[("active", "white")])
        style.configure("TEntry", padding=10, relief="flat",
                        background="#4a4a4a", foreground="white",
                        font=("Arial", 14))
        style.configure("TLabel", padding=10,
                        background="#2d2d2d", foreground="white",
                        font=("Arial", 10))

        title_label = ttk.Label(main_frame, text="Password Generator",
                                font=("Arial", 16, "bold"),
                                foreground="#00b0ff") # Light blue color for title
        title_label.pack(pady=(0, 10))

        self.password_entry = ttk.Entry(main_frame, textvariable=self.password_var, state="readonly",
                                        justify="center")
        self.password_entry.pack(fill="x", expand=True, ipady=5)

        self.strength_label = ttk.Label(main_frame, text="",
                                        font=("Arial", 10, "italic"))
        self.strength_label.pack(pady=(5, 0))

        button_frame = ttk.Frame(main_frame, style="App.TFrame")
        button_frame.pack(fill="x", expand=True, pady=10)

        self.copy_button = ttk.Button(button_frame, text="Copy to Clipboard",
                                      command=self.copy_to_clipboard)
        self.copy_button.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.generate_button = ttk.Button(button_frame, text="Generate New Password",
                                          command=self.regenerate_password)
        self.generate_button.pack(side="right", expand=True, fill="x", padx=(5, 0))

        self.update_strength_indicator()

    def update_strength_indicator(self):
        """
        Updates the password strength indicator.
        """
        strength = self.get_password_strength(self.password)
        self.strength_label.config(text=f"Password Strength: {strength}",
                                   foreground=self.get_strength_color(strength))

    def get_strength_color(self, strength):
        """
        Returns the color for the password strength.
        """
        if strength == "Weak":
            return "red"
        elif strength == "Medium":
            return "orange"
        else:
            return "green"

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
        self.update_strength_indicator()

if __name__ == "__main__":
    app = PasswordGenerator()
    app.mainloop()
