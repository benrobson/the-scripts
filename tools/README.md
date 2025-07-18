# Password Generator GUI

This script generates a random password and displays it in a simple GUI.
The password consists of two random words, a random number, and a random symbol.
The user can copy the password to the clipboard or generate a new one.

## How to Compile

1.  **Install Python:** If you don't have Python installed, download and install it from [python.org](https://python.org).
2.  **Install Dependencies:** Open a terminal or command prompt and run the following command to install the required dependencies:
    ```bash
    pip install pyinstaller
    ```
3.  **Run the Build Script:** Navigate to the `tools` directory in your terminal or command prompt and run the following command:
    ```bash
    python build.py
    ```
    This will create a `dist` folder in the `tools` directory containing the `GeneratePasswordGUI.exe` executable.

    To create the executable and add it to the startup folder (Windows only), run the following command:
    ```bash
    python build.py --startup
    ```
