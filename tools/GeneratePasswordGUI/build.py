import PyInstaller.__main__
import os
import sys
import shutil

def get_startup_folder():
    if sys.platform == 'win32':
        return os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    else:
        return None

def build(add_to_startup=False):
    """
    Builds the executable.
    """
    # Get the absolute path to the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script = os.path.join(script_dir, 'GeneratePasswordGUI.py')
    name = 'GeneratePasswordGUI'
    icon = os.path.join(script_dir, 'favicon.ico')

    pyinstaller_args = [
        '--name=%s' % name,
        '--onefile',
        '--windowed',
        '--icon=%s' % icon,
        '--distpath=%s' % os.path.join(script_dir, 'dist'),
        '--workpath=%s' % os.path.join(script_dir, 'build'),
        script,
    ]

    PyInstaller.__main__.run(pyinstaller_args)

    if add_to_startup:
        startup_folder = get_startup_folder()
        if startup_folder:
            try:
                shutil.copy(os.path.join(script_dir, 'dist', f'{name}.exe'), startup_folder)
                print(f"Added {name}.exe to startup folder.")
            except PermissionError:
                print("Permission denied: Could not copy to startup folder.")
                print("Please run this script as an administrator to add the application to startup.")

if __name__ == '__main__':
    add_to_startup = '--startup' in sys.argv
    build(add_to_startup)
