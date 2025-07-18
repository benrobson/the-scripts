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
    script = 'GeneratePasswordGUI.py'
    name = 'GeneratePasswordGUI'
    # Get the absolute path to the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon = os.path.join(script_dir, '..', '..', 'scripts', 'assets', 'favicon.ico')

    pyinstaller_args = [
        '--name=%s' % name,
        '--onefile',
        '--windowed',
        '--icon=%s' % icon,
        script,
    ]

    PyInstaller.__main__.run(pyinstaller_args)

    if add_to_startup:
        startup_folder = get_startup_folder()
        if startup_folder:
            shutil.copy(os.path.join('dist', f'{name}.exe'), startup_folder)
            print(f"Added {name}.exe to startup folder.")

if __name__ == '__main__':
    add_to_startup = '--startup' in sys.argv
    build(add_to_startup)
