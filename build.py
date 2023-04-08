import sys
import subprocess

if __name__ == '__main__':
    pyinstaller_command = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--uac-admin',
        '--noconsole',
        '--add-data', 'context_menu.py;.',
        'app.py'
    ]

    subprocess.run(pyinstaller_command, check=True)