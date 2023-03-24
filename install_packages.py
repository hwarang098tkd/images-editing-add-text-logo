import subprocess

# List of packages to be installed
required_packages = ['pandas', 'Pillow', 'openpyxl', 'random']


def install_packages():
    # Check if each package is installed and install it if not
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"{package} is not installed, installing...")
            subprocess.check_call(['pip', 'install', package])
