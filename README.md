# SAP GUI Scripting Library

## Table of contents

* [ToC](#table-of-contents)
* [Python installation](#python-installation)
* [Install](#install)

## Python installation
1. Download [last version of Python 3.x installer](https://www.python.org/downloads/)
2. Run the installer
3. While installation choose folowing option:
    - Add python 3.x to PATH

## Install

### Pip installation (recomended)
Installation is easy. Run in windows console (command line interpreter - cmd):
```sh
pip install PySapGUI
```
If your computer is behind a proxy set additional option --proxy in following format:
```sh
pip install sapsec --proxy http://user:password@proxyserver:port
```

### Installation from github
If for some reason the installation was not successful (with pip) there is an opportunity to install sapsec from github source files.
1. Download [zip archive](https://github.com/gutskodv/PySapGUI/archive/master.zip) with project source codes. Or use git clone:
```sh
git clone https://github.com/gutskodv/PySapGUI.git
```
2. Unpack files from dowloaded zip archive. And go to project directory with setup.py file.
3. Ugrade pip, Install Wheel package, Collect sapsec package:
```sh
python -m pip install --upgrade pip
pip install wheel
python setup.py bdist_wheel
```
4. Install sapsec package from generaed python wheel in dist subdirectory:
```sh
python setup.py dist\PySapGUI*.whl
```

### Requirements
You can manually intall requirements if they were not installed in automatic mode.
1. PyWin32 (Python extensions for Microsoft Windows Provides access to much of the Win32 API, the ability to create and use COM objects, and the Pythonwin environment).
```sh
pip install pywin32
```
