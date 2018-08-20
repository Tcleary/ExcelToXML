# ExcelToXML
Excel to Xml for batch uploads for Windows PC
use run.py for Linux/Mac after installing requirments in "requirments.txt"

### Procedure

* Open your command prompt and install virtualenv (a Python virtual enviroment package) using Python's pip tool:
```
pip install virtualenv
```
* from within a directory for projects, clone the repository:
```
git clone https://github.com/Tcleary/ExcelToXML.git
```
* change into the cloned directory, Ex. C:\Users\Username\Directory\ExcelToXML:
```
cd C:\Users\Username\Directory\ExcelToXML
```
* Create and load virtual environment in a subdirectory, for example we created one called "venv"
```
virtualenv venv
```
* Activate the virtual environment (it will say venv infront of the directory location on the command prompt, that is how you know you are in the virtual environment)
```
venv\Scripts\activate
```
* install the required python libraries listed in the requirements.txt file
```
pip install -r requirements.txt
```
* Run the script:
```
python run.py
```
Follow instructions instruction in the script to find and process excel file
