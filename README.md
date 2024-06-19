# Python Victron Modbus Excel Converter V1

This is a python script, that uses the Victron Excel file to create ready to use entries for the Homeassistant Modbus section in the configruation.yaml file.

# Python dependencies:

- Pandas 
```
Ubuntu:
sudo apt install python3-pandas

Fedora:
sudo dnf install python3-pandas

For venv:
pip3 install pandas
```


# Operational dependencies:

- Victron Excel Sheet (can optionally be downloaded by the script --> ToDo)
- User acces to the Victron GX UI to get modbus device addresses


# Workflow:

1. Specify (local) or download (online) the latest Victron Modbus register excel sheet
2. Selection to list all registers from the file or skip this step
3. User input which register indices to use as colon separated string
4. User input for the corresponding Victron Modbus addresses (Victron GX settings UI)
5. Output is written to the console, optionally it can be saved as file with timestamp.


Tested with the lates Victron Modbus sheet (Rev 48) and python3.


# ToDo:

- Download of the Victron sheet for Github if not specified
- Specify local sheet as parameter
- Proper error handling