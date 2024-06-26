import os.path
import os
import time
import pandas
import requests

colCharSize_registerOverviewIndex = 10
colCharSize_registerDescription = 60
colCharSize_registerModbusAdr = 20
colCharSize_registerDbusObjPath = 60
registerOverviewPagesEntryCount = 30
victronExcelFileHeaderRowIndexNumber = 1
baseStr = "modbus:\n  - name: victron\n    type: tcp\n    host: x.x.x.x\n    port: 502\n    sensors:"
fileURL = "https://raw.githubusercontent.com/victronenergy/dbus_modbustcp/master/CCGX-Modbus-TCP-register-list.xlsx"
fileTarget = "ModbusRegisterList.xlsx"
spaces = '      '

def cls(): os.system('cls' if os.name=='nt' else 'clear')

def fillStringUpWithSpaces(inputStr: str, size: int) -> str:
    tmpStr = inputStr
    origSize = len(inputStr)
    fillUpAmount = size - origSize

    for x in range(fillUpAmount):
        tmpStr += " "

    return tmpStr

def parseExcelToDict(filepath: str, sheetName: str, headerRowNr: int):
    excel_data_df = pandas.read_excel(filepath, sheet_name = sheetName, header = headerRowNr)
    return excel_data_df.to_dict()

def convertDictEntryToHassString(xlsxDict: dict, row: int, unitIDDict: dict) -> str:
    expString = "    - "

    value_desc = xlsxDict.get('description').get(int(row))
    if value_desc is None:
        print("Row " + row + " has no valid descriptor, please check the file. This row is now skipped.")
        return ""
    expString += 'name: \'Victron ' + str(value_desc) + '\'\n'

    serviceStr = xlsxDict.get('dbus-service-name').get(int(row))
    value_slave = unitIDDict.get(serviceStr)
    expString += spaces + 'slave: ' + str(value_slave) + '\n'

    value_adr = xlsxDict.get('Address').get(int(row))
    expString += spaces + 'address: ' + str(value_adr) + '\n'

    value_type = xlsxDict.get('Type').get(int(row))
    if value_type == 'int16' or value_type == 'uint16' or value_type == 'int64' or value_type == 'uint64' or value_type == 'int32' or value_type == 'uint32':
        expString += spaces + 'data_type: ' + str(value_type) + '\n'

    value_scale = xlsxDict.get('Scalefactor').get(int(row))
    expString += spaces + 'scale: ' + str(f'{(1 / value_scale):g}') + '\n'

    value_unit = xlsxDict.get('dbus-unit').get(int(row))
    value_unit = value_unit.split()[0]
    if len(value_unit) <= 5:        
        value_deviceClass = ""
        value_precision = 0
        if value_unit == "W" or value_unit == "VA":
            value_deviceClass = "power"
            value_precision = 0
            expString += spaces + 'device_class: ' + str(value_deviceClass) + '\n'
            expString += spaces + 'precision: ' + str(value_precision) + '\n'
            value_unit = "W"
        elif value_unit == "kWh":
            value_deviceClass = "energy"
            value_precision = 1
            expString += spaces + 'device_class: ' + str(value_deviceClass) + '\n'
            expString += spaces + 'precision: ' + str(value_precision) + '\n'
        elif value_unit == "V":
            value_deviceClass = "voltage"
            value_precision = 1
            expString += spaces + 'device_class: ' + str(value_deviceClass) + '\n'
            expString += spaces + 'precision: ' + str(value_precision) + '\n'
        elif value_unit == "A":
            value_deviceClass = "current"
            value_precision = 1
            expString += spaces + 'device_class: ' + str(value_deviceClass) + '\n'
            expString += spaces + 'precision: ' + str(value_precision) + '\n'
        
        expString += spaces + 'unit_of_measurement: ' + '\"' + value_unit + '\"' + '\n'
        
    else:
        expString += spaces + '# Unit:\t' + str(value_unit) + '\n'

    # expString += spaces + 'scan_interval: 5' + '\n'   # Maybe for features, default is 15sec

    value_range = xlsxDict.get('Range').get(int(row))
    expString += spaces + '# Range:\t' + str(value_range) + '\n'

    return expString

def getAllCellValuesFromColumn(xlsxDict: dict, colName: str):
    counter = 0
    tmpDict = {}
    for entryName in xlsxDict[colName].values():
        tmpDict[counter] = entryName
        counter += 1
    return tmpDict

def main():
    cls()
    print("----------------------------------------------------------------------------------------------------------------------------")
    print("Welcome to the Victron GX Modbus Register Converter for Homeassistant configuration.yaml")
    print("First we need the actual CCGX Modubus Register Definition file (xlsx file) to parse the available entries.")
    print("----------------------------------------------------------------------------------------------------------------------------")

    fileValidFlag = False
    inputStr = input("Do you want to load the latest CCGX Modbus register file from the Victron Github repository? [Y/n]: ")
    if inputStr.lower() == "y" or len(inputStr) == 0:
        try:
            r = requests.get(fileURL, allow_redirects=True)
            open(fileTarget, 'wb').write(r.content)
            filepath = "./" + fileTarget
            print("File download successful")
            fileValidFlag = True
        except:
            print("Something went wrong while loading the file, aborting the script.")
            exit(-2)
    else:
        print("Please get it and specify its location here:")        
        while fileValidFlag == False:
            filepath = input("File: ")
            if not os.path.isfile(filepath):
                print("Specified file does not exist, try again.")
            elif 'xlsx' not in filepath:
                print("Specified file is not a xlsx file, try again.")
            else:
                print("Ok, the specified file is valid")
                fileValidFlag = True

    dict_complete = parseExcelToDict(filepath, 'Field list', victronExcelFileHeaderRowIndexNumber)
    dict_names = getAllCellValuesFromColumn(dict_complete, 'description')
    dict_paths = getAllCellValuesFromColumn(dict_complete, 'dbus-obj-path')
    dict_modbusAdr = getAllCellValuesFromColumn(dict_complete, 'Address')
    dict_serviceName = getAllCellValuesFromColumn(dict_complete, 'dbus-service-name')
    list_serviceNamesSelected = []
    countEntries = len(dict_names)

    print("I have found " + str(countEntries) + " registers in the list")
    print("Now we should look for the registers, that you want to use")
    inputStr = input("Should I display the register list so that you can look for your entries? [Y/n]: ")
    cls()

    if inputStr.lower() == "y" or len(inputStr) == 0:
        print("Ok, now I'll list all registers with index numbers in pages of " + str(registerOverviewPagesEntryCount) + " items")
        print("Please note the indices of the entry that you wanna use")
        print("Next page with any key, to abort listing type \"end\"")
        input()
        cls()
        
        for pageStartIndex in range(0, countEntries, registerOverviewPagesEntryCount):
            print(fillStringUpWithSpaces("Index", colCharSize_registerOverviewIndex) + \
                  fillStringUpWithSpaces("Address", colCharSize_registerModbusAdr) + \
                  fillStringUpWithSpaces("Description", colCharSize_registerDescription) + \
                  fillStringUpWithSpaces("Modbus Path", colCharSize_registerDbusObjPath))
            print("----------------------------------------------------------------------------------------------------------------------------")
            for entryIndex in range(pageStartIndex, pageStartIndex + registerOverviewPagesEntryCount):
                if entryIndex < (countEntries):
                    print(fillStringUpWithSpaces(str(entryIndex), colCharSize_registerOverviewIndex) + \
                          fillStringUpWithSpaces(str(dict_modbusAdr[entryIndex]), colCharSize_registerModbusAdr) + \
                          fillStringUpWithSpaces(str(dict_names[entryIndex]), colCharSize_registerDescription) + \
                          fillStringUpWithSpaces(str(dict_paths[entryIndex]), colCharSize_registerDbusObjPath))
            print("")
            userInput = input()
            if len(userInput) > 0 and "end" in userInput.lower():
                break
            cls()

    cls()

    print("Ok, now we need the register indices, that I should convert.")
    indicesSelectedStr = input("Please fill in the index numbers separated by colon \',\': ")
    indicesSelectedList = indicesSelectedStr.split(',')

    if len(indicesSelectedList) <= 0:
        print("Wrong selection, aborting program")
        exit(-1)

    for selectedEntry in indicesSelectedList:
        if dict_serviceName.get(int(selectedEntry)) != None and dict_serviceName.get(int(selectedEntry)) not in list_serviceNamesSelected:
            list_serviceNamesSelected.append(dict_serviceName.get(int(selectedEntry)))

    cls()
    print("Selection accepted, you have selected items with " + str(len(list_serviceNamesSelected)) + " different service name(s).")
    print("For these items we need the Unit IDs from your GX System.")
    print("Please log in to the UI of your GX device and browse to Settings => Services => ModbusTCP => Set it to active and click on \'Available services\'")
    print("There look for the foollowing service and type the corresponding Unit ID here as number:")
    print("")

    dict_idMappings = {} 
    for i in range(len(list_serviceNamesSelected)):
        inputStr = input(list_serviceNamesSelected[i] + ": ")
        if len(inputStr) > 0:
            dict_idMappings[list_serviceNamesSelected[i]] = inputStr
        else:
            print("Input error, aborting program")
            exit(-1)

    cls()
    
    print("Generated Output:")
    print("(Please remember to edit IP address, and port settings to your needs)")
    print("----------------------------------------------------------------------------------------------------------------------------")
    print("")
    outStr = baseStr + '\n'
    for selectedEntry in indicesSelectedList:
        outStr += convertDictEntryToHassString(dict_complete, selectedEntry, dict_idMappings) + '\n'
    print(outStr)
    print("")
    print("----------------------------------------------------------------------------------------------------------------------------")

    inputStr = input("Should the output be saved as file? [Y/n]: ")
    if inputStr.lower() == "y" or len(inputStr) == 0:
        filename = time.strftime("%Y.%m.%d_%H%M%S_modbus_configuration.yaml")
        text_file = open(filename, "w")
        text_file.write(outStr)
        text_file.close()
        print("Output is saved to: \'" + filename + "\'")


if __name__ == "__main__":
    main()