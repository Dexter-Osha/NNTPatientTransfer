import pandas as pd
import openpyxl
import os
import keyboard
import time

# Main File of script, EDIT AT OWN RISK.
if __name__ == '__main__':
    # Basic way of importing .csv into python.
    file_path = 'C:/PatientTransfer/PatientMapping.xlsx'
    newFile = "C:\PatientTransfer\PatientEntryCommands.txt"
    errorFile= "C:\PatientTransfer\PatientEntryErrors.txt"
    df = pd.read_excel(file_path)
    # create a blank patient profile to be updated with each patient's corresponding info
    rawNNTPatient = {"Unique Index": "n/a", "/patid": "n/a", "/name": "n/a", "/surname": "n/a", "/dateb": "n/a",
                  "/SEX": "n/a"}

# Iterate through all rows in Excel files
    for row in df.itertuples():
        # Blank list which will be used to update NNT Patient object
        patientInfoList = []
        # Iterate through each row inside of Excel file, adding all patient info into list used to update NNT Patient object
        for value in range(len(row)):
            keyValue = row[value]
            patientInfoList.append(keyValue)
        # Update NNT Patient Object with values saved into List, only grabbing the keys we need for successful import.
        for i in range(len(patientInfoList)):
            rawNNTPatient.update(zip(rawNNTPatient.keys(), patientInfoList))
            NNTPatient = dict((k, rawNNTPatient[k]) for k in ['/patid', '/name', '/surname', '/SEX'] if k in rawNNTPatient)
        # Find the date of birth(Timestamp object) for each patient and format it into desired string
        for key, value in rawNNTPatient.items():
            if key == "/dateb":
                dateOfBirth_Timestamp = value
                formattedDateOfBirth = (dateOfBirth_Timestamp.strftime('%d,%m,%Y'))
        #Output patient entry command for each patient in newly created file.
        with open(newFile, 'a') as file:
            file.write("C:\\NNT\\NNTBridge.exe " + (str(NNTPatient).replace(":","").replace("'", "").replace(",","").strip("{}")) +
                  " /dateb " + formattedDateOfBirth + "\n")





