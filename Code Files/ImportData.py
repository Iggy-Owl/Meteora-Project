"""
==================================================================================================================
This code enables the import of data and processes it in order to make the needed calculations later on
==================================================================================================================
"""
import openpyxl
import pandas as pd
import sys
import glob
from PySide2 import QtWidgets
from datetime import timedelta

pd.set_option('display.max_columns', None)
# Opens dialog window in order to import the folder with the Excel files, containing the meteor data for a given date.

app = QtWidgets.QApplication(sys.argv)
window = QtWidgets.QMainWindow()
DirectoryPath = QtWidgets.QFileDialog.getExistingDirectory(None, "Select Directory") \
 \
 # Declaring data frames containing the information for each observer (file).
# The data they contain (columns) is described down below.

break_df = pd.DataFrame()  # Index number, Start(UT), End(UT)
sky_obscured_df = pd.DataFrame()  # Index number, Time of Assessment (UT), Obscurity Coefficient (K)
limitingMagnitudeDF = pd.DataFrame()  # Index number, Time of Assessment (UT), Limiting Magnitude (LM)
centerOfFieldDF = pd.DataFrame()  # Index number, Right  Ascension, Declination
meteorsDF = pd.DataFrame()  # Index number, Time of Occurrence(UT), Visible Magnitude (Mv), Velocity (V) (deg/sec),
# Shower, Map, Begin X, Begin Y, End X, End Y (coordinates), Dead Time

# Declaring data frames, which stitch the above-mentioned data together in single master charts

meteors_MasterDF = pd.DataFrame()

# Filling the Master Data Frames with the data from the separate files
for file in glob.glob(DirectoryPath + "/*.xlsx"):
    break_df = pd.read_excel(file, sheet_name='ObservationData',
                             skiprows=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], usecols="A,B,C",
                             header=None)

    sky_obscured_df = pd.read_excel(file, sheet_name='ObservationData',
                                    skiprows=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], usecols="E,F,G",
                                    header=None)

    limitingMagnitudeDF = pd.read_excel(file, sheet_name='ObservationData',
                                        skiprows=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], usecols="I,J,K",
                                        header=None)

    centerOfFieldDF = pd.read_excel(file, sheet_name='ObservationData',
                                    skiprows=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14], usecols="M,N,O",
                                    header=None)

    meteorsDF = pd.read_excel(file, sheet_name='Meteors', skiprows=[1])
    meteors_MasterDF = meteors_MasterDF.append(meteorsDF, ignore_index=True)

    # PROCESSING DATA FOR ESTIMATING ZHR (ZENITH HOURLY RATE)

    # Estimating active time between the start of the observation and the occurrence of the first seen meteor

    workBook = openpyxl.load_workbook(file)  # Loads the Excel Workbook

    # Getting the times of the start and the end of the observation
    ObservationData = workBook["ObservationData"]
    observationStartUT = ObservationData['A6'].value
    observationEndUT = ObservationData['B6'].value

    # Getting the times of the first and the last meteors seen.
    Meteors = workBook["Meteors"]
    firstMeteorSighted = Meteors['B3'].value
    lastedMeteorSighted = Meteors['B' + str(len(list(Meteors.rows)))].value  # Gets the value from the last row.

    # Active time elapsed between the start of the observation and occurrence of the first meteor.
    startToFirst = (timedelta(hours=firstMeteorSighted.hour, minutes=firstMeteorSighted.minute) - timedelta(
        hours=observationStartUT.hour, minutes=observationStartUT.minute)).total_seconds()

    # Active time elapsed between the end of the observation and occurrence of the last meteor.
    endToLast = (timedelta(hours=observationEndUT.hour, minutes=observationEndUT.minute) - timedelta(
        hours=lastedMeteorSighted.hour, minutes=lastedMeteorSighted.minute)).total_seconds()
