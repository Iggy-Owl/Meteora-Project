"""
==================================================================================================================
 This class is used for calculating ZHR based on 1 hour chunks
==================================================================================================================
"""

import pandas as pd


class ZHRCalculationChunk:
    breaks = pd.DataFrame()
    sky_obscured = pd.DataFrame()
    limiting_magnitude = pd.DataFrame()
    meteors_hour = pd.DataFrame()

    def __init__(self, breaks, sky_obscured, limiting_magnitude, meteors_hour):
        self.breaks = breaks
        self.sky_obscured = sky_obscured
        self.limiting_magnitude = limiting_magnitude
        self.meteors_hour = meteors_hour


file = "TestData/Test1.xlsx"

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

test = ZHRCalculationChunk(break_df, sky_obscured_df, limitingMagnitudeDF, meteorsDF)

i = 0
time_intervals_list =list()

for time in meteorsDF.Time:
    s = meteorsDF.Time[i].strftime("%H:%M")
    meteorsDF.Time[i] = s
    i = i + 1

S = pd.to_datetime(meteorsDF.Time)
time_interval = [g.reset_index(drop=True)
                 for i, g in
                 meteorsDF.groupby([(S - S[0]).astype('timedelta64[h]')])]
time_intervals_list.append(time_interval)

print(time_intervals_list)
