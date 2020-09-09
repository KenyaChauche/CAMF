# First On Scene Identifier tool for Camano Island Fire Department
# This tool will process a call records document and return that
# document with some columns added. These columns will show which
# units were the first on scene for that incident. There will be
# one column where string values of "yes" or "no" will dictate
# whether the unit of that row was the first on scene. Another
# column will display the unit that is first on scene, and the
# next column will display the turnout time of that unit. This
# will only do this for the row of that unit, all other rows for
# that incident will be left blank.

import pandas as pd
import numpy as np

from datetime import timedelta
from itertools import compress

try:
    header_index = 2

    # gather inputs from user
    print('We will ask you for three inputs.')
    while True:
        # read input
        path = input("Copy/ paste file path: ")
        if path == "":
            print("File path missing value, please double check filepath")
            continue

        # cleaning
        path = path.rstrip()

        path = path.replace('"', '')

        # read in file

        try:
            df = pd.read_excel(path, header = header_index)
        except:
            print("Error with file path, please double check filepath")
            continue
        break


    # process excel spreadsheet
    print("Determining First On Scene Units...")

    try:
        df = pd.read_excel(path, header = header_index)

        df = df.iloc[0:len(df) - header_index]

        # create dictionary of incident numbers with that incident's first on scene unit
        fos_units = {}

        def arrival_rank(incident_number, apparatus):
            try:
                ordered_list = list(df[df['Incident Number'] == incident_number].sort_values(by = ['Arrival Date'])['Apparatus Name'].values)
                return int(ordered_list.index(apparatus) + 1)
            except:
                return None

        df['Rank of Arrival'] = [arrival_rank(df.loc[i, 'Incident Number'], df.loc[i, 'Apparatus Name']) for i in df.index]

        df['Number Apparatuses Involved'] = [len(df[df['Incident Number'] == df.loc[i, 'Incident Number']]) for i in df.index]

        for i in df["Incident Number"]:
            time_list = list(df[df['Incident Number'] == i]['Arrival Date'])
            bool_list = [not b for b in list(pd.isnull(time_list))]
            time_list = list(compress(time_list, bool_list))
            try:
                min_time = min(time_list)
            except:
                min_time = None
            unit = df[(df['Arrival Date'] == min_time) & (df['Incident Number'] == i)]['Apparatus Name'].to_numpy()
            fos_units[i] = unit

        print("Processing Document...")

        filler = [None for i in df.index]

        df['Turn Out (seconds)'] = filler
        df['Response Time (seconds)'] = filler
        df['Travel Time (seconds)'] = filler

        df['FOS unit'] = filler
        df['FOS Turn Out (seconds)'] = filler
        df['FOS Response Time (seconds)'] = filler
        df['FOS Travel Time (seconds)'] = filler
        df['Is FOS'] = filler
        df['Incident Turn Out Goal Met'] = filler

        goal = 9 * 60 + 30

        for i in df.index:
            inum = df.loc[i, 'Incident Number']
            try:
                df['Turn Out (seconds)'] = [pd.Timedelta((df.loc[i, 'En Route Date']- df.loc[i, 'Dispatched Date'])).seconds for i in df.index]
                df['Response Time (seconds)'] = [pd.Timedelta((df.loc[i, 'Arrival Date'] - df.loc[i, 'Dispatched Date'])).seconds for i in df.index]
                df['Travel Time (seconds)'] = [pd.Timedelta((df.loc[i, 'Arrival Date'] - df.loc[i, 'En Route Date'])).seconds for i in df.index]

                fos = fos_units[inum][0]
                df_slice = df[(df['Incident Number'] == inum) & (df['Apparatus Name'] == fos)]

                if df.loc[i, 'Apparatus Name'] == fos:
                    df.loc[i, 'FOS unit'] = fos
                    df.loc[i, 'FOS Turn Out (seconds)'] = df_slice['Turn Out (seconds)'].to_numpy()[0]
                    df.loc[i, 'FOS Response Time (seconds)'] = df_slice['Response Time (seconds)'].to_numpy()[0]
                    df.loc[i, 'FOS Travel Time (seconds)'] = df_slice['Travel Time (seconds)'].to_numpy()[0]
                    df.loc[i, 'Is FOS'] = 'YES'
                    if df.loc[i, 'FOS Response Time (seconds)'] <= goal:
                        df.loc[i, 'Incident Turn Out Goal Met'] = 'YES'
                    else:
                        df.loc[i, 'Incident Turn Out Goal Met'] = 'NO'
                else:
                    df.loc[i, 'Is FOS'] = 'NO'
            except:
                pass
    except Exception as e:
        print(e)


    while True:
        file_name = input("What would you like the file called? ")
        if file_name == "":
            print("File name empty, please enter a value.")
            continue

        destination = input("Where would you like the file to go? ")
        if destination == "":
            print("File path missing value, please double check destination filepath and re-enter file name and destination filepath")
            continue

        # cleaning

        destination = destination.rstrip()

        destination = destination.replace('"', '')

        if not destination.endswith("\\"):
            destination = destination + "\\"

        file_name = file_name.rstrip()

        if file_name.endswith(".xls"):
            file_name.remove(".xls")

        if ".xlsx" not in file_name:
            file_name = file_name + ".xlsx"

        try:
            df.to_excel(f"{destination + file_name}")
        except Exception as e:
            print(e)
            print("An error has occurred. Please re-enter the folder destination and file name.")
            continue

        break


    name = (f'{destination + file_name}')

    writer = pd.ExcelWriter(f'{name}', engine = 'xlsxwriter')

    df.to_excel(writer, index = False)

    writer.save()

    print("Done!")
except Exception as e:
    print(e)
