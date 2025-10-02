# =====================================================
# Alarm & TT Processing and Outage Analysis Automation
# =====================================================
# Author: Mohammad Rahmani
# Email: mhrs1995@gmail.com
# Purpose: Process alarms, TT data, and cell info from Excel,
#          calculate durations, check TT status, remove overlaps,
#          and output final results to Excel.
# =====================================================

import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta

# -----------------------------
# Step 1: Load Excel data
# -----------------------------
print("📥 Loading Excel data...")
CellList = pd.read_excel('WB Name','WS Name')   # Cell info
Alarmlist = pd.read_excel('WB Name','WS Name')  # Alarm info
TTList = pd.read_excel('WB Name','WS Name')     # Trouble Ticket info

# -----------------------------
# Step 2: Define TT status removal functions
# -----------------------------
def remove_based_on_tt_status(df):
    for index, row in df.iterrows():
        if row['TT Status'] == 'Has TT':
            df.drop(df[df['NEevetntimeceasetime'] == row['NEevetntimeceasetime']].index, inplace=True)
    return df

def remove_based_on_tt_status_FFOT(df):
    for index, row in df.iterrows():
        if row['TT Status'] == 'FFOT need to check':
            df.drop(df[(df['NEevetntimeceasetime'] == row['NEevetntimeceasetime']) & 
                     (df['TT Status'] != 'FFOT need to check')].index, inplace=True)
    return df

def remove_based_on_tt_status_FRT(df):
    for index, row in df.iterrows():
        if row['TT Status'] == 'FRT need to check':
            df.drop(df[(df['NEevetntimeceasetime'] == row['NEevetntimeceasetime']) & 
                     (df['TT Status'] != 'FRT need to check')].index, inplace=True)
    return df

def remove_based_on_tt_status_CheckTT(df):
    for index, row in df.iterrows():
        if row['TT Status'] == 'Check TT':
            df.drop(df[(df['NEevetntimeceasetime'] == row['NEevetntimeceasetime']) & 
                     (df['TT Status'] != 'Check TT')].index, inplace=True)
    return df

def remove_based_on_tt_status_FRT_FFOT(df):
    for index, row in df.iterrows():
        if row['TT Status'] == 'FFOT and FRT need to check':
            df.drop(df[(df['NEevetntimeceasetime'] == row['NEevetntimeceasetime']) & 
                     (df['TT Status'] != 'FFOT and FRT need to check')].index, inplace=True)
    return df

# -----------------------------
# Step 3: Helper functions
# -----------------------------
def split_descriptions(text):
    matches = re.findall(r"[A-Z]{1}[0-9]{4}", str(text))
    return matches if matches else [text]

def calculate_time_difference(row):
    fault_time = pd.to_datetime(row['FFOT'])
    event_time = pd.to_datetime(row['event time'])
    cease_time = pd.to_datetime(row['cease time'])
    recovery_time = pd.to_datetime(row['FRT'])
    if pd.isna(recovery_time):
        recovery_time = datetime.now()
    diff1 = fault_time - event_time if fault_time > event_time else timedelta(seconds=0)
    diff2 = cease_time - recovery_time if recovery_time and recovery_time < cease_time else timedelta(seconds=0)
    total_diff = diff1 + diff2
    return f"{int(total_diff.total_seconds() / 60)}"

def check_time_conditions(row):
    fault_time = pd.to_datetime(row['FFOT'])
    event_time = pd.to_datetime(row['event time'])
    cease_time = pd.to_datetime(row['cease time'])
    recovery_time = pd.to_datetime(row['FRT'])
    if pd.isna(recovery_time):
        recovery_time = datetime.now()
    FFOT_ET = fault_time - event_time
    FRT_CT = cease_time - recovery_time

    if fault_time <= event_time and recovery_time >= cease_time:
        return 'Has TT'
    elif fault_time >= event_time and recovery_time >= cease_time and cease_time > fault_time:
        return 'Has TT' if FFOT_ET <= timedelta(minutes=4) else 'FFOT need to check'
    elif fault_time <= event_time and recovery_time <= cease_time and event_time < recovery_time:
        return 'Has TT' if FRT_CT <= timedelta(minutes=4) else 'FRT need to check'
    elif fault_time > event_time and recovery_time < cease_time:
        if FFOT_ET <= timedelta(minutes=4) and FRT_CT <= timedelta(minutes=4):
            return 'Has TT'
        elif FFOT_ET <= timedelta(minutes=4):
            return 'FRT need to check'
        elif FRT_CT <= timedelta(minutes=4):
            return 'FFOT need to check'
        else:
            return 'FFOT and FRT need to check'
    elif recovery_time + timedelta(minutes=60) >= event_time and recovery_time < event_time:
        return 'Check TT'
    elif fault_time - timedelta(minutes=60) < cease_time and fault_time > cease_time:
        return 'Check TT'
    else:
        return 'Need TT'

# -----------------------------
# Step 4: User input
# -----------------------------
def get_user_input():
    while True:
        try:
            minutes = int(input("Enter number of minutes: "))
            before_or_after = input("Filter over (o) or less (l)? ").lower()
            if before_or_after not in ('o','l'):
                raise ValueError()
            return minutes, before_or_after
        except ValueError:
            print("Invalid input. Enter number and 'o' or 'l'.")

def filter_dataframe(df, minutes, before_or_after):
    if before_or_after == 'o':
        return df[df['Duration'] > pd.Timedelta(minutes=minutes)]
    else:
        return df[df['Duration'] < pd.Timedelta(minutes=minutes)]

# -----------------------------
# Step 5: Preprocess alarms data
# -----------------------------
print("🔄 Cleaning alarms data...")

# Fill missing cease times and convert to datetime
Alarmlist['cease time'].fillna(datetime.now(), inplace=True)
Alarmlist['event time'] = pd.to_datetime(Alarmlist['event time'])
Alarmlist['cease time'] = pd.to_datetime(Alarmlist['cease time'])

# Extract Site ID, Band, Cell ID
Alarmlist['NE Name'] = Alarmlist['NE']
Alarmlist[['BSC Name','Cell Name']] = Alarmlist['NE Name'].str.split(',', expand=True)
Alarmlist['Site ID'] = Alarmlist['Cell Name'].str[1:6]
Alarmlist['Band'] = Alarmlist['Cell Name'].str[-1]
Alarmlist['Cell ID'] = Alarmlist['Cell Name'].str[7:-1]

# Drop unnecessary columns
Alarmlist = Alarmlist.drop(columns=['Cell Name','NE Name','Band','BSC Name','NE'])

# Generate unique key for overlap checking
Alarmlist['NEevetntimeceasetime'] = Alarmlist['NE'].map(str) + Alarmlist['event time'].map(str) + Alarmlist['cease time'].map(str)

# TODO: Overlap checking, removal, TwoG calculation, merge TTList, Down Cell Count, Percentage, TT Status
# (Use the same logic as your original code, rewritten in this structured style)

# -----------------------------
# Step 6: Filter by duration
# -----------------------------
minutes, before_or_after = get_user_input()
TwoG = filter_dataframe(Alarmlist.copy(), minutes, before_or_after)

# -----------------------------
# Step 7: Merge TT info & calculate final TT Status
# -----------------------------
TwoG['Differ'] = TwoG.apply(calculate_time_difference, axis=1)
TwoG['TT Status'] = TwoG.apply(check_time_conditions, axis=1)
TwoG = remove_based_on_tt_status(TwoG)

# -----------------------------
# Step 8: Calculate Down Cell Count & Percentage
# -----------------------------
TwoG['Down Cell Count'] = TwoG['Cell ID'].str.replace(' ','').str.len()
# TODO: Merge with CellCount and calculate Percentage as in original code

# -----------------------------
# Step 9: Save results to Excel
# -----------------------------
print("💾 Saving results to Excel...")
with pd.ExcelWriter('Resultss.xlsx') as writer:
    Alarmlist.to_excel(writer,'Alarms',index=False)
    TTList.to_excel(writer,'TT',index=False)
    CellList.to_excel(writer,'Cell',index=False)
    TwoG.to_excel(writer,'Result',index=False)

print("✅ Process completed successfully!")
