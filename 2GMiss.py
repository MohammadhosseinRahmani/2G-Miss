from joblib import PrintTime
import pandas as pd
import numpy as np
import datetime 
from datetime import datetime
from datetime import timedelta
import re
from sqlalchemy import null
from sympy import subsets
print('Get data from Excel')
#get data from Excel
CellList = pd.read_excel('Sample FINAL 50% ORIGINAL.xlsx','Cells')
Alarmlist = pd.read_excel('Sample FINAL 50% ORIGINAL.xlsx','Alarms')
TTList = pd.read_excel('Sample FINAL 50% ORIGINAL.xlsx','TTs')

#remove Row for has TT
def remove_based_on_tt_status(df):
  for index, row in df.iterrows():
    # Check if TT Status is "Has TT"
    if row['TT Status'] == 'Has TT':
      value_to_remove = row['NEevetntimeceasetime']
      # Filter and drop rows with matching NEevetntimeceasetime
      df.drop(df[df['NEevetntimeceasetime'] == value_to_remove].index, inplace=True)
  return df

#Remove Row for FFOT
def remove_based_on_tt_status_FFOT(df):
  for index, row in df.iterrows():
    # Check if TT Status is "Has TT"
    if row['TT Status'] == 'FFOT need to check':
      value_to_Keep = row['NEevetntimeceasetime']
      # Filter and drop rows with matching NEevetntimeceasetime
      df = df.drop(df[(df['NEevetntimeceasetime'] == value_to_Keep) & (df['TT Status'] != 'FFOT need to check')].index, inplace=True)
  return df
#Remove Row for FRT
def remove_based_on_tt_status_FRT(df):
  for index, row in df.iterrows():
    # Check if TT Status is "Has TT"
    if row['TT Status'] == 'FRT need to check':
      value_to_Keep = row['NEevetntimeceasetime']
      # Filter and drop rows with matching NEevetntimeceasetime
      df = df.drop(df[(df['NEevetntimeceasetime'] == value_to_Keep) & (df['TT Status']!= 'FRT need to check')].index, inplace=True)
  return df

#Remove Row for Check TT
def remove_based_on_tt_status_CheckTT(df):
  for index, row in df.iterrows():
    # Check if TT Status is "Has TT"
    if row['TT Status'] == 'Check TT':
      value_to_Keep = row['NEevetntimeceasetime']
      # Filter and drop rows with matching NEevetntimeceasetime
      df = df.drop(df[(df['NEevetntimeceasetime'] == value_to_Keep) & (df['TT Status']!= 'Check TT')].index, inplace=True)
  return df

#Remove Row for FRT and FFOT
def remove_based_on_tt_status_FRT_FFOT(df):
  for index, row in df.iterrows():
    # Check if TT Status is "Has TT"
    if row['TT Status'] == 'FFOT and FRT need to check':
      value_to_Keep = row['NEevetntimeceasetime']
      # Filter and drop rows with matching NEevetntimeceasetime
      df = df.drop(df[(df['NEevetntimeceasetime'] == value_to_Keep) & (df['TT Status']!= 'FFOT and FRT need to check')].index, inplace=True)
  return df

# Define function to extract text (without splitting)
def split_descriptions(text):
  """Splits text containing the format "[A-Z]{1}[0-9]{4}" into a list.

  Args:
      text: The text to split.

  Returns:
      A list of strings where each element has the format "[A-Z]{1}[0-9]{4}"
      or the original text if no matches are found.
  """
  matches = re.findall(r"[A-Z]{1}[0-9]{4}", text)
  return matches if matches else [text]

def calculate_time_difference(row):
  fault_time = pd.to_datetime(row['FFOT'])
  event_time = pd.to_datetime(row['event time'])
  cease_time = pd.to_datetime(row['cease time'])
  recovery_time = pd.to_datetime(row['FRT'])
  if pd.isna(recovery_time):  # Check for missing values (NaT)
    recovery_time = datetime.now()   

  # Calculate time difference between event and fault (ensure fault is after event)
  if fault_time > event_time:
    time_diff_1 = fault_time - event_time
  else:
    time_diff_1 = timedelta(seconds=0)  # Set to zero if fault is not after event

  # Calculate time difference between recovery and cease (ensure recovery is before cease)
  if recovery_time and recovery_time < cease_time:  # Check if recovery time exists and is before cease time
    time_diff_2 = cease_time - recovery_time
  else:
    time_diff_2 = timedelta(seconds=0)  # Set to zero if not both conditions are met

  # Total time difference in seconds
  total_difference = time_diff_1 + time_diff_2

  # Convert total difference to minutes and format as string
  minutes_diff = total_difference.total_seconds() / 60
  formatted_minutes = f"{int(minutes_diff)}"

  # Return the sum of time differences
  return formatted_minutes


# Assuming your DataFrame is called 'df'
def check_time_conditions(row):
  """
  Checks time conditions for fault time (FFOT), event time, cease time, and recovery time (FRT).

  Args:
      row (pandas.Series): A row of data from the DataFrame.

  Returns:
      str: The verdict based on time conditions: 'TT Found', 'FFOT need to check', 'FRT need to check', 
           'FFOT and FRT need to check', or 'Check TT'.
  """
 
  fault_time = pd.to_datetime(row['FFOT'])
  event_time = pd.to_datetime(row['event time'])
  cease_time = pd.to_datetime(row['cease time'])

  # Handle missing values in ResolvedFault Recovery Time
  recovery_time = pd.to_datetime(row['FRT'])
  if pd.isna(recovery_time):  # Check for missing values (NaT)
      recovery_time = datetime.now()  # Fill with current local date and time
  FFOT_ET = fault_time - event_time
  FRT_CT = cease_time-recovery_time
  # Set verdict based on additional conditions
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
            return 'FRT Need to check'
        elif FRT_CT <= timedelta(minutes=4):
            return 'FFOT need to check'
        else:
            return 'GFFOT and FRT need to check'
  elif recovery_time + timedelta(minutes=60) >= event_time and recovery_time < event_time:
        return 'HCheck TT'
  elif fault_time - timedelta(minutes=60) < cease_time and fault_time > cease_time:
        return 'HCheck TT'
  else:
        return 'Need TT'
  

def split_and_update(data):
  # Cast to string before splitting (assuming it's not already a string)
  if not pd.api.types.is_string_dtype(data):
    data = str(data)
  # Split the string, remove duplicates using set, and convert back to list
  data_split = list(set(data.split(',')))
  return data_split

def overlap_checker_tmp(AL):
        AL['NEevetntimeceasetime']=AL['NE'].map(str)+AL['event time'].map(str)+AL['cease time'].map(str)
        df2 = pd.merge(AL, AL, on='NE')
        df2['Overlap']=np.where((df2['event time_x']<=df2['cease time_y'])&(df2['cease time_x']>=df2['event time_y']) & (df2['NEevetntimeceasetime_x'] != df2['NEevetntimeceasetime_y']), 'Overlapped','not overlapped')
        df2['event time_x']=np.where(df2['Overlap'].eq('Overlapped'),df2[['event time_x','event time_y']].min(axis=1),df2['event time_x'])
        df2['cease time_x']=np.where(df2['Overlap'].eq('Overlapped'),df2[['cease time_x','cease time_y']].max(axis=1),df2['cease time_x'])
        df2['Cell ID']=df2['Cell ID_x'].map(str)+df2['Cell ID_y'].map(str)
        df2['Cell ID'] = df2['Cell ID'].apply(lambda x: ' '.join(sorted(set(x))))
        df2.rename(columns = {'event time_x':'event time','cease time_x':'cease time'}, inplace = True)
        return df2, AL

def overlap_checker(AL):
        AL['NEevetntimeceasetime']=AL['Site ID'].map(str)+AL['event time'].map(str)+AL['cease time'].map(str)
        df2 = pd.merge(AL, AL, on='Site ID')
        df2['Overlap']=np.where((df2['event time_x']<=df2['cease time_y'])&(df2['cease time_x']>=df2['event time_y']) & (df2['NEevetntimeceasetime_x'] != df2['NEevetntimeceasetime_y']), 'Overlapped','not overlapped')
        df2['event time_x']=np.where(df2['Overlap'].eq('Overlapped'),df2[['event time_x','event time_y']].min(axis=1),df2['event time_x'])
        df2['cease time_x']=np.where(df2['Overlap'].eq('Overlapped'),df2[['cease time_x','cease time_y']].max(axis=1),df2['cease time_x'])
        df2['Cell ID']=df2['Cell ID_x'].map(str)+df2['Cell ID_y'].map(str)
        df2['Cell ID'] = df2['Cell ID'].apply(lambda x: ' '.join(sorted(set(x))))
        df2.rename(columns = {'event time_x':'event time','cease time_x':'cease time','BSC Name_x':'BSC Name'}, inplace = True)
        return df2, AL

def overlap_remover_tmp(df, AL):
        df2= df[(df['Overlap']=="Overlapped")]
        AL1=AL[~AL['NEevetntimeceasetime'].isin(df2['NEevetntimeceasetime_x'])]
        df2 = df2.drop(columns=['event time_y','cease time_y','Overlap','NEevetntimeceasetime_x','NEevetntimeceasetime_y','Cell ID_x','Cell ID_y'])
        df2 = df2.drop_duplicates()
        bigdata = pd.concat([df2,AL1],ignore_index=True,sort=False)
        return bigdata

def overlap_remover(df, AL):
        df2= df[(df['Overlap']=="Overlapped")]
        AL1=AL[~AL['NEevetntimeceasetime'].isin(df2['NEevetntimeceasetime_x'])]
        df2 = df2.drop(columns=['event time_y','cease time_y','Overlap','NEevetntimeceasetime_x','NEevetntimeceasetime_y','Cell ID_x','Cell ID_y', 'BSC Name_y'])
        df2 = df2.drop_duplicates()
        bigdata = pd.concat([df2,AL1],ignore_index=True,sort=False)
        return bigdata

def get_user_input():
    """Gets user input for minutes and before/after selection."""

    while True:
        try:
            minutes = int(input("Enter the number of minutes: "))
            before_or_after = input("Filter over (o) or Less (l) this time? ").lower()
            if before_or_after not in ('o', 'l'):
                raise ValueError("Invalid choice. Please enter 'o' or 'l'.")
            return minutes, before_or_after
        except ValueError:
            print("Invalid input. Please enter a number for minutes and 'o' or 'l' for before/after.")

def filter_dataframe(df, minutes, before_or_after):
    """Filters the DataFrame based on user input."""

    # Assuming your DataFrame has a datetime column named 'event time'
    reference_time = pd.to_datetime('now')
    if before_or_after == 'o':
        filtered_df = df[df['Duration'] > pd.Timedelta(minutes=minutes)]
    else:
        filtered_df = df[df['Duration'] < pd.Timedelta(minutes=minutes)]
    return filtered_df


def filter_rows(group):
    if any(group['TT Status'] != 'Need TT'):
        return group[group['TT Status'] != 'Need TT']
    else:
        return group.drop_duplicates(subset=['NEevetntimeceasetime'])
    
Alarmlist['cease time'].fillna(datetime.now(), inplace=True)
Alarmlist['event time'] = pd.to_datetime(Alarmlist['event time'])
Alarmlist['cease time'] = pd.to_datetime(Alarmlist['cease time'])
Alarmlist['NE Name'] = Alarmlist['NE'] 
Alarmlist['NE Name'] = Alarmlist.loc[:, 'NE']
Alarmlist[['BSC Name', 'Cell Name']] = Alarmlist['NE Name'].str.split(',', expand=True)
Alarmlist = Alarmlist[~(Alarmlist['NE'].str.contains('B0365'))]
Alarmlist = Alarmlist[~(Alarmlist['NE'].str.contains('E0340'))]
Alarmlist = Alarmlist[~(Alarmlist['NE'].str.contains('B0809'))]
Alarmlist = Alarmlist[~(Alarmlist['NE'].str.contains('T3566'))]
Alarmlist['Site ID'] = Alarmlist['Cell Name'].str.slice(start=1,stop=6)
Alarmlist['Band'] = Alarmlist['Cell Name'].str[-1]
Alarmlist['Cell ID'] = Alarmlist['Cell Name'].str.slice(start=7)
Alarmlist['Cell ID'] = Alarmlist['Cell ID'].str[:-1]
Alarmlist = Alarmlist.drop(columns=['Cell Name','NE Name', 'Band','Site ID','BSC Name'])
dftmp, AlarmList = overlap_checker_tmp(Alarmlist)
while dftmp['Overlap'].str.contains('Overlapped').any():
    Alarmlist = overlap_remover_tmp(dftmp,AlarmList)
    dftmp, AlarmList = overlap_checker_tmp(Alarmlist)

Alarmlist['event time'] = pd.to_datetime(Alarmlist['event time'])
Alarmlist['cease time'] = pd.to_datetime(Alarmlist['cease time'])
Alarmlist['NE Name'] = Alarmlist['NE'] 
Alarmlist['NE Name'] = Alarmlist.loc[:, 'NE']
Alarmlist[['BSC Name', 'Cell Name']] = Alarmlist['NE Name'].str.split(',', expand=True)
Alarmlist['Site ID'] = Alarmlist['Cell Name'].str.slice(start=1,stop=6)
Alarmlist['Band'] = Alarmlist['Cell Name'].str[-1]
Alarmlist['Cell ID'] = Alarmlist['Cell Name'].str.slice(start=7)
Alarmlist['Cell ID'] = Alarmlist['Cell ID'].str[:-1]
Alarmlist['Site ID'] = Alarmlist['Site ID'].astype(str)+"_"+Alarmlist['Band']
Alarmlist = Alarmlist.drop(columns=['Cell Name','NE Name','NE', 'Band'])
dftmp, AlarmList = overlap_checker(Alarmlist)
while dftmp['Overlap'].str.contains('Overlapped').any():
    Alarmlist = overlap_remover(dftmp,AlarmList)
    dftmp, AlarmList = overlap_checker(Alarmlist)

TwoG = Alarmlist.groupby(["NEevetntimeceasetime"]).agg({
    "Cell ID": lambda x: ', '.join(set(x)),  # Combine unique Cell IDs
    "Site ID": 'first',
    "event time": 'max',
    "cease time": 'max',
    "BSC Name": 'first'
}).reset_index()

TwoG["Cell ID"] = TwoG["Cell ID"].apply(lambda x: ' '.join(sorted(x.split(', '))))
TwoG["Cell ID"] = TwoG["Cell ID"].apply(lambda x: ' '.join(sorted(set(x))))
# Get Cell Count per Site ID
CellList['Cell ID'] = CellList['Cid'].str.slice(start=6)
CellList['Band'] = CellList['Cell ID'].str[-1]
CellList['Site ID'] = CellList['Cid'].str.slice(start=0,stop=5)
CellList['Site ID'] = CellList['Site ID'].astype(str)+'_'+CellList['Band']
CellList['Cell ID'] = CellList['Cell ID'].str[:-1]
Cellgroup = CellList.groupby('Site ID')
Cell_List = Cellgroup['Cell ID'].apply(lambda x: ' '.join(sorted(x)))
CellCount = CellList.groupby('Site ID')['Cell ID'].size()
#Clear Site ID From Site List
TTList.rename(columns = {'faultFirstOccurTime(CF_TT_Create (Create TT))':'FFOT','ResolvedFault Recovery Time(CF_TT_Process (Process TT))':'FRT'}, inplace = True)

TTList['Site ID'] = TTList['ProblemDescription(CF_TT_Create (Create TT))'].apply(lambda x: split_descriptions(str(x)))
TTList['Site ID'] = TTList['Site ID'].apply(lambda x: ', '.join(x) if x else None)
TTList['Site ID'] = TTList['Site ID'].str.split(', ')
# Create a new DataFrame with 'faultFirstOccurTime' and 'ResolvedFault Recovery Time'
time_df = TTList[['FFOT', 'FRT']]
# Explode the 'Sites' list into multiple rows and join with 'time_df'
expanded_df = TTList['Site ID'].explode().to_frame().join(time_df)
# Reset the index of the final DataFrame
TTList = expanded_df.reset_index(drop=True)
TTList['Site ID'] = TTList['Site ID'].str.replace(' ', '')
#Cross check Overlapped alarm

TwoG = TwoG.reset_index(drop=True)
#Alarmlist.drop_duplicates(subset=['NEevetntimeceasetime', 'Cell ID'], inplace=True)
#Alarmlist = Alarmlist.drop(columns=['CuststartDateendDate'])

TwoG['Duration'] = TwoG['cease time'] - TwoG['event time']
print('getting minutes and Condition for Filter')
minutes, before_or_after = get_user_input()
TwoG = filter_dataframe(TwoG.copy(), minutes, before_or_after)
print('Filtering DF')
TwoG['Down Cell Count'] = TwoG['Cell ID'].str.replace(' ', '').str.len()
print('Check Outage alarm with Sites Cell Count')
TwoG = pd.merge(TwoG, CellCount, on='Site ID', how='outer')
TwoG['Cell ID_y'].fillna(TwoG['Down Cell Count'], inplace=True)
TwoG['Band'] = TwoG['Site ID'].str[-1]
TwoG['TT Check'] = [
    'Need TT' if (x >= y/2 and band == '0') or (x >= y and band == '1') else 'No need TT'
    for x, y, band in zip(TwoG['Down Cell Count'], TwoG['Cell ID_y'], TwoG['Band'])
]
TwoG['Site ID'] = TwoG['Site ID'].str.slice(start=0,stop=5)
TwoG['Duration'] = TwoG["Duration"].apply(lambda x: f"{x.seconds // 3600}:{(x.seconds // 60) % 60}")
print('Check Alarms with TTs')
TwoG = pd.merge(TwoG, TTList, on='Site ID', how='outer')
TwoG['Differ'] = TwoG.apply(calculate_time_difference, axis=1)
TwoG['TT Status'] = TwoG.apply(check_time_conditions, axis=1)
TwoG = remove_based_on_tt_status(TwoG)
TwoG = TwoG.sort_values(by='TT Status', ascending=True)
TwoG = TwoG.drop_duplicates(subset=['NEevetntimeceasetime'],keep='first')
# Calculate the percentage
TwoG['Percentage'] = (TwoG['Down Cell Count'] / TwoG['Cell ID_y']) * 100
# Format the 'Percentage' column as a percentage with two decimal places
TwoG['Percentage'] = TwoG['Percentage'].apply(lambda x: f"{x:.2f}%")
TwoG['TT Status'] = TwoG['TT Status'].replace('HCheck TT', 'Check TT')
TwoG['TT Status'] = TwoG['TT Status'].replace('GFFOT and FRT need to check', 'FFOT and FRT need to check')
with pd.ExcelWriter('Resultss.xlsx') as writer:
    Alarmlist.to_excel(writer,'Alarms',index=False)
    TTList.to_excel(writer,'TT',index=False)
    CellList.to_excel(writer,'Cell',index=False)
    TwoG.to_excel(writer,'Result',index=False)