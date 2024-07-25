import sys 
import pandas as pd
from openpyxl import load_workbook, Workbook
import math
import datetime
from datetime import timedelta
import re
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File

from io import BytesIO

#Need access to sharepoint and the correct document_library
def sharepoint():
    sharepoint_site = 'https://navistarinc.sharepoint.com/sites/BuildOperations'
    document_library = 'ace3f3d7-ed91-49d0-808f-16f6d50694bb'

    # Read credentials from file
    file = 'cred.txt'
    with open(file, 'r') as reader:
        uname = reader.readline().strip()
        pwd = reader.readline().strip()

    credentials = UserCredential(uname, pwd)
    ctx = ClientContext(sharepoint_site).with_credentials(credentials)
    
    # Correct the document URL
    file_url = f"{document_library}"
    
    # Load the Excel file from SharePoint
    target_file = ctx.web.get_file_by_server_relative_url(file_url)
    
    # Download the file content
    with open("Build Ops Master Schedule.xlsm", "wb") as local_file:
        target_file.download(local_file).execute_query()

    # Load the content into an openpyxl workbook
    with open("Build Ops Master Schedule.xlsm", "rb") as file_content:
        bo_workbook = load_workbook(file_content)
    
    bo_sheet = bo_workbook['Vehicle Master Build Schedule']
    return bo_sheet
'''
def sharepoint():

    sharepoint_site = 'https://navistarinc.sharepoint.com/:x:/r/sites/BuildOperations/_layouts/15/Doc.aspx?sourcedoc=%7B89CA8623-8556-4938-A0E5-F3D085E92F3D%7D&file=Build%20Ops%20Master%20Schedule.xlsm&action=default&mobileredirect=true&DefaultItemOpen=1'
    document_library = 'ace3f3d7-ed91-49d0-808f-16f6d50694bb'
    #'/sites/BuildOperations/Build Engineering/Master Schedule'
    #/sites/BuildOperations/Build%20Engineers/Master%20Schedule/Build%20Ops%20Master%20Schedule.xlsm?d=w89ca862385564938a0e5f3d085e92f3d&csf=1&web=1&e=9y9gJM
    file = 'cred.txt'
    with open(file, 'r') as reader:
            uname = reader.readline(1)
            pwd = reader.readline(2)
            reader.close()
    credentials = UserCredential(uname, pwd)
    ctx = ClientContext(sharepoint_site).with_credentials(credentials)
    

    # Load the Excel file from SharePoint
    response = ctx.web.get_file_by_server_relative_url(document_library).download()
    response.execute_query()       

    # Load the content into an openpyxl workbook
    excel_content = BytesIO(response.content)
    bo_workbook = load_workbook(excel_content)
    bo_sheet = bo_workbook['Vehicle Master Build Schedule']
    return bo_sheet
'''
#bo_sheet = sharepoint()

source_file = "Copy of Build Ops Master Schedule.xlsm"
target_file = "EES Test Work Load Projections Master -(Cooper).xlsx"


#Build Ops - read only data
bo_workbook = load_workbook(filename=source_file, read_only=True)
bo_sheet = bo_workbook['Vehicle Master Build Schedule'] 
headers = [cell.value for cell in bo_sheet[9]]



# Logic to get column names and indices
def get_column_index(column_name):
    if column_name in headers:
        return headers.index(column_name)
    else:
        raise ValueError(f"Column '{column_name}' not found in headers")

column_names = ['PRGM', 'WRTS #', 'Truck ID', 'STATUS', 'BUILD SITE', 'Shake Down Duration (working days)', 'BUILD END/EES START      PLANNED', 'BUILD END/EES START        ACTUAL'] 

'''
PRGM : 0
WRTS #: 1
Truck ID: 2
STATUS: 3
BUILD_SITE: 4
Shake down duration: 5
ees start planned: 6
actual: 7
'''
column_indices = [get_column_index(name) for name in column_names]

def gifc(column): #get index from column
    i = 0
    while i < len(column_names):
       # print(i)
        if column == column_names[i]:
            break
        i += 1
    return i



def normalize_vehicle_id(vehicle_id):
    # Use regex to separate the alphabetic part from the numeric part
    if 'LG1' in vehicle_id:
        return vehicle_id.replace("LG1", "")   # Take care of all weird formats 
    if 'TBD' in vehicle_id:
        return 0
    if 'CV' not in vehicle_id and 'DV' not in vehicle_id and 'CERT' not in vehicle_id: 
        return 0
    match = re.match(r"([A-Za-z]+)(\d+)", vehicle_id)
    if match:
        # Extract alphabetic and numeric parts
        alphabetic_part = match.group(1)
        numeric_part = match.group(2).lstrip('0')  # Remove leading zeros from the numeric part
        return f"{alphabetic_part}{numeric_part}"
    else:
        # If the input doesn't match the pattern, return it as is
        return vehicle_id    

def add_projections(p, program_list, v_len, row_place, i):
    #program_list is a list of tuples, must iterate accordingly
    
    #Create list of only vehicle info from program_list
    if len(program_list) < 1: #No vehicles to add 
        return
    
    vehicle_info = [normalize_vehicle_id(item[2]) for item in program_list if item[2] != None and item[2] != 'TBD']
    
    #Convert the DV notation
    

    handoff_predict = [item[6] for item in program_list] #This is type str list
    projection_headers = [cell.value for cell in projection_sheet[2]] #Dates throughout the year
    predict_date = format_dates(p, handoff_predict) #This is type datetime list
    #print("Date predict: ", predict_date)
    
    
    
    
    row_place -= 1
    
    #Logic for checking cols and rows 
    formated_cols = [normalize_vehicle_id(item) if type(item) == str else 0 for item in col_2]
    print("New formatted colums : ", formated_cols)

    if len(vehicle_info) < 1:
        return
    print("Vehicle Info: ", vehicle_info)
    #The outer for loop traverses the build ops vehicles
    #This inner loop traverses through the projection excel sheet
    reset_row_idx = row_place
    for dv, v in enumerate(vehicle_info):
        row_place = reset_row_idx #Once the vehicle is found, need to start over the local idx
        while row_place < prog_start_row[i+1]: #While in program section
            if dv >= len(vehicle_info): #Idx out of range check
                return
            if v == 0: #If vehicle is in 'TBD' status or otherwise from prev checks
                break
            if formated_cols[row_place] == 0: #If it is weird format (aka 0 - see normalize_vehicle_id()) then go to next row and vehicle
                row_place += 1
                
            elif v in formated_cols[row_place]:
                
                dat = predict_date[dv] #Need current vehicle ID
                add_to_sheet(row_place, dat, projection_headers)
                print(f" {p} -> r : {v}", row_place+1)
                
                row_place += 1
            else:
                row_place += 1
    return 
def add_to_sheet(r, target_date, ph): #Given the day started, populate the correct spots on excel sheet
   
    column_index = None
    for index, header in enumerate(ph):
        #print("Header : ", header, "\n")
        #print("target_date : ", target_date, "\n")
        if isinstance(header, datetime.datetime):
            start_of_week = header 
            end_of_week = header + datetime.timedelta(days = 6)
            if start_of_week <= target_date <= end_of_week:
                column_index = index + 1
                r += 1
                projection_sheet.cell(row=r, column=column_index, value=1)
                projection_sheet.cell(row=r, column=column_index+1, value=1)
                
                print("Added to cells")
                break

    




def iterate_program(prg, build_site):

    
    program_list = []
    print(f"Adding {prg} ...")

    for row in bo_sheet.iter_rows(min_row=11, values_only=True):  # Start from the 11th row (skip header)
        skip = True #Flag set to nullify Canceled and Complete statuses 
        if skip:
            selected_columns = list(row[idx] for idx in column_indices)
            
            if selected_columns[gifc('Shake Down Duration (working days)')] > 0: #Ensure it is a shakedown (Avoids date problems)

                
                if  isinstance(selected_columns[gifc('BUILD END/EES START      PLANNED')], datetime.date):
                    #selected_columns[gifc('BUILD END/EES START      PLANNED')] = selected_columns[gifc('BUILD END/EES START      PLANNED')].strftime("%m/%d/%Y")
                    if isInPast(selected_columns[gifc('BUILD END/EES START      PLANNED')]):
                        skip = False
                    else:
                        selected_columns[gifc('BUILD END/EES START      PLANNED')] = selected_columns[gifc('BUILD END/EES START      PLANNED')].strftime("%m/%d/%Y")

                if  isinstance(selected_columns[gifc('BUILD END/EES START        ACTUAL')], datetime.date):
                    selected_columns[gifc('BUILD END/EES START        ACTUAL')] = selected_columns[gifc('BUILD END/EES START        ACTUAL')].strftime("%m/%d/%Y") 
                if selected_columns[gifc('STATUS')] == 'Complete' or selected_columns[gifc('STATUS')] == 'Canceled':
                    skip = False
                if selected_columns[gifc('PRGM')] == prg and selected_columns[gifc('BUILD SITE')] == build_site and skip == True and (selected_columns[gifc('BUILD END/EES START      PLANNED')] or selected_columns[gifc('BUILD END/EES START        ACTUAL')]  != None): #Only current prg, ATC builds, 
                    wrts_info = tuple(selected_columns)
                    program_list.append(wrts_info)

    print("Success")
    #print("program list ", program_list)
    return program_list

def isInPast(t): #If date is before July 1, don't add 
    if t < datetime.datetime(2024, 7, 1, 0, 0) or t.year > 2024:
        return True
    else:
        return False
def format_dates(prg, handoff_predict):
    
    '''
    Format data from iterate_program to :

    -change to datetime 
    
    '''
    print(f"Formatting dates for {prg} ...")
    
    format_data = "%m/%d/%Y"
    predict_date = []
    actual_date = []
    
    
    for t in handoff_predict: #Convert to format -> datetime
        if t == 'N/A':
            continue
        else:
            predict_date.append(datetime.datetime.strptime(t, format_data))


    return predict_date

    ''' This will be for later, for now we will only do handoff_predict

    
    for i, t in enumerate(handoff_actual):
        if t == None:
            return i
        else:
            actual_date.append(datetime.datetime.strptime(t, format_data))
    '''

def read_target(): #This function will grow or shrink as new vehicles come in 

    i = 0
    for ranges in projection_sheet.iter_cols(min_row=1, max_row=124, min_col=1, max_col=2, values_only=True):
        print("")
        #normalize_vehicle_id(ranges[i])
    
    return list(ranges)

if __name__ == '__main__':

    
    projection_workbook = load_workbook(filename=target_file)
    if len(sys.argv) > 1:
        if sys.argv[1] == 'N':
            build_site = sys.argv[1]
            projection_sheet = projection_workbook['2024 NPG Bay Space(Shakedown)']
        elif sys.argv[1] == 'A':
            build_site = sys.argv[1]
            projection_sheet = projection_workbook['2024 ATC Bay Space(Shakedown)']
    else: #default to ATC if no arguments given
        build_site = 'A'
        projection_sheet = projection_workbook['2024 ATC Bay Space(Shakedown)']
      
    #Projections - read and write data
    
    projection_sheet = projection_workbook['2024 ATC Bay Space(Shakedown)']
    
    all_progs = []
    vehicle_len=[18,8,25,2,18,34,14,3]
    prog_start_row = [3, 21, 29, 54, 56, 74, 108, 122, 124]
    '''
    e82_length = 18
    h06_length = 8
    h08_length = 25
    j01_length = 2
    j07_length = 17
    j08_length = 34
    j09_length = 14
    s26_length = 3
    '''
    ranges = read_target()
   
    col_2 = [0 if x == None or x =='TBD' else x for x in ranges]
        
    for i, p in enumerate(all_progs):
        program_list = iterate_program(p, build_site) #Create new sheet for different build sites
        add_projections(p, program_list, vehicle_len[i], prog_start_row[i], i)
        

    projection_workbook.save(target_file)    

    