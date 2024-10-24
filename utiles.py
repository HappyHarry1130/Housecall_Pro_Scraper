import openai
from openai import OpenAIError  # Import the correct exception
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import traceback
import requests
import re
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()
api_key = os.getenv('OPENAI_API_KEY')
openai.api_key = api_key


def parse_time(time_str):
    return datetime.strptime(time_str, '%I:%M %p')

def extract_level(data):
    match = re.search(r'LV\d+', data)
    return match.group(0) if match else None

def write_to_google_sheet(data, spreadsheet_id, sheet_name):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)

        service = build('sheets', 'v4', credentials=creds)

        spreadsheet = client.open_by_key(spreadsheet_id)
        try:
            sheet = spreadsheet.worksheet(sheet_name)
            # Check if today is Monday and clear the sheet if it exists
            if datetime.now().weekday() == 0:  # 0 is Monday
                # sheet.clear()
                None
        except gspread.exceptions.WorksheetNotFound:
            # Create the sheet if it doesn't exist
            sheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")
        try:
            sheet = spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")

        level = extract_level(sheet_name)
        if level == 'lV1':
            level_value = 140
        elif level == 'LV2':
            level_value = 185
        elif level == 'LV3':
            level_value = 230
        else:
            level_value = 200
        current_headers = sheet.row_values(1)
        headers = ["Date", "Customer", "Link",  "Claim Approved or Denied",  "OMW",  "ClockIn", "ClockOut", "Drive Time", "Time Spent on Job", "What was done", "Tech Called Office (Yes or No)", "What Was Call About?","Misdiagnosis (Yes or No)","Job Pics", "Job Notes","Tech Submitted Estimate For Upsell Opportunity (Yes or No)", "Missed opportunities", "Maintenance Plan Offered?",  "Seg 1 Subtotal", "Seg 2 Subtotal" , "Discount", "Final Invoice", "Our Cost On Parts", "Tech Burden to run call", "Day rate per job Cost",  "Commision Paid To Tech","Daily Revenue  (+/-):", "Conclusions", "Techs Reasoning"]
        if current_headers:
            sheet.update('A1:AG1', [headers])
        else:
            sheet.update('A1:AG1', [headers])
        row_height_request = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "ROWS",
                            "startIndex": 0,  
                            "endIndex": 1
                        },
                        "properties": {
                            "pixelSize": 100  
                        },
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=row_height_request).execute()
        
        column_width_request = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,  
                            "endIndex": 35
                        },
                        "properties": {
                            "pixelSize": 200  
                        },
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=column_width_request).execute()
        column0_width_request = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,  
                            "endIndex": 1
                        },
                        "properties": {
                            "pixelSize": 100  
                        },
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=column0_width_request).execute()

        column0_width_request = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 1,  
                            "endIndex": 2
                        },
                        "properties": {
                            "pixelSize": 300  
                        },
                        "fields": "pixelSize"
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 2,  
                            "endIndex": 3
                        },
                        "properties": {
                            "pixelSize": 400  
                        },
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=column0_width_request).execute()

        column0_width_request = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 2,  
                            "endIndex": 4
                        },
                        "properties": {
                            "pixelSize": 250  
                        },
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=column0_width_request).execute()

        # Set the background color of the first row
        background_color_request = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": len(headers)
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 1.0,
                                    "green": 0.949,
                                    "blue": 0.8
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=background_color_request).execute()



        driven_time_colnum = 8
        time_spent_colnum= 9
        job_pics_colume = 14
        job_note_colume = 15
        # pay_colume = 21

        driven_time = data[driven_time_colnum-1]

        time_spent = data[time_spent_colnum-1]

                
        job_pics = data[job_pics_colume-1] 
        data[job_pics_colume-1] = "Yes" if job_pics > 0 else "No"
                
        job_notes_length = data[job_note_colume-1]  # I column is the 9th column, 0-indexed as 8
        data[job_note_colume-1] = "Yes" if job_notes_length > 0 else "No"
        
        # pay =data[pay_colume-1]
        # data [pay_colume-1] = "Yes" if pay == 10 else "No"


        existing_rows = sheet.get_all_values()
        today = datetime.now().date()
        row_to_update = None
        count_of_calls =0
        rows_to_delete = []
        for i, row in enumerate(existing_rows[1:], start=2):  # Skip header row, start index at 2 for Google Sheets
            try:
                row_date = datetime.strptime(row[0], '%Y-%m-%d').date()  # Adjust date format as needed
                if (today - row_date).days > 7:
                    rows_to_delete.append(i)
            except ValueError:
                print(f"Skipping row {i} due to invalid date format: {row[0]}")

        # Delete rows in reverse order to avoid index shifting issues
        for row_index in reversed(rows_to_delete):
            sheet.delete_rows(row_index)
        for i, row in enumerate(existing_rows):
            if row[0] == data[0] and row[1] == data[1]:  
                row_to_update = i + 1  
                break
        for i, row in enumerate(existing_rows):
            if row[0] == data[0]:  
                count_of_calls+=1 
                break
        if row_to_update:
            None
        else:
            count_of_calls+=1
        print (f'count of calls : {count_of_calls}')
        one_cost_of_tech = level_value / count_of_calls

        if data[4]:
            omw_time = parse_time(data[4])
        else: omw_time = 0
        if data[6]:
            clockout_time = parse_time(data[6])
        else: clockout_time = omw_time
        time_range = clockout_time - omw_time
        print(f'time_range: {time_range}')
        if time_range ==0: tech_Burden_to_run_call = 0
        else :tech_Burden_to_run_call = 0.8 * time_range.total_seconds()/60
        print(f'tech_Burden_to_run_call: {tech_Burden_to_run_call}')
        data[23] = tech_Burden_to_run_call

        if row_to_update:
            # Exclude "Claim Submitted(Yes or No)" column (index 5) from the update
            existing_row = sheet.row_values(row_to_update)
            data[3] = existing_row[3]
            data[9] = existing_row[9]
            data[10] = existing_row[10]
            data[11] = existing_row[11]
            data[12] = existing_row[12]
            data[15] = existing_row[15]
            data[16] = existing_row[16]
            data[17] = existing_row[17]
            data[18] = existing_row[18]
            data[19] = existing_row[19]
            data[20] = existing_row[20]
            data[21]= existing_row[21]
            data[22] = existing_row[22]
            data[24] = existing_row[24]
            data[25] = existing_row[25]
            data[26] = existing_row[26]
            data[27] = existing_row[27]
            data[28] = existing_row[28]
            
            # data_to_update = data[:4] + data[6:10] + data[15:16] + data[21:27]   # Ensure data[29] is a list
            cell_range = f'A{row_to_update}:AG{row_to_update}'
            sheet.update(cell_range, [data])
            next_row = row_to_update
        else:
            next_row = len(existing_rows) + 1
            sheet.append_row(data)

        if driven_time < 20:
            cell = sheet.cell(next_row, driven_time_colnum)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        elif driven_time<30:
            cell = sheet.cell(next_row, driven_time_colnum)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        else:
            cell = sheet.cell(next_row, driven_time_colnum)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 0,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)

        if time_spent < 45:
            cell = sheet.cell(next_row, time_spent_colnum)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        elif time_spent<100:
            cell = sheet.cell(next_row, time_spent_colnum)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        else:
            cell = sheet.cell(next_row, time_spent_colnum)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 0,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        

        if job_pics > 10:
            cell = sheet.cell(next_row, job_pics_colume)  
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        elif job_pics > 0:
            cell = sheet.cell(next_row, job_pics_colume)  
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        else:
            cell = sheet.cell(next_row, job_pics_colume)  
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 0,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)


        if job_notes_length < 25:
            cell = sheet.cell(next_row, job_note_colume)  
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 0,
                    "blue": 0
                }
            }
        elif job_notes_length < 40:
            cell = sheet.cell(next_row, job_note_colume)  
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 1,
                    "blue": 0
                }
            }
        else:
            cell = sheet.cell(next_row, job_note_colume)  
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
        sheet.format(cell.address, cell_format)
        

        # if pay == 1:
        #     cell = sheet.cell(next_row, pay_colume)  
        #     cell_format = {
        #         "backgroundColor": {
        #             "red": 1,
        #             "green": 0,
        #             "blue": 0
        #         }
        #     }
        #     sheet.format(cell.address, cell_format)
        # elif pay ==10:
        #     cell = sheet.cell(next_row, pay_colume)  
        #     cell_format = {
        #         "backgroundColor": {
        #             "red": 0,
        #             "green": 1,
        #             "blue": 0
        #         }
        #     }
        #     sheet.format(cell.address, cell_format)
        # else:
        #     cell = sheet.cell(next_row, pay_colume)  
        #     cell_format = {
        #         "backgroundColor": {
        #             "red": 1,
        #             "green": 1,
        #             "blue": 0
        #         }
        #     }
        #     sheet.format(cell.address, cell_format)
        
        # if data[29].strip().isdigit() and int(data[29]) > 0:
        #     cell = sheet.cell(next_row, 30)  
        #     cell_format = {
        #         "backgroundColor": {
        #             "red": 0,
        #             "green": 1,
        #             "blue": 0
        #         }
        #     }
        #     sheet.format(cell.address, cell_format)
        # else:
        #     cell = sheet.cell(next_row, 30) 
        #     cell_format = {
        #         "backgroundColor": {
        #             "red": 1,
        #             "green": 0,
        #             "blue": 0
        #         }
        #     }
        #     sheet.format(cell.address, cell_format)
        # Set up data validation for "Job Pics" and "Job Notes" columns in the newly added row
        service = build('sheets', 'v4', credentials=creds)
        body = {
            "requests": [

                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": next_row - 1,  # Adjust to the specific row (0-indexed)
                            "endRowIndex": next_row,        # Adjust to the specific row + 1 (0-indexed)
                            "startColumnIndex": job_pics_colume-1,          # Adjust to the specific column (0-indexed)
                            "endColumnIndex": job_pics_colume             # Adjust to the specific column + 1 (0-indexed)
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [
                                    {"userEnteredValue": "Yes"},
                                    {"userEnteredValue": "No"}
                                ]
                            },
                            "showCustomUi": True
                        }
                    }
                },
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": next_row - 1,  # Adjust to the specific row (0-indexed)
                            "endRowIndex": next_row,        # Adjust to the specific row + 1 (0-indexed)
                            "startColumnIndex": job_note_colume-1,          # Adjust to the specific column (0-indexed)
                            "endColumnIndex": job_note_colume             # Adjust to the specific column + 1 (0-indexed)
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [
                                    {"userEnteredValue": "Yes"},
                                    {"userEnteredValue": "No"}
                                ]
                            },
                            "showCustomUi": True
                        }
                    }
                },
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

        

        # Center-align all cells in the sheet
        alignment_request = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": 0,
                            "endRowIndex": next_row,
                            "startColumnIndex": 0,
                            "endColumnIndex": 35
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE"
                            }
                        },
                        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=alignment_request).execute()

        # Sort the sheet by Date column (A column)
        sort_request = {
            "requests": [
                {
                    "sortRange": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": 1,  # Skip the header row
                            "endRowIndex": next_row,  # Sort up to the last row
                            "startColumnIndex": 0,  # A column
                            "endColumnIndex": 35  # M column
                        },
                        "sortSpecs": [
                            {
                                "dimensionIndex": 0,  # Sort by the first column (Date)
                                "sortOrder": "ASCENDING"
                            }
                        ]
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=sort_request).execute()

        freeze_request = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet.id,
                            "gridProperties": {
                                "frozenRowCount": 1,
                                "frozenColumnCount": 3
                            }
                        },
                        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
                    }
                }
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=freeze_request).execute()

    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Error: Spreadsheet with ID '{spreadsheet_id}' not found.")
    except gspread.exceptions.APIError as api_error:
        print(f"APIError: {api_error}")
        if api_error.response:
            try:
                print(f"Response content: {api_error.response.content}")
                print(f"Response JSON: {api_error.response.json()}")
            except requests.exceptions.JSONDecodeError:
                print("Response is not in JSON format.")
    except Exception as e:
        print(f"An error occurred: {e}")
        traceback.print_exc()
   

def Warranty_Retail(question, max_retries=5):
    messages = [
        {"role": "system", "content": f"""You need to separate warranty and retail.
                    If it is warranty, you need to find the warranty company.
                    Warranty and retail can be one of this list: "Warranty, Retail, Recall, Warranty Repair, MCM"
                    Warranty company can be one of this list: "None, AHS, Choice, ORHP, Platinum"
                    Your output should be in json format {{"Warranty and retail":XXX,"Warranty company":XXX}}, nothing else needed such as explain.
                    You must print only json format. don't need ```json  .... ```.
                    """},
        {"role": "user", "content": f"""
            {question}
        """}
    ]

    for attempt in range(max_retries):
        try:
            completion = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",  # Correct model name
                messages=messages
            )
            return completion.choices[0].message['content']
        except OpenAIError as e:  # Use the correct exception
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                print("Retrying...")
            else:
                print("Max retries reached. Exiting.")
                raise

def get_informations(question, max_retries=5):
    messages = [
        {"role": "system", "content": f"""You are an information analyst.
                    You need to analyze the information and return these values ​​in the corresponding json format.
                    The information you need to get is:
                    Name, Address, Email, Phone Number
                     {{"Name":XXX,"Address":XXX, "Email":XXX, "Phone Number":XXX}}, nothing else needed such as explain.
                    You must print only json format. don't need ```json  .... ```.
                    """},
        {"role": "user", "content": f"""
            {question}
        """}
    ]

    for attempt in range(max_retries):
        try:
            completion = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",  # Correct model name
                messages=messages
            )
            return completion.choices[0].message['content']
        except OpenAIError as e:  # Use the correct exception
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                print("Retrying...")
            else:
                print("Max retries reached. Exiting.")
                raise

def write_to_google_sheet_2(data, spreadsheet_id, sheet_name="Data"):
    # Define the scope and authenticate
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('./credentials.json', scope)
    client = gspread.authorize(creds)

    # Open the spreadsheet
    spreadsheet = client.open_by_key(spreadsheet_id)

    if datetime.now().weekday() == 0:
        sheet = spreadsheet.worksheet(sheet_name)
        sheet.clear()
    # Check if the sheet already exists
    try:
        sheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")


    # Define headers
    headers = ["DATE", "Customer", "Link", "PART", "TOTAL SALE", "TOTAL"]
    current_headers = sheet.row_values(1)        
    if current_headers:
        None
    else:
        sheet.update('A1:F1', [headers])


    # Get all existing data in the sheet (after clearing, this will only be the headers)
    existing_rows = sheet.get_all_values()

    # Check for existing entry
    row_to_update = None
    for i, row in enumerate(existing_rows):
        if row[0] == data[0] and row[1] == data[1]:  # Compare date and customer
            row_to_update = i + 1  # 1-indexed for Google Sheets
            break

    if row_to_update:
        # Update the existing row
        cell_range = f'A{row_to_update}:F{row_to_update}'
        existing_row = sheet.row_values(row_to_update)
        data[3] = existing_row[3]
        data[4] = existing_row[4]
        data[5] = existing_row[5]
        sheet.update(cell_range, [data])
        next_row = row_to_update
    else:
        # Append new data
        sheet.append_row(data)
        next_row = len(existing_rows) + 1

    # Formatting for the header row
    header_range = f"A1:F1"  # Adjust based on your headers
    sheet.format(header_range, {
        "backgroundColor": {"red": 0, "green": 0.8, "blue": 0.8},
        "textFormat": {"bold": True},
        "horizontalAlignment": "CENTER"
    })
    header_range = f"A1:I1"  # Adjust based on your headers

    # Set column widths and row heights using Google Sheets API
    service = build('sheets', 'v4', credentials=creds)
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,  # Column A
                        "endIndex": 1,    # Column A
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 1,  # Column B
                        "endIndex": 2,
                    },
                    "properties": {
                        "pixelSize": 200
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 2,  # Column C
                        "endIndex": 3,
                    },
                    "properties": {
                        "pixelSize": 450
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 3,  # Column D
                        "endIndex": 4,
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 4,  # Column E
                        "endIndex": 5,
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 5,  # Column F
                        "endIndex": 6,
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 0,  # Row 1 (header)
                        "endIndex": 1,    # Row 1 (header)
                    },
                    "properties": {
                        "pixelSize": 70  # Set the height for the header row
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 1,  # Row 2 (first data row)
                        "endIndex": next_row,  # Up to the last row
                    },
                    "properties": {
                        "pixelSize": 50  # Set the height for the data rows
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "sortRange": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 1,  # Skip the header row
                        "endRowIndex": next_row,  # Sort up to the last row
                        "startColumnIndex": 0,  # A column
                        "endColumnIndex": 35  # M column
                    },
                    "sortSpecs": [
                        {
                            "dimensionIndex": 0,  # Sort by the first column (Date)
                            "sortOrder": "ASCENDING"
                        }
                    ]
                }
            }
        ]
    }

    # Execute the request to set column widths, row heights, and sort the data
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    alignment_request = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": next_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 35
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=alignment_request).execute()


def write_to_google_sheet_3(data, spreadsheet_id, sheet_name="Data"):
    # Define the scope and authenticate
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('./credentials.json', scope)
    client = gspread.authorize(creds)

    # Open the spreadsheet
    spreadsheet = client.open_by_key(spreadsheet_id)

    # Check if the sheet already exists
    try:
        sheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")


    # Define headers
    headers = ["Date", "Customer",  "Link" , "Adress", "Email", "Phone Number", "What have done", "Material", "Service"]
    current_headers = sheet.row_values(1)        
    if current_headers:
        None
    else:
        sheet.update('A1:I1', [headers])


    # Get all existing data in the sheet (after clearing, this will only be the headers)
    existing_rows = sheet.get_all_values()

    # Check for existing entry
    row_to_update = None
    for i, row in enumerate(existing_rows):
        if row[2] == data[2]:
            row_to_update = i + 1  # 1-indexed for Google Sheets
            break

    if row_to_update:
        # Update the existing row
        cell_range = f'A{row_to_update}:I{row_to_update}'
        sheet.update(cell_range, [data])
        next_row = row_to_update
    else:
        # Append new data
        sheet.append_row(data)
        next_row = len(existing_rows) + 1

    # Formatting for the header row
    header_range = f"A1:I1"  # Adjust based on your headers
    sheet.format(header_range, {
        "backgroundColor": {"red": 0, "green": 0.8, "blue": 0.8},
        "textFormat": {"bold": True},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE" 
    })

    # Set column widths and row heights using Google Sheets API
    service = build('sheets', 'v4', credentials=creds)
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,  # Column A
                        "endIndex": 1,    # Column A
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 1,  # Column A
                        "endIndex": 2,    # Column A
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 2,  # Column A
                        "endIndex": 3,    # Column A
                    },
                    "properties": {
                        "pixelSize": 350
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 3,  # Column B
                        "endIndex": 4,
                    },
                    "properties": {
                        "pixelSize": 350
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 4,  # Column C
                        "endIndex": 5,
                    },
                    "properties": {
                        "pixelSize": 240
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 5,  # Column D
                        "endIndex": 6,
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 6,  # Column E
                        "endIndex": 7,
                    },
                    "properties": {
                        "pixelSize": 400
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 7,  # Column F
                        "endIndex": 8,
                    },
                    "properties": {
                        "pixelSize": 400
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 8,  # Column F
                        "endIndex": 9,
                    },
                    "properties": {
                        "pixelSize": 400
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 0,  # Row 1 (header)
                        "endIndex": 1,    # Row 1 (header)
                    },
                    "properties": {
                        "pixelSize": 70  # Set the height for the header row
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 1,  # Row 2 (first data row)
                        "endIndex": next_row,  # Up to the last row
                    },
                    "properties": {
                        "pixelSize": 30  # Set the height for the data rows
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "sortRange": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 1,  # Skip the header row
                        "endRowIndex": next_row,  # Sort up to the last row
                        "startColumnIndex": 0,  # A column
                        "endColumnIndex": 7  # M column
                    },
                    "sortSpecs": [
                        {
                            "dimensionIndex": 0,  # Sort by the first column (Date)
                            "sortOrder": "ASCENDING"
                        }
                    ]
                },
                
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": next_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 7
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"  # Add this line for vertical centering
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": next_row,  # Adjust to the number of rows you have
                        "startColumnIndex": 6,  # Column E
                        "endColumnIndex": 9
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "wrapStrategy": "WRAP"  # Enable text wrapping
                        }
                    },
                    "fields": "userEnteredFormat.wrapStrategy"
                }
            },

            
        ]
    }
    freeze_request = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet.id,
                        "gridProperties": {
                            "frozenRowCount": 1,
                            "frozenColumnCount": 3
                        }
                    },
                    "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=freeze_request).execute()

    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def write_to_google_sheet_for_recall(data, spreadsheet_id, sheet_name="Data"):

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('./credentials.json', scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(spreadsheet_id)

    try:
        sheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")

    headers = ["Date", "Customer",  "Link" , "Tags"]
    current_headers = sheet.row_values(1)        
    if current_headers:
        None
    else:
        sheet.update('A1:D1', [headers])

    existing_rows = sheet.get_all_values()
    row_to_update = None
    for i, row in enumerate(existing_rows):
        if row[2] == data[2]:
            row_to_update = i + 1  
            break

    if row_to_update:
        cell_range = f'A{row_to_update}:I{row_to_update}'
        sheet.update(cell_range, [data])
        next_row = row_to_update
    else:
        sheet.append_row(data)
        next_row = len(existing_rows) + 1

    header_range = f"A1:D1" 
    sheet.format(header_range, {
        "backgroundColor": {"red": 0, "green": 0.8, "blue": 0.8},
        "textFormat": {"bold": True},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE" 
    })


    service = build('sheets', 'v4', credentials=creds)
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 0, 
                        "endIndex": 1,  
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 1,  
                        "endIndex": 2,    
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 2,  
                        "endIndex": 3,    
                    },
                    "properties": {
                        "pixelSize": 350
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 3, 
                        "endIndex": 4,
                    },
                    "properties": {
                        "pixelSize": 300
                    },
                    "fields": "pixelSize"
                }
            },
            
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 0,  
                        "endIndex": 1,   
                    },
                    "properties": {
                        "pixelSize": 70  
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 1, 
                        "endIndex": next_row,  
                    },
                    "properties": {
                        "pixelSize": 30  
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "sortRange": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 1,  
                        "endRowIndex": next_row,  
                        "startColumnIndex": 0,  
                        "endColumnIndex": 7  
                    },
                    "sortSpecs": [
                        {
                            "dimensionIndex": 0,  
                            "sortOrder": "ASCENDING"
                        }
                    ]
                },
                
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": next_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 7
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"  
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": next_row,  
                        "startColumnIndex": 6,  
                        "endColumnIndex": 9
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "wrapStrategy": "WRAP"  
                        }
                    },
                    "fields": "userEnteredFormat.wrapStrategy"
                }
            },

            
        ]
    }
    freeze_request = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet.id,
                        "gridProperties": {
                            "frozenRowCount": 1,
                            "frozenColumnCount": 3
                        }
                    },
                    "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=freeze_request).execute()
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def write_to_google_sheet_for_warrancy(data, spreadsheet_id, sheet_name="Data"):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('./credentials.json', scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)

    try:
        sheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")

    headers = ["Date", "Customer",  "Link" , "Adress", "Email", "Phone Number"]
    current_headers = sheet.row_values(1)        
    if current_headers:
        None
    else:
        sheet.update('A1:F1', [headers])



    existing_rows = sheet.get_all_values()


    row_to_update = None
    for i, row in enumerate(existing_rows):
        if row[2] == data[2]:
            row_to_update = i + 1  
            break

    if row_to_update:

        cell_range = f'A{row_to_update}:F{row_to_update}'
        sheet.update(cell_range, [data])
        next_row = row_to_update
    else:
        sheet.append_row(data)
        next_row = len(existing_rows) + 1

    header_range = f"A1:F1" 
    sheet.format(header_range, {
        "backgroundColor": {"red": 0, "green": 0.8, "blue": 0.8},
        "textFormat": {"bold": True},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE" 
    })


    service = build('sheets', 'v4', credentials=creds)
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 1,   
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 1, 
                        "endIndex": 2,    
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 2, 
                        "endIndex": 3,   
                    },
                    "properties": {
                        "pixelSize": 350
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 3, 
                        "endIndex": 4,
                    },
                    "properties": {
                        "pixelSize": 350
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 4,  
                        "endIndex": 5,
                    },
                    "properties": {
                        "pixelSize": 240
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 5,  
                        "endIndex": 6,
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 6,  
                        "endIndex": 7,
                    },
                    "properties": {
                        "pixelSize": 400
                    },
                    "fields": "pixelSize"
                }
            },


            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 0,  
                        "endIndex": 1,    
                    },
                    "properties": {
                        "pixelSize": 70  
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": 1,  
                        "endIndex": next_row,  
                    },
                    "properties": {
                        "pixelSize": 30  
                    },
                    "fields": "pixelSize"
                }
            },
            {
                "sortRange": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 1, 
                        "endRowIndex": next_row,  
                        "startColumnIndex": 0,  
                        "endColumnIndex": 7 
                    },
                    "sortSpecs": [
                        {
                            "dimensionIndex": 0,  
                            "sortOrder": "ASCENDING"
                        }
                    ]
                },
                
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": next_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 7
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE" 
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": next_row,  
                        "startColumnIndex": 6,  
                        "endColumnIndex": 9
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "wrapStrategy": "WRAP"  
                        }
                    },
                    "fields": "userEnteredFormat.wrapStrategy"
                }
            },

            
        ]
    }
    freeze_request = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet.id,
                        "gridProperties": {
                            "frozenRowCount": 1,
                            "frozenColumnCount": 3
                        }
                    },
                    "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=freeze_request).execute()

    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()