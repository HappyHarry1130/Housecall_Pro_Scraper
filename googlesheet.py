import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import traceback
import requests

def write_to_google_sheet(data, spreadsheet_id='1Wah4JVOkaiGRvsYO0QOpcTkrg9i59NrGr5LA5YD8CrU', sheet_name='43'):
    try:
        # Set up the credentials and client
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)

        # Open the spreadsheet and select the sheet
        spreadsheet = client.open_by_key(spreadsheet_id)
        try:
            sheet = spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            # Create a new sheet if it does not exist
            sheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")

        # Write headers to the first row
        headers = ["Date", "Customer", "WARRANTY CALL OR RETAIL CALL", "WARRANTY COMPANY",  "OMW",  "ClockIn", "ClockOut", "Drive Time", "Time Spent on Job", "Job Pics", "Job Notes", "DID THE TECH COLLECT PAYMENT", "JOB COMPLEXITY"]
        sheet.update('A1:M1', [headers])

        driven_time = data[7]
        # Determine "Job Pics" value
        job_pics = data[9]  # H column is the 8th column, 0-indexed as 7
        data[9] = "Yes" if job_pics > 0 else "No"

        # Determine "Job Notes" value
        job_notes_length = data[10]  # I column is the 9th column, 0-indexed as 8
        data[10] = "Yes" if job_notes_length > 0 else "No"
        
        pay =data[11]
        data [11] = "Yes" if pay > 1 else "No"
        # Check if the row with the same Date and Customer exists
        existing_rows = sheet.get_all_values()
        row_to_update = None
        for i, row in enumerate(existing_rows):
            if row[0] == data[0] and row[1] == data[1]:  # Compare Date and Customer
                row_to_update = i + 1  # Google Sheets is 1-indexed
                break

        if row_to_update:
            # Update the existing row
            cell_range = f'A{row_to_update}:M{row_to_update}'
            sheet.update(cell_range, [data])
            next_row = row_to_update
        else:
            # Write data at the next available row
            next_row = len(existing_rows) + 1
            sheet.append_row(data)

        # Apply conditional formatting for job pics
        if job_pics < 10:
            cell = sheet.cell(next_row, 10)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 0,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        
        if pay < 10:
            cell = sheet.cell(next_row, 12)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 0,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        else:
            cell = sheet.cell(next_row, 12)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)

        if driven_time < 20:
            cell = sheet.cell(next_row, 8)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        elif driven_time<45:
            cell = sheet.cell(next_row, 8)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)
        else:
            cell = sheet.cell(next_row, 8)  # J column is the 10th column, 1-indexed as 10
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 1,
                    "blue": 0
                }
            }
            sheet.format(cell.address, cell_format)

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
                            "startColumnIndex": 9,          # Adjust to the specific column (0-indexed)
                            "endColumnIndex": 10             # Adjust to the specific column + 1 (0-indexed)
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
                            "startColumnIndex": 10,          # Adjust to the specific column (0-indexed)
                            "endColumnIndex": 11             # Adjust to the specific column + 1 (0-indexed)
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
                            "startColumnIndex": 11,          # Adjust to the specific column (0-indexed)
                            "endColumnIndex": 12             # Adjust to the specific column + 1 (0-indexed)
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

        # Apply conditional formatting for job notes based on length
        if job_notes_length < 100:
            cell = sheet.cell(next_row, 11)  # K column is the 11th column, 1-indexed as 11
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 0,
                    "blue": 0
                }
            }
        elif job_notes_length < 200:
            cell = sheet.cell(next_row, 11)  # K column is the 11th column, 1-indexed as 11
            cell_format = {
                "backgroundColor": {
                    "red": 1,
                    "green": 1,
                    "blue": 0
                }
            }
        else:
            cell = sheet.cell(next_row, 11)  # K column is the 11th column, 1-indexed as 11
            cell_format = {
                "backgroundColor": {
                    "red": 0,
                    "green": 1,
                    "blue": 0
                }
            }
        sheet.format(cell.address, cell_format)

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
                            "endColumnIndex": 13
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
                            "endColumnIndex": 13  # M column
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

# Example usage
data = ["2023-10-12", "Customer2", "", "",  "OMW1",  "8:00 AM", "9:00 AM", 30, 60, 5, 90, 10]  # 8th value is job_pics count, 9th value is job_notes length
spreadsheet_id = '1Wah4JVOkaiGRvsYO0QOpcTkrg9i59NrGr5LA5YD8CrU'
write_to_google_sheet(data, spreadsheet_id, 'JordanP')