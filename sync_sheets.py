import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yaml


with open("settings.yaml") as stream:
    try:
        formatted_params = yaml.safe_load(stream)
        print()
    except yaml.YAMLError as exc:
        print(exc)

source_sheet_url = formatted_params["sheet_settings"]["source_sheet_url"]
target_sheet_url = formatted_params["sheet_settings"]["target_sheet_url"]
source_worksheet_name = formatted_params["sheet_settings"]["source_worksheet_name"]
target_worksheet_name = formatted_params["sheet_settings"]["target_worksheet_name"]

print("Loaded settings.")

def sync_sheets(source_sheet_url, target_sheet_url, source_worksheet_name, target_worksheet_name):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('sac_creds/credentials.json', scope)
        client = gspread.authorize(creds)
        print("Credentials loaded successfully.")

        source_sheet = client.open_by_url(source_sheet_url).worksheet(source_worksheet_name)
        print(f"Opened source sheet: {source_sheet_url}")
        target_sheet = client.open_by_url(target_sheet_url).worksheet(target_worksheet_name)
        print(f"Opened target sheet: {target_sheet_url}")

        source_data = source_sheet.get_all_values()
        print(f"Source sheet dimensions: {len(source_data)} rows, {len(source_data[0]) if source_data else 0} columns")

        target_data = target_sheet.get_all_values()
        print(f"Target sheet dimensions: {len(target_data)} rows, {len(target_data[0]) if target_data else 0} columns")

        updates = []
        for row_index, row in enumerate(source_data, start=1):
            for col_index, value in enumerate(row, start=1):
                if (row_index > len(target_data) or 
                    col_index > len(target_data[row_index-1]) or 
                    value != target_data[row_index-1][col_index-1]):
                    updates.append({
                        'range': gspread.utils.rowcol_to_a1(row_index, col_index),
                        'values': [[value]]
                    })

        if updates:
            target_sheet.batch_update(updates)
            print(f"Updated {len(updates)} cells in the target sheet.")
        else:
            print("No updates were necessary. The sheets are identical.")

        if len(source_data) > len(target_data):
            target_sheet.add_rows(len(source_data) - len(target_data))
            print(f"Added {len(source_data) - len(target_data)} rows to the target sheet.")
        
        if len(source_data[0]) > len(target_data[0]):
            target_sheet.add_cols(len(source_data[0]) - len(target_data[0]))
            print(f"Added {len(source_data[0]) - len(target_data[0])} columns to the target sheet.")

    except Exception as e:
        print(f"An error occurred: {str(e)}")


sync_sheets(source_sheet_url, target_sheet_url, source_worksheet_name, target_worksheet_name)