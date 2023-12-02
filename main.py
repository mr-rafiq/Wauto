import time
import pywhatkit as kit
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import os
import schedule
from dotenv import load_dotenv
# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
load_dotenv()
path = os.getenv("APIPath")
# Set up Google Sheets API credentials
creds = ServiceAccountCredentials.from_json_keyfile_name(path, SCOPES)
client = gspread.authorize(creds)

# Load existing news data from Google Sheets
sheet_name = "WhatsAPP"

#Initiate the sheet in Pandas
mess_df = pd.DataFrame(columns=['Title', 'Message', 'Link', 'Timestamp'])


#Reload the sheet
def reload_sheet():
    try:
        sheet = client.open(sheet_name)
        contacts_sheet = sheet.worksheet("Contacts")
        content_sheet = sheet.worksheet("Content")
        contacts_sheet_values = contacts_sheet.get_all_values()
        content_sheet_vaues = content_sheet.get_all_values()
        
        if contacts_sheet_values and len(contacts_sheet_values) > 1 and content_sheet_vaues and len(content_sheet_vaues) > 1:
            contacts_header = contacts_sheet_values[0]
            contacts_records = contacts_sheet_values[1:]
            contacts_sheet_df = pd.DataFrame(contacts_records, columns=contacts_header)
            
            content_header = content_sheet_vaues[0]
            content_records = content_sheet_vaues[1:]
            content_sheet_df = pd.DataFrame(content_records, columns=content_header)
            print(f"Opened existing sheet: {sheet_name}")
            
            return content_sheet_df,contacts_sheet_df
        else:
            print(f"Sheet {sheet_name} exists but has no data.")
    except gspread.exceptions.SpreadsheetNotFound:
        sheet = client.create(sheet_name)
        sheet.add_worksheet(title="Content", rows=100, cols=20)
        sheet.add_worksheet(title="Contacts", rows=100, cols=20)
        content_sheet = sheet.worksheet("Content")
        contacts_sheet = sheet.worksheet("Contacts")
        content_sheet.insert_row(['Title', 'Message', 'Link', 'Timestamp'], 1)  # Add the header row
        contacts_sheet.insert_row(['Name', 'PhoneNr'], 1)  # Add the header row
        sheet.share('rafiqrooney17@gmail.com', perm_type='user', role='writer')
        content_sheet_df = pd.DataFrame(columns=['Title', 'Message', 'Link',  'Timestamp'])
        contacts_sheet_df = pd.DataFrame(columns=['Name', 'PhoneNr'])
        print(f"Created new sheet: {sheet_name}")
        return content_sheet_df, contacts_sheet_df



def send_message(phone_nr, message):
    kit.sendwhatmsg_instantly(phone_nr, message, 10, True, 10)
    
def send_instanly(phone_nr, message):
    kit.sendwhatmsg_instantly(phone_nr, message, 10, True, 10)


def main():
    content_sheet_df, contacts_sheet_df = reload_sheet()
    if len(content_sheet_df) == 0 or len(contacts_sheet_df) == 0:
        print("No data in sheet")
        return
    
    #if it has image in the link then send image
    if content_sheet_df.iloc[0]['Link'] is not None:
        print("Sending Image")
        for index, row in contacts_sheet_df.iterrows():
            formatted_number = f"+{row['PhoneNr']}"
            print(row['Name'], formatted_number)
            print(content_sheet_df.iloc[0]['Link'])
            print(content_sheet_df.iloc[0]['Message'])
            kit.sendwhats_image(formatted_number, content_sheet_df.iloc[0]['Link'], content_sheet_df.iloc[0]['Message'], 10, True, 10)
        return
    
    message = content_sheet_df.iloc[0]['Message']
    
    for index, row in contacts_sheet_df.iterrows():
        formatted_number = f"+{row['PhoneNr']}"
        print(row['Name'], formatted_number)
        send_message(formatted_number,message)

if __name__ == "__main__":
    #schedule every day at 9:00
    schedule.every().day.at("09:00").do(main)
    while True:
        schedule.run_pending()
        time.sleep(1)