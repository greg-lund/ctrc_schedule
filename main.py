from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseDownload
from datetime import datetime,timedelta
import os.path
import io
import pickle
from docx import Document
import re
import pytz

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']
SCOPES += ["https://www.googleapis.com/auth/drive"]
FOLDER = "CTRC Schedules"

def open_services():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    calendar_service = build('calendar', 'v3', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return calendar_service,drive_service

def add_calendar(service,calendarName):
    # Check for calendarName calendar and create it if it doesn't exist
    calendar_list = service.calendarList().list().execute().get('items',[])
    for c in calendar_list:
        if c['summary'] == calendarName:
            print(calendarName,"already exists")
            return c
    else:
        new_calendar = {
                'summary': calendarName,
                'timeZone': 'America/Denver'
                }
        created_calendar = service.calendars().insert(body=new_calendar).execute()
        print(f"Created calendar: {created_calendar['id']}")
        return created_calendar

def scrape_docx(doc,docName="DL"):
    # Events array
    events = []

    for table in doc.tables:

        # Get date from top of table
        date_arr = []
        for cell in table.rows[0].cells:
            if len(date_arr) == 0:
                date_arr.append(cell.text)
                continue
            if cell.text == date_arr[-1]:
                continue
            date_arr.append(cell.text)

        # We only care about first table per page
        if(date_arr[0] == "Date:"):

            # Remove unnnecesary fields
            date_arr.pop(0)
            if date_arr[-1] == '':
                date_arr.pop()

            # Find Docs initials in column
            try:
                _,doc_name = table.rows[1].cells[-1].text.split("\n")
            except:
                continue

            if re.search(docName,doc_name,re.IGNORECASE) is not None:

                i = 2
                count = 0
                while i < len(table.rows):
                    row = table.rows[i]

                    # Get time
                    time = row.cells[0].text

                    # Look at last column for patient
                    pat_name = row.cells[-1].text
                    if re.search("[a-zA-Z][a-zA-Z]",pat_name) is not None:
                        date = " ".join(date_arr[1:-2])
                        duration = 15
                        while (i+1) < len(table.rows) and table.rows[i+1].cells[-1].text == pat_name:
                            i = i + 1
                            duration = duration + 15
                        events.append(create_event(date_arr,time,duration,pat_name))
                        count = count + 1
                    i = i + 1
                if count == 0:
                    events.append(create_available_event(date_arr))

    return events

def correct_hrs(hrs):
    if int(hrs) <= 5:
        hrs = str(int(hrs) + 12)
    elif int(hrs) < 10:
        hrs = '0' + hrs
    return hrs

def make_datetime(date_arr,time):
    """ 
    Create a datetime object from the following objects:
     date_arr will be in the following format: ["Monday", "January", "10", "2024"]
     time will be in the following format: "1230"
    """
    _, month_str, day_str, year_str = date_arr
    month = datetime.strptime(month_str, "%B").month  # Convert month name to month number
    day = int(day_str)
    year = int(year_str)

    mins = int(time[-2:])
    hrs = int(correct_hrs(time[:-2]))

    return datetime(year,month,day,hrs,mins)

def format_dt(dt):
    """
    Formats a datetime object into string that calendar needs in format YYYY-MM-DDTHH:MM:SS
    """
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def create_available_event(date_arr):
    _, month_str, day_str, year_str = date_arr
    month = datetime.strptime(month_str, "%B").month  # Convert month name to month number
    date = year_str + "-" + str(month) + "-" + day_str
    event = {
            'summary': "Available for CTRC",
            'description': "Scheduled for today",
            'start': {
                'date': date,
                'timeZone': 'America/Denver'
                },
            'end': {
                'date': date,
                'timeZone': 'America/Denver'
                },
            }
    return event

def create_event(date_arr,time,duration,name):
    """
    Create an event object for the given date and time scraped from 
     the docx file

     Date will be in the following format: ["Monday", "January", "10", "2024"]
     time will be in the following format: "1230"
     duration will be an integer in minutes
     name is the patient name as a string

    """
    dt = make_datetime(date_arr,time)
    start_time = format_dt(dt)
    end_time = format_dt(dt + timedelta(minutes=duration))

    event = {
            'summary': "CTRC Trial with Patient " + name,
            'description': 'CTRC Trial',
            'start': {
                'dateTime': start_time,
                'timeZone': 'America/Denver',
                },
            'end': {
                'dateTime': end_time,
                'timeZone': 'America/Denver',
                },
            }

    return event

def convertUTC(time,timezone="America/Denver"):
    """
    Convert local timestamp string to UTC format
    """

    local_tz = pytz.timezone(timezone)
    
    local_time = datetime.strptime(time, '%Y-%m-%dT%H:%M:%S')
    
    localized_time = local_tz.localize(local_time)
    
    utc_time = localized_time.astimezone(pytz.utc)
    
    return utc_time.isoformat()

def clearCalendar(service,calendar):
    eventList = service.events().list(calendarId=calendar['id'],singleEvents=True,maxResults=9999).execute().get('items',[])
    if not eventList:
        print("No events found to delete")
        return
    print(f"Found {len(eventList)} events. Deleting all...")
    for event in eventList:
        try:
            service.events().delete(calendarId=calendar['id'],eventId=event['id']).execute()
        except Exception as e:
            print(f"Failed to delete event {event['id']}")
    print("Done deleting events")

def upload_events(service,calendar,events):
    print(f"Adding {len(events)} scheduled events")
    for event in events:
        service.events().insert(calendarId=calendar['id'], body=event).execute()

def get_file_ids(service):
    results = service.files().list(pageSize=10,fields="nextPageToken, files(id,name)").execute()
    items = results.get("files",[])
    for item in items:
        if item['name'] == FOLDER:
            query = "'" + item['id'] + "' in parents"
            results = service.files().list(q = query, fields="nextPageToken, files(id, name)").execute()
            items = results.get("files",[])
            if len(items) > 0:
                key = "^[0-9]{4}-[0-9]+-[0-9]+\.docx"
                files = []
                for x in items:
                    if re.search(key,x['name']) is not None:
                        files.append(x['id'])
    print(f"Found {len(files)} schedule files")
    return files

def read_file(service,file_id):
    # Create a request to download the file
    request = service.files().get_media(fileId=file_id)
    
    # Create a BytesIO stream to hold the file content
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)

    done = False
    while done is False:
        status, done = downloader.next_chunk()

    print(f"Loaded file {file_id}")

    # Reset the stream position to the beginning
    file_stream.seek(0)

    doc = Document(file_stream)

    return doc

def get_events(service,file_ids,docName="DL"):
    """
    Loop through all file_ids and create calendar events
    """
    events = []
    for id in file_ids:
        doc = read_file(service,id)
        events.extend(scrape_docx(doc,docName))
    return events

def main():

    docNames = ["DL","CR","LG"]
    calendarNames = ["CTRC Physician Schedule","Christine CTRC Schedule","Laurence CTRC Schedule"]

    # Open services to access google account
    calendar_service,drive_service = open_services()

    # Locate all the schedule files
    file_ids = get_file_ids(drive_service)

    # Get events from all these files for each doc
    for i in range(len(docNames)):
        events = get_events(drive_service,file_ids,docNames[i])
        calendar = add_calendar(calendar_service,calendarNames[i])
        clearCalendar(calendar_service,calendar)
        upload_events(calendar_service,calendar,events)

if __name__ == '__main__':
    main()
    #calendarNames = ["CTRC Physician Schedule","Christine CTRC Schedule"]
    #calendar_service,drive_service = open_services()
    #calendar = add_calendar(calendar_service,calendarNames[0])
    #result = calendar_service.events().list(calendarId=calendar['id']).execute()
    #eventList =  result.get('items',[])
    #print(eventList)
