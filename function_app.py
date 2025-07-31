import azure.functions as func
import logging
import requests
import json
import os
from datetime import datetime, timedelta

app = func.FunctionApp()

emails = {
    'Paul Thornton': 'paul.thornton@nbif.ca',
    'Cameron Horwood': 'cameron.horwood@nbif.ca',
    'Hallie Wedge': 'hallie.wedge@nbif.ca',
    'Jaime Christian': 'jaime.christian@nbif.ca',
    'Jessica McQuarrie': 'jessica.veinotte@nbif.ca',
    'Jessica Veinotte': 'jessica.veinotte@nbif.ca',
    'Jeff White': 'jeff.white@nbif.ca',
    'John Alexander': 'john.alexander@nbif.ca',
    'Eden Hoang': 'eden.hoang@nbif.ca',
    'Heather Libbey': 'heather.libbey@nbif.ca',
    'Hilary Lenihan': 'hilary.lenihan@nbif.ca',
    'Laila Theriault': 'laila.theriault@nbif.ca',
    'Kamrul Arefin': 'kamrul.arefin@nbif.ca',
    'Nitin Bhalla': 'nitin.bhalla@nbif.ca'
}

def get_email(name):
    return emails.get(name)

def get_requests():
    auth_token = os.getenv('AUTH_TOKEN')
    requests_url = "https://nbif.bamboohr.com/api/v1/time_off/requests"
    
    headers = {
        "accept": "application/json",
        "authorization": f"Basic {auth_token}"
    }
    
    today = datetime(2025, 7, 29).date()
    future = today + timedelta(days=90)
    params = {
        'start': today.isoformat(),
        'end': future.isoformat(),
        'status': 'approved'
    }

    try:
        response = requests.get(requests_url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()
    except requests.RequestException as e:
        logging.error(f"Error fetching time-off requests: {e}")
        return []
    except json.JSONDecodeError:
        logging.error("Failed to decode requests JSON")
        return []

    return data

def get_graph_token(client_id, client_secret, tenant_id):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        'client_id': client_id,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': client_secret,
        'grant_type': 'client_credentials'
    }
    try:
        response = requests.post(url, data=data)
        response.raise_for_status()
        return response.json()["access_token"]
    except Exception as e:
        logging.error(f"Failed to retrieve Microsoft Graph token: {e}")
        return None

def get_user_id(access_token, email):
    url = f"https://graph.microsoft.com/v1.0/users/{email}"
    headers = {'Authorization': f'Bearer {access_token}'}
    try:
        r = requests.get(url, headers=headers)
        r.raise_for_status()
        return r.json().get('id')
    except requests.RequestException:
        logging.warning(f"Could not find user with email: {email}")
        return None

def create_event(access_token, request_id, user_id, start, end):
    try:
        start_dt = datetime.strptime(start, "%Y-%m-%d").replace(hour=8, minute=30)
        end_dt = datetime.strptime(end, "%Y-%m-%d").replace(hour=17, minute=0)
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/calendar/events"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        event = {
            "subject": "Out of Office",
            "body": {
                "contentType": "Text",
                "content": f"Out of Office, BambooHR ID:{request_id}"
            },
            "start": {
                "dateTime": start_dt.isoformat(),
                "timeZone": "America/Halifax"
            },
            "end": {
                "dateTime": end_dt.isoformat(),
                "timeZone": "America/Halifax"
            },
            "showAs": "oof",
            "isAllDay": False,
            "attendees": [],
            "location": {"displayName": "Out of Office"},
            "categories": ["Out of Office"]
        }

        response = requests.post(url, headers=headers, json=event)
        response.raise_for_status()
        return response.status_code, response.json()
    except Exception as e:
        logging.error(f"Failed to create event for request {request_id}: {e}")
        return 500, {"error": str(e)}

def get_events(access_token, user_id):
    today = datetime(2025, 7, 29).isoformat()
    future = (datetime.now() + timedelta(days=90)).isoformat()

    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/calendarView"
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    params = {'startDateTime': today, 'endDateTime': future}

    try:
        r = requests.get(url, headers=headers, params=params)
        r.raise_for_status()
        data = r.json()
        return r.status_code, data
    except Exception as e:
        logging.error(f"Failed to fetch calendar events for user {user_id}: {e}")
        return 500, {"error": str(e)}

def check_duplicates(events_data, request_id):
    duplicate_marker = f"BambooHR ID:{request_id}"
    events = events_data.get('value', [])
    for event in events:
        try:
            if duplicate_marker in event['body']['content']:
                return True
        except KeyError:
            logging.warning(f"Event missing expected fields: {event}")
        except Exception as e:
            logging.error(f"Error checking for duplicates: {e}")
    return False

@app.timer_trigger(schedule="0 */5 * * * *", arg_name="myTimer", run_on_startup=True,
              use_monitor=False) 
def timer_trigger(myTimer: func.TimerRequest) -> None:
    if myTimer.past_due:
        logging.info('The timer is past due!')

    logging.info('WhosOut daily sync function started')
    
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    tenant_id = os.getenv('TENANT_ID')
    
    if not all([client_id, client_secret, tenant_id]):
        logging.error("Missing required environment variables")
        return
    
    requests_list = get_requests()
    logging.info(f"Retrieved {len(requests_list)} approved requests")

    access_token = get_graph_token(client_id, client_secret, tenant_id)
    if not access_token:
        logging.error("No Graph access token retrieved, aborting.")
        return

    if isinstance(requests_list, list) and requests_list:
        for request in requests_list:
            try:
                request_id = request['id']
                name = request['name']
                start = request['start']
                end = request['end']
                amount = request['amount']['amount']

                email = get_email(name)
                if not email:
                    logging.warning(f"{name} not found in email list")
                    continue

                user_id = get_user_id(access_token, email)
                if not user_id:
                    continue

                status, events_data = get_events(access_token, user_id)
                if status != 200:
                    logging.warning(f"Could not retrieve events for {name}")
                    continue
                
                # Process for Cameron Horwood only (as in original code)
                if name == 'Cameron Horwood':
                    if not check_duplicates(events_data, request_id):
                        status_code, response = create_event(access_token, request_id, user_id, start, end)
                        logging.info(f"Event created for {name} from {start} to {end} for {amount} day(s) (Status: {status_code})")
                    else:
                        logging.info(f"Duplicate event for {name}, request {request_id}")
            except KeyError as e:
                logging.error(f"Missing key in request object: {e}")
            except Exception as e:
                logging.exception(f"Unexpected error processing request {request}: {e}")
    else:
        logging.info("No approved requests to process")

    logging.info('WhosOut daily sync function completed')