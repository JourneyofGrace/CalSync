import os
import requests
import msal
from icalendar import Calendar
from datetime import date, datetime, timedelta

# --- Configuration ---
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
GROUP_ID = os.environ["GROUP_ID"]

# Map feeds to their respective prefixes
PCO_FEEDS = [
    {
        "url": "https://calendar.planningcenteronline.com/icals/eJxj4ajmsGLLz2TO62Sy4kotzi8oqWa34khOzPFU4jY0MzI3M2OzYnMNsWIrzWQ2r4624i5ILErMLa5mAAClxQ8Kc13fbb9a4c94004a94cf9db15e7aa867aa2dcb22",
        "prefix": "⛪ "
    },
    {
        "url": "https://calendar.planningcenteronline.com/icals/eJxj4ajmsGLLz2TO62Sy4kotzi8oqWa34khOzPFU4rQ0NTNhs2JzDbFiK81kNq-OtuIuSCxKzC2uZgAAi3cOpA==de88363ec427d412c80a30dd9f2185b68e050305",
        "prefix": "🏢 "
    }
]

# Unique namespace for your schema extension inside your tenant
EXTENSION_NAME = "org.church.pcoSync"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

def get_graph_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed to acquire token: {result.get('error_description')}")

def format_graph_datetime(dt_obj):
    """Normalizes icalendar date/datetime components to UTC Graph strings."""
    if isinstance(dt_obj, datetime):
        return {"dateTime": dt_obj.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"}
    elif isinstance(dt_obj, date):
        return {"dateTime": dt_obj.strftime("%Y-%m-%dT00:00:00"), "timeZone": "UTC"}

def get_existing_group_events(headers):
    """Fetches existing group events that contain our custom PCO metadata extension."""
    start_date = (datetime.utcnow() - timedelta(days=30)).strftime("%Y-%m-%dT00:00:00Z")
    url = f"https://graph.microsoft.com/v1.0/groups/{GROUP_ID}/calendar/events?$expand=extensions($filter=id eq '{EXTENSION_NAME}')&$filter=start/dateTime ge '{start_date}'&$top=1000"
    
    existing_map = {}
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            print(f"Error fetching existing events: {response.text}")
            break
        
        data = response.json()
        for event in data.get("value", []):
            extensions = event.get("extensions", [])
            pco_uid = next((ext.get("pcoUid") for ext in extensions if ext.get("id") == EXTENSION_NAME), None)
            
            if pco_uid:
                existing_map[pco_uid] = {
                    "id": event["id"],
                    "subject": event["subject"],
                    "start": event["start"]["dateTime"][:19],
                    "end": event["end"]["dateTime"][:19]
                }
        url = data.get("@odata.nextLink")
    return existing_map

def sync_calendars():
    token = get_graph_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    print("Mapping existing Office 365 Group Calendar states...")
    graph_events = get_existing_group_events(headers)
    
    pco_current_uids = set()
    endpoint = f"https://graph.microsoft.com/v1.0/groups/{GROUP_ID}/events"

    for feed in PCO_FEEDS:
        print(f"Processing remote feed: {feed['url']}")
        try:
            res = requests.get(feed['url'])
            res.raise_for_status()
            cal = Calendar.from_ical(res.content)
        except Exception as e:
            print(f"Failed to pull or parse feed: {e}")
            continue

        for component in cal.walk('vevent'):
            if not component.get('uid'):
                continue
                
            pco_uid = str(component.get('uid'))
            pco_current_uids.add(pco_uid)
            
            raw_summary = str(component.get('summary', 'Untitled Event'))
            # Inject the emoji prefix
            summary = f"{feed['prefix']}{raw_summary}"
            description = str(component.get('description', ''))
            
            start_graph = format_graph_datetime(component.get('dtstart').dt)
            end_graph = format_graph_datetime(component.get('dtend').dt)
            
            event_payload = {
                "subject": summary,
                "body": {"contentType": "Text", "content": description},
                "start": start_graph,
                "end": end_graph,
            }

            if pco_uid in graph_events:
                existing = graph_events[pco_uid]
                if (existing["subject"] != summary or 
                    existing["start"] != start_graph["dateTime"] or 
                    existing["end"] != end_graph["dateTime"]):
                    
                    print(f"Updating shifted event: {summary}")
                    patch_url = f"{endpoint}/{existing['id']}"
                    requests.patch(patch_url, headers=headers, json=event_payload)
            else:
                print(f"Creating new event: {summary}")
                event_payload["extensions"] = [
                    {
                        "@odata.type": "microsoft.graph.openTypeExtension",
                        "extensionName": EXTENSION_NAME,
                        "pcoUid": pco_uid
                    }
                ]
                requests.post(endpoint, headers=headers, json=event_payload)

    for old_pco_uid, event_meta in graph_events.items():
        if old_pco_uid not in pco_current_uids:
            print(f"Purging canceled/deleted event: {event_meta['subject']}")
            delete_url = f"{endpoint}/{event_meta['id']}"
            requests.delete(delete_url, headers=headers)

    print("Sync process completed successfully.")

if __name__ == "__main__":
    sync_calendars()
