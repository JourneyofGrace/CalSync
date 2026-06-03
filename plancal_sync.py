import os
import requests
import msal
from icalendar import Calendar
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

# --- Configuration ---
TENANT_ID = os.environ["TENANT_ID"].strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')
CLIENT_ID = os.environ["CLIENT_ID"].strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')
CLIENT_SECRET = os.environ["CLIENT_SECRET"].strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')
TARGET_MAILBOX = os.environ["TARGET_MAILBOX"].strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')

# Timezone Configuration
LOCAL_TZ = ZoneInfo("America/Phoenix")
GRAPH_TZ_STRING = "America/Phoenix"

# Map feeds to their respective prefixes and Outlook Categories
PCO_FEEDS = [
    {
        "url": "https://calendar.planningcenteronline.com/icals/eJxj4ajmsGLLz2TO62Sy4kotzi8oqWa34khOzPFU4jY0MzI3M2OzYnMNsWIrzWQ2r4624i5ILErMLa5mAAClxQ8Kc13fbb9a4c94004a94cf9db15e7aa867aa2dcb22",
        "prefix": "⛪ ",
        "category": "Church Events"
    },
    {
        "url": "https://calendar.planningcenteronline.com/icals/eJxj4ajmsGLLz2TO62Sy4kotzi8oqWa34khOzPFU4rQ0NTNhs2JzDbFiK81kNq-OtuIuSCxKzC2uZgAAi3cOpA==de88363ec427d412c80a30dd9f2185b68e050305",
        "prefix": "🏢 ",
        "category": "Office Events"
    }
]

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

def setup_category_colors(headers):
    """Programs the Shared Mailbox to assign specific colors to our categories."""
    print(f"Verifying Master Category colors for {TARGET_MAILBOX}...")
    url = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/outlook/masterCategories"
    
    # Microsoft Graph Color Presets: preset9 = Blue, preset4 = Green
    desired_categories = {
        "Church Events": "preset9", 
        "Office Events": "preset4"
    }
    
    res = requests.get(url, headers=headers)
    
    # Fail gracefully if the Entra ID app lacks MailboxSettings.ReadWrite
    if res.status_code == 403:
        print("  -> Notice: App lacks 'MailboxSettings.ReadWrite' permission. Colors must be set manually in Outlook.")
        return
    elif res.status_code != 200:
        print(f"  -> Warning: Could not fetch master categories: {res.text}")
        return

    existing_cats = {c["displayName"]: c for c in res.json().get("value", [])}

    for name, color in desired_categories.items():
        if name not in existing_cats:
            requests.post(url, headers=headers, json={"displayName": name, "color": color})
            print(f"  -> Created Master Category '{name}' with color {color}")
        else:
            if existing_cats[name].get("color") != color:
                cat_id = existing_cats[name]["id"]
                requests.patch(f"{url}/{cat_id}", headers=headers, json={"color": color})
                print(f"  -> Updated Master Category '{name}' to color {color}")

def format_graph_datetime(dt_obj):
    if isinstance(dt_obj, datetime):
        if dt_obj.tzinfo is None:
            dt_obj = dt_obj.replace(tzinfo=timezone.utc)
        local_dt = dt_obj.astimezone(LOCAL_TZ)
        return {"dateTime": local_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": GRAPH_TZ_STRING}
    elif isinstance(dt_obj, date):
        return {"dateTime": dt_obj.strftime("%Y-%m-%dT00:00:00"), "timeZone": GRAPH_TZ_STRING}

def get_existing_events(headers, cutoff_date_utc_str):
    url = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/events?$expand=extensions($filter=id eq '{EXTENSION_NAME}')&$filter=start/dateTime ge '{cutoff_date_utc_str}'&$top=1000"
    
    graph_events = {}
    all_fingerprints = {}
    
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            print(f"Error fetching existing events: {response.text}")
            break
        
        data = response.json()
        for event in data.get("value", []):
            extensions = event.get("extensions", [])
            pco_uid = next((ext.get("pcoUid") for ext in extensions if ext.get("id") == EXTENSION_NAME), None)
            
            subject = event.get("subject", "Untitled Event")
            start_str = event.get("start", {}).get("dateTime", "")[:19]
            end_str = event.get("end", {}).get("dateTime", "")[:19]
            
            fp = f"{subject}_{start_str}"
            all_fingerprints[fp] = event["id"]
            
            if pco_uid:
                graph_events[pco_uid] = {
                    "id": event["id"],
                    "subject": subject,
                    "start": start_str,
                    "end": end_str,
                    "categories": event.get("categories", [])
                }
        url = data.get("@odata.nextLink")
    return graph_events, all_fingerprints

def sync_calendars():
    token = get_graph_token()
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Prefer": f'outlook.timezone="{GRAPH_TZ_STRING}"'
    }
    
    # Trigger the color setup before processing events
    setup_category_colors(headers)
    
    cutoff_date_utc = datetime.now(timezone.utc) - timedelta(days=60)
    cutoff_date_utc_str = cutoff_date_utc.strftime("%Y-%m-%dT00:00:00Z")
    
    cutoff_date_local = datetime.now(LOCAL_TZ) - timedelta(days=60)
    cutoff_date_local_str = cutoff_date_local.strftime("%Y-%m-%dT00:00:00")
    
    print(f"Mapping existing calendar states for {TARGET_MAILBOX} (Looking back 60 days in {GRAPH_TZ_STRING})...")
    graph_events, all_fingerprints = get_existing_events(headers, cutoff_date_utc_str)
    
    pco_current_uids = set()
    endpoint = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/events"

    for feed in PCO_FEEDS:
        print(f"Processing remote feed: {feed['url']} (Category: {feed['category']})")
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
                
            start_graph = format_graph_datetime(component.get('dtstart').dt)
            end_graph = format_graph_datetime(component.get('dtend').dt)
            
            if start_graph["dateTime"] < cutoff_date_local_str:
                continue 

            pco_uid = str(component.get('uid'))
            pco_current_uids.add(pco_uid)
            
            raw_summary = str(component.get('summary', 'Untitled Event'))
            summary = f"{feed['prefix']}{raw_summary}"
            description = str(component.get('description', ''))
            target_category = feed['category']
            
            event_payload = {
                "subject": summary,
                "body": {"contentType": "Text", "content": description},
                "start": start_graph,
                "end": end_graph,
                "categories": [target_category]
            }

            fingerprint = f"{summary}_{start_graph['dateTime']}"

            if pco_uid in graph_events:
                existing = graph_events[pco_uid]
                if (existing["subject"] != summary or 
                    existing["start"] != start_graph["dateTime"] or 
                    existing["end"] != end_graph["dateTime"] or
                    target_category not in existing.get("categories", [])):
                    
                    print(f"Updating shifted or un-categorized event: {summary}")
                    patch_url = f"{endpoint}/{existing['id']}"
                    res_patch = requests.patch(patch_url, headers=headers, json=event_payload)
                    if res_patch.status_code not in [200, 204]:
                        print(f"  -> Failed to update: {res_patch.text}")
            else:
                if fingerprint in all_fingerprints:
                    print(f"Re-linking existing event: {summary}")
                    existing_id = all_fingerprints[fingerprint]
                    patch_url = f"{endpoint}/{existing_id}"
                    
                    event_payload["extensions"] = [
                        {
                            "@odata.type": "microsoft.graph.openTypeExtension",
                            "extensionName": EXTENSION_NAME,
                            "pcoUid": pco_uid
                        }
                    ]
                    requests.patch(patch_url, headers=headers, json=event_payload)
                    
                    stale_uids = [k for k, v in graph_events.items() if v["id"] == existing_id]
                    for k in stale_uids:
                        del graph_events[k]
                else:
                    print(f"Creating new event: {summary}")
                    event_payload["extensions"] = [
                        {
                            "@odata.type": "microsoft.graph.openTypeExtension",
                            "extensionName": EXTENSION_NAME,
                            "pcoUid": pco_uid
                        }
                    ]
                    res_post = requests.post(endpoint, headers=headers, json=event_payload)
                    if res_post.status_code not in [200, 201]:
                        print(f"  -> Failed to create: {res_post.text}")

    for old_pco_uid, event_meta in graph_events.items():
        if old_pco_uid not in pco_current_uids:
            print(f"Purging canceled/deleted event: {event_meta['subject']}")
            delete_url = f"{endpoint}/{event_meta['id']}"
            res_delete = requests.delete(delete_url, headers=headers)
            if res_delete.status_code != 204:
                print(f"  -> Failed to delete: {res_delete.text}")

    print("Sync process completed successfully.")

if __name__ == "__main__":
    sync_calendars()
