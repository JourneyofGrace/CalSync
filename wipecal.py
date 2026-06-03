import os
import requests
import msal

# --- Configuration (Using the same aggressive scrubbing) ---
TENANT_ID = os.environ.get("TENANT_ID", "").strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')
CLIENT_ID = os.environ.get("CLIENT_ID", "").strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "").strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')
TARGET_MAILBOX = os.environ.get("TARGET_MAILBOX", "").strip().replace('"', '').replace("'", "").replace('\\n', '').replace('\\r', '')

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

def get_graph_token():
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result: return result["access_token"]
    raise Exception(f"Failed to acquire token: {result.get('error_description')}")

def nuke_calendar():
    token = get_graph_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    print(f"Fetching ALL events in {TARGET_MAILBOX}...")
    
    # Notice we removed all filters. We are just asking for every event ID.
    url = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/events?$select=id,subject&$top=1000"
    
    delete_count = 0
    while url:
        res = requests.get(url, headers=headers)
        if res.status_code != 200:
            print(f"Error fetching events: {res.text}")
            break
            
        data = res.json()
        events = data.get("value", [])
        
        if not events:
            break
            
        for event in events:
            print(f"Deleting: {event.get('subject', 'Untitled Event')}")
            del_url = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/events/{event['id']}"
            del_res = requests.delete(del_url, headers=headers)
            
            if del_res.status_code == 204:
                delete_count += 1
            else:
                print(f"  -> Failed to delete: {del_res.text}")
                
        url = data.get("@odata.nextLink")
        
    print(f"\nCalendar Wipe Complete. Successfully deleted {delete_count} events.")

if __name__ == "__main__":
    nuke_calendar()
