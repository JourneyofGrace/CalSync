import os
import requests
import msal

# --- Configuration ---
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

def clean_duplicates():
    token = get_graph_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    print(f"Scanning {TARGET_MAILBOX} for duplicates...")
    
    # Fetch all events (grabbing only the ID, Subject, and Times to make it fast)
    url = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/events?$select=id,subject,start,end&$top=1000"
    
    seen_fingerprints = set()
    delete_count = 0
    
    while url:
        res = requests.get(url, headers=headers)
        if res.status_code != 200:
            print(f"Error fetching events: {res.text}")
            break
            
        data = res.json()
        events = data.get("value", [])
        
        for event in events:
            event_id = event['id']
            subject = event.get('subject', 'Untitled')
            # Extract the raw UTC time strings
            start_time = event.get('start', {}).get('dateTime', '')
            end_time = event.get('end', {}).get('dateTime', '')
            
            # Create a unique fingerprint: "Title_StartTime_EndTime"
            fingerprint = f"{subject}_{start_time}_{end_time}"
            
            if fingerprint in seen_fingerprints:
                # We have seen this exact event at this exact time before. It's a clone.
                print(f"Deleting duplicate: {subject} at {start_time}")
                del_url = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/events/{event_id}"
                del_res = requests.delete(del_url, headers=headers)
                
                if del_res.status_code == 204:
                    delete_count += 1
                else:
                    print(f"  -> Failed to delete: {del_res.text}")
            else:
                # First time seeing this event. Add it to the safe list.
                seen_fingerprints.add(fingerprint)
                
        url = data.get("@odata.nextLink")
        
    print(f"\nSweep Complete. Surgically removed {delete_count} duplicates. The originals are safe.")

if __name__ == "__main__":
    clean_duplicates()
