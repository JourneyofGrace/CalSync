import os
import time

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
# Outlook resources store open extensions under the legacy fully-qualified id
# ("Microsoft.OutlookServices.OpenTypeExtension.<name>"); newer responses can
# use the "microsoft.graph.openTypeExtension." prefix. $expand requires one of
# these ids in its filter, so query both and merge.
EXTENSION_FILTER_IDS = [
    f"Microsoft.OutlookServices.OpenTypeExtension.{EXTENSION_NAME}",
    f"microsoft.graph.openTypeExtension.{EXTENSION_NAME}",
]
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

def graph_request(method, url, **kwargs):
    """Call Microsoft Graph with a timeout and retries for throttling/transient errors."""
    kwargs.setdefault("timeout", 30)
    for attempt in range(5):
        res = requests.request(method, url, **kwargs)
        if res.status_code not in (429, 500, 502, 503, 504):
            return res
        retry_after = res.headers.get("Retry-After")
        delay = int(retry_after) if retry_after and retry_after.isdigit() else min(60, 2 ** attempt * 2)
        print(f"  -> Graph {res.status_code}; retrying in {delay}s ({attempt + 1}/5)")
        time.sleep(delay)
    return res

def setup_category_colors(headers):
    """Programs the Shared Mailbox to assign specific colors to our categories."""
    print(f"Verifying Master Category colors for {TARGET_MAILBOX}...")
    url = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/outlook/masterCategories"
    
    # Microsoft Graph Color Presets: preset9 = Blue, preset4 = Green
    desired_categories = {
        "Church Events": "preset9", 
        "Office Events": "preset4"
    }
    
    res = graph_request("GET", url, headers=headers)
    
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
            graph_request("POST", url, headers=headers, json={"displayName": name, "color": color})
            print(f"  -> Created Master Category '{name}' with color {color}")
        else:
            if existing_cats[name].get("color") != color:
                cat_id = existing_cats[name]["id"]
                graph_request("PATCH", f"{url}/{cat_id}", headers=headers, json={"color": color})
                print(f"  -> Updated Master Category '{name}' to color {color}")

def format_graph_datetime(dt_obj):
    if isinstance(dt_obj, datetime):
        if dt_obj.tzinfo is None:
            dt_obj = dt_obj.replace(tzinfo=timezone.utc)
        local_dt = dt_obj.astimezone(LOCAL_TZ)
        return {"dateTime": local_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": GRAPH_TZ_STRING}
    elif isinstance(dt_obj, date):
        return {"dateTime": dt_obj.strftime("%Y-%m-%dT00:00:00"), "timeZone": GRAPH_TZ_STRING}

def _fetch_events_for_extension(headers, cutoff_date_utc_str, extension_id):
    """Fetch events that carry one specific open-extension id."""
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/events"
        f"?$expand=extensions($filter=id eq '{extension_id}')"
        f"&$filter=start/dateTime ge '{cutoff_date_utc_str}'&$top=1000"
    )
    events = {}
    while url:
        response = graph_request("GET", url, headers=headers)
        if response.status_code != 200:
            # Abort rather than return a partial map: creating events from a
            # partial view would duplicate the entire calendar.
            raise RuntimeError(f"Error fetching existing events: {response.text}")

        data = response.json()
        for event in data.get("value", []):
            # Merge extensions if the same event comes back for both prefixes.
            if event["id"] not in events:
                events[event["id"]] = event
            else:
                have = {e.get("id") for e in events[event["id"]].get("extensions", [])}
                for ext in event.get("extensions", []):
                    if ext.get("id") not in have:
                        events[event["id"]].setdefault("extensions", []).append(ext)
        url = data.get("@odata.nextLink")
    return events


def get_existing_events(headers, cutoff_date_utc_str):
    # $expand=extensions *requires* a server-side id filter; it cannot be used
    # bare (ErrorGraphExtensionExpandRequiresFilter). Query the documented
    # Outlook prefix first, then the newer microsoft.graph prefix as fallback.
    events_by_id = {}
    for index, extension_id in enumerate(EXTENSION_FILTER_IDS):
        try:
            fetched = _fetch_events_for_extension(headers, cutoff_date_utc_str, extension_id)
        except RuntimeError as exc:
            if index == 0:
                raise
            print(f"  Warning: fallback extension query failed: {exc}")
            continue
        for event_id, event in fetched.items():
            if event_id not in events_by_id:
                events_by_id[event_id] = event
                continue
            have = {e.get("id") for e in events_by_id[event_id].get("extensions", [])}
            for ext in event.get("extensions", []):
                if ext.get("id") not in have:
                    events_by_id[event_id].setdefault("extensions", []).append(ext)

    graph_events = {}
    all_fingerprints = {}

    for event in events_by_id.values():
        extensions = event.get("extensions", [])
        pco_ext = next(
            (
                ext
                for ext in extensions
                if ext.get("extensionName") == EXTENSION_NAME
                or str(ext.get("id", "")).endswith(EXTENSION_NAME)
            ),
            None,
        )
        pco_uid = pco_ext.get("pcoUid") if pco_ext else None
        extension_id = pco_ext.get("id") if pco_ext else None

        subject = event.get("subject", "Untitled Event")
        start_str = event.get("start", {}).get("dateTime", "")[:19]
        end_str = event.get("end", {}).get("dateTime", "")[:19]

        fp = f"{subject}_{start_str}"
        all_fingerprints[fp] = {
            "id": event["id"],
            "subject": subject,
            "extensionId": extension_id,
            "pcoUid": pco_uid,
        }

        if pco_uid:
            graph_events[pco_uid] = {
                "id": event["id"],
                "subject": subject,
                "start": start_str,
                "end": end_str,
                "categories": event.get("categories", []),
                "isAllDay": bool(event.get("isAllDay", False))
            }
    return graph_events, all_fingerprints

def build_extension_object(pco_uid):
    return {
        "@odata.type": "microsoft.graph.openTypeExtension",
        "extensionName": EXTENSION_NAME,
        "pcoUid": pco_uid
    }

def build_extension_payload(pco_uid):
    return [build_extension_object(pco_uid)]

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
    feeds_ok = True

    for feed in PCO_FEEDS:
        print(f"Processing remote feed: {feed['url']} (Category: {feed['category']})")
        try:
            res = requests.get(feed['url'], timeout=60)
            res.raise_for_status()
            cal = Calendar.from_ical(res.content)
        except Exception as e:
            print(f"Failed to pull or parse feed: {e}")
            feeds_ok = False
            continue

        feed_event_count = 0
        for component in cal.walk('vevent'):
            dtstart = component.get('dtstart')
            if not component.get('uid') or not dtstart:
                continue

            dtstart_value = dtstart.dt
            dtend = component.get('dtend')
            if dtend:
                dtend_value = dtend.dt
            elif isinstance(dtstart_value, datetime):
                dtend_value = dtstart_value + timedelta(hours=1)
            else:
                dtend_value = dtstart_value + timedelta(days=1)

            start_graph = format_graph_datetime(dtstart_value)
            end_graph = format_graph_datetime(dtend_value)
            if not start_graph or not end_graph:
                continue

            is_all_day = not isinstance(dtstart_value, datetime)

            if start_graph["dateTime"] < cutoff_date_local_str:
                continue

            pco_uid = str(component.get('uid'))
            pco_current_uids.add(pco_uid)
            feed_event_count += 1
            
            raw_summary = str(component.get('summary', 'Untitled Event'))
            summary = f"{feed['prefix']}{raw_summary}"
            description = str(component.get('description', ''))
            target_category = feed['category']
            
            event_payload = {
                "subject": summary,
                "body": {"contentType": "Text", "content": description},
                "start": start_graph,
                "end": end_graph,
                "isAllDay": is_all_day,
                "categories": [target_category]
            }

            fingerprint = f"{summary}_{start_graph['dateTime']}"

            if pco_uid in graph_events:
                existing = graph_events[pco_uid]
                if (existing["subject"] != summary or
                    existing["start"] != start_graph["dateTime"] or
                    existing["end"] != end_graph["dateTime"] or
                    bool(existing.get("isAllDay", False)) != is_all_day or
                    target_category not in existing.get("categories", [])):

                    print(f"Updating shifted or un-categorized event: {summary}")
                    patch_url = f"{endpoint}/{existing['id']}"
                    res_patch = graph_request("PATCH", patch_url, headers=headers, json=event_payload)
                    if res_patch.status_code not in [200, 204]:
                        print(f"  -> Failed to update: {res_patch.text}")
                    else:
                        existing.update({
                            "subject": summary,
                            "start": start_graph["dateTime"],
                            "end": end_graph["dateTime"],
                            "categories": [target_category],
                            "isAllDay": is_all_day
                        })
            else:
                if fingerprint in all_fingerprints:
                    # The event already exists but isn't linked to this PCO UID:
                    # either it has no extension (created by an older version
                    # of this sync, or by hand) or it carries a stale UID (the
                    # event was re-created in Planning Center). Attach/refresh
                    # the extension in place instead of deleting and
                    # re-creating: churning the event ID every run races with
                    # Exchange (ConcurrentItemSave) and wipes Outlook edits.
                    print(f"Linking existing event: {summary}")
                    existing = all_fingerprints[fingerprint]
                    existing_id = existing["id"]

                    # Drop any stale UID entry for this event so the purge loop
                    # doesn't delete the event we are adopting.
                    stale_uids = [k for k, v in graph_events.items() if v["id"] == existing_id]
                    for k in stale_uids:
                        del graph_events[k]

                    if existing.get("extensionId"):
                        # Update the stale pcoUid on the existing extension.
                        ext_url = f"{endpoint}/{existing_id}/extensions/{existing['extensionId']}"
                        res_ext = graph_request("PATCH", ext_url, headers=headers, json=build_extension_object(pco_uid))
                    else:
                        ext_url = f"{endpoint}/{existing_id}/extensions"
                        res_ext = graph_request("POST", ext_url, headers=headers, json=build_extension_object(pco_uid))

                    if res_ext.status_code in [200, 201]:
                        print(f"  -> Attached/updated PCO UID on event {existing_id}")
                    elif res_ext.status_code == 409:
                        # Already linked (e.g. a race); treat as success.
                        print("  -> Extension already present; treating as linked.")
                    else:
                        print(f"  -> Failed to attach extension: {res_ext.text}")
                        continue

                    res_patch = graph_request("PATCH", f"{endpoint}/{existing_id}", headers=headers, json=event_payload)
                    if res_patch.status_code not in [200, 204]:
                        print(f"  -> Failed to update linked event: {res_patch.text}")

                    graph_events[pco_uid] = {
                        "id": existing_id,
                        "subject": summary,
                        "start": start_graph["dateTime"],
                        "end": end_graph["dateTime"],
                        "categories": [target_category],
                        "isAllDay": is_all_day
                    }
                    all_fingerprints.pop(fingerprint, None)
                else:
                    print(f"Creating new event: {summary}")
                    event_payload["extensions"] = build_extension_payload(pco_uid)
                    res_post = graph_request("POST", endpoint, headers=headers, json=event_payload)
                    if res_post.status_code not in [200, 201]:
                        print(f"  -> Failed to create: {res_post.text}")
                    else:
                        new_ev = res_post.json()
                        graph_events[pco_uid] = {
                            "id": new_ev["id"],
                            "subject": summary,
                            "start": start_graph["dateTime"],
                            "end": end_graph["dateTime"],
                            "categories": [target_category],
                            "isAllDay": is_all_day
                        }

        if feed_event_count == 0:
            print(f"Feed returned no usable events; skipping purge for safety: {feed['category']}")
            feeds_ok = False

    if feeds_ok:
        for old_pco_uid, event_meta in graph_events.items():
            if old_pco_uid not in pco_current_uids:
                print(f"Purging canceled/deleted event: {event_meta['subject']}")
                delete_url = f"{endpoint}/{event_meta['id']}"
                res_delete = graph_request("DELETE", delete_url, headers=headers)
                if res_delete.status_code != 204:
                    print(f"  -> Failed to delete: {res_delete.text}")
    else:
        print("Skipping purge because one or more feeds failed to load.")

    print("Sync process completed successfully.")

if __name__ == "__main__":
    sync_calendars()
