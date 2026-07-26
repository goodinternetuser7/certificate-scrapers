#!/usr/bin/env python3
"""
Scrapes the member register of the Roundtable on Sustainable Palm Oil (RSPO):
  https://rspo.org/search-members/

That page is only a WordPress shell around an iframe to a Salesforce
Visualforce page (`.../membership/AT_SearchMember_VFPage`), which in turn boots
a Lightning Out component (`c:AT_SearchMembers`). There is no HTML register to
parse — but the component's data comes from a *guest-accessible* Apex action,
so this is a plain-HTTP scraper (no browser): bootstrap the Aura framework
context from the Lightning Out app descriptor, then POST a single
`getApplicationsByFilter` action with an empty filter and `queryLimit = 0`
("no limit"). The whole register (~6.4k members) comes back in one response.

The action also reports its own `RecordCount`, so the pull is self-checking: if
a response is short we retry, then fall back to unioning per-category pulls
(the register partitions cleanly by membership category), which keeps each
query small should the register outgrow a single response.

Output columns:
  Membership Number, Member Name, Country, Membership Category, Sector,
  Status, Last Update, Group Members, Group Member Names, Profile URL
"""

import csv
import json
import time
from datetime import datetime, timezone

import requests

BASE = "https://rspo.my.salesforce-sites.com/membership"
BOOT_URL = f"{BASE}/c/AT_SearchMember_LightningOut.app"
AURA_URL = f"{BASE}/aura"
PAGE_URI = "/membership/AT_SearchMember_VFPage"
APP = "c:AT_SearchMember_LightningOut"
CONTROLLER = "AT_SearchMembers"
ACTION = "getApplicationsByFilter"
USER_AGENT = "Mozilla/5.0 RSPO-member-scraper/1.0"

# Membership categories (the `type__c` picklist) — used only by the fallback
# pull, to split the register into smaller queries.
CATEGORIES = ("Ordinary", "Associate", "Affiliate")
MAX_ATTEMPTS = 3
RETRY_DELAY = 5

FIELDNAMES = [
    "Membership Number", "Member Name", "Country", "Membership Category",
    "Sector", "Status", "Last Update", "Group Members", "Group Member Names",
    "Profile URL",
]


def boot_context(session):
    """Read the live framework UID + app descriptor from the Lightning Out app.

    Both rotate whenever Salesforce redeploys, so they must be fetched per run
    rather than hard-coded; a stale fwuid makes every Aura POST fail.
    """
    r = session.get(BOOT_URL,
                    params={"aura.format": "JSON", "aura.formatAdapter": "LIGHTNING_OUT"},
                    timeout=60)
    r.raise_for_status()
    ctx = r.json()["auraConfig"]["context"]
    return {"mode": "PROD", "fwuid": ctx["fwuid"], "app": APP,
            "loaded": ctx["loaded"], "dn": [], "globals": {}, "uad": True}


def aura_call(session, context, params):
    """POST one Apex action and return its returnValue."""
    message = {"actions": [{
        "id": "1;a",
        "descriptor": f"apex://{CONTROLLER}/ACTION${ACTION}",
        "callingDescriptor": f"markup://c:{CONTROLLER}",
        "params": params,
        "version": None,
    }]}
    r = session.post(AURA_URL,
                     params={"r": 1, f"other.{CONTROLLER}.{ACTION}": 1},
                     data={"message": json.dumps(message),
                           "aura.context": json.dumps(context),
                           "aura.pageURI": PAGE_URI,
                           "aura.token": "null"},
                     timeout=300)
    r.raise_for_status()
    body = r.text
    if body.startswith("while(1);"):            # Aura's JSON-hijacking guard
        body = body[len("while(1);"):]
    payload = json.loads(body)
    if payload.get("exceptionEvent"):
        raise RuntimeError(f"Aura exception: {str(payload.get('event'))[:300]}")
    action = payload["actions"][0]
    if action.get("state") != "SUCCESS":
        raise RuntimeError(f"Action {action.get('state')}: {str(action.get('error'))[:300]}")
    return action["returnValue"]


def query(session, context, category="", with_count=True):
    params = {"memberName": "", "category": category, "sector": "", "country": "",
              "status": "", "withGroupMember": True, "queryLimit": 0,
              "withRecordCount": with_count}
    rv = aura_call(session, context, params) or {}
    return rv.get("Applications") or [], rv.get("RecordCount"), rv.get("MembershipCMSSetting") or {}


def profile_base(cms):
    """Members' public profile URLs, as the component itself composes them."""
    host = (cms.get("RSPO_Website_Host_URL__c") or "https://rspo.org").rstrip("/")
    path = (cms.get("Member_CMS_Detail_URL__c") or "/members").strip("/")
    return f"{host}/{path}"


def to_row(app, base_url):
    account = app.get("Account__r") or {}
    groups = app.get("Group_Memberships__r") or []
    names = [g.get("Group_Member_name__c") or "" for g in groups]
    cms_id = account.get("CMS_IntID__c")
    return {
        "Membership Number": app.get("membershipNo__c") or "",
        "Member Name": account.get("Name") or "",
        "Country": app.get("Country__c") or "",
        "Membership Category": app.get("type__c") or "",
        "Sector": app.get("Sector__c") or "",
        "Status": app.get("Status__c") or "",
        "Last Update": app.get("Last_Update__c") or "",
        "Group Members": len(groups),
        # Group members are only listed by name; the register exposes no other
        # detail for them, so they stay a single joined cell rather than rows.
        "Group Member Names": "; ".join(n for n in names if n),
        "Profile URL": f"{base_url}/{cms_id}" if cms_id else "",
    }


def collect(session, context):
    """Pull every member, verifying against the register's own RecordCount."""
    members, expected, cms = {}, None, {}

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            apps, count, setting = query(session, context)
        except (requests.RequestException, ValueError, KeyError, RuntimeError) as exc:
            # A monthly unattended run shouldn't die on one flaky response.
            print(f"  Attempt {attempt} failed: {exc}")
            if attempt == MAX_ATTEMPTS:
                break
            time.sleep(RETRY_DELAY)
            continue
        cms = setting or cms
        expected = count if count is not None else expected
        for app in apps:
            members[app["Id"]] = app
        print(f"  Attempt {attempt}: {len(apps)} returned, {len(members)} unique"
              f"{f' of {expected}' if expected else ''}.")
        if expected and len(members) >= expected:
            return members, expected, cms
        if attempt < MAX_ATTEMPTS:
            time.sleep(RETRY_DELAY)

    # Short after retries: split the register by membership category so each
    # query returns a fraction of the rows, and union the results.
    print("  Full pull came up short — falling back to per-category pulls …")
    for category in CATEGORIES:
        try:
            apps, _, setting = query(session, context, category=category, with_count=False)
        except (requests.RequestException, ValueError, KeyError, RuntimeError) as exc:
            print(f"    {category} failed: {exc}")
            continue
        cms = setting or cms
        for app in apps:
            members[app["Id"]] = app
        print(f"    {category}: {len(apps)} returned, {len(members)} unique.")

    return members, expected, cms


def main():
    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    print(f"Bootstrapping Aura context from {BOOT_URL} …")
    context = boot_context(session)

    print("Querying the member register …")
    members, expected, cms = collect(session, context)
    if not members:
        raise SystemExit("Fetched 0 members — the Aura endpoint may have changed.")
    if expected and len(members) < expected:
        print(f"WARNING: register reports {expected} members but only "
              f"{len(members)} were retrieved.")

    base_url = profile_base(cms)
    records = [to_row(app, base_url) for app in members.values()]
    records.sort(key=lambda r: (r["Member Name"].lower(), r["Membership Number"]))

    active = sum(1 for r in records if r["Status"].lower() == "active")
    grouped = sum(1 for r in records if r["Group Members"])
    print(f"\nParsed {len(records)} members ({active} Active, "
          f"{grouped} with group members).")

    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"RSPO members {date_str}.csv", "RSPO members latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
        print(f"Saved → {path}")


if __name__ == "__main__":
    main()
