#!/usr/bin/env python3
"""
Scrapes active (valid) FSC certificates from the FSC Certificates Public Dashboard
(a Power BI "publish to web" report):
  https://app.powerbi.com/view?r=<token>

Power BI reports have no HTML to scrape, but a published report exposes its data
through the public `querydata` API. We replay the exact semantic query behind the
report's detail-table visual (captured once) and page through it with Power BI's
restart tokens — so this is a plain HTTP scraper, no browser required.

Output columns:
  Licence Code, Certificate Code, Certificate Type, Status, Controlled Wood,
  Valid From, Valid To, Organization, Role, Site Status, State/Province, Country
"""

import copy
import csv
import json
import re
import sys
import time
import uuid
from datetime import datetime, timezone

import requests

# The published report. The resource key (auth) is the "k" field of the ?r= token;
# "c":9 selects the West-Europe query cluster.
RESOURCE_KEY = "7e74c25a-e015-435a-a16c-98af7bb481cd"
QUERYDATA_URL = ("https://wabi-west-europe-f-primary-api.analysis.windows.net"
                 "/public/reports/querydata?synchronous=true")
WINDOW = 500                # rows per request (the visual's page size)
REQUEST_DELAY = 0.1

# Semantic query for the certificate detail-table visual, captured from the report.
# Columns (projection order) map to OUTPUT_FIELDS below. Filters: status = Valid,
# Cert_Status not Applicant/null, sites not hidden; ordered by start date desc.
QUERY_BODY = json.loads(r"""
{"version":"1.0.0","queries":[{"Query":{"Commands":[{"SemanticQueryDataShapeCommand":{"Query":{"Version":2,"From":[{"Name":"c","Entity":"Certificate","Type":0},{"Name":"o","Entity":"Organization","Type":0},{"Name":"c1","Entity":"Country_R","Type":0}],"Select":[{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"License"},"Name":"Certificate.License"},{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"Full_Certificate_Code__c"},"Name":"Certificate.Full_Certificate_Code__c"},{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"status"},"Name":"Certificate.status"},{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"CW"},"Name":"Certificate.CW"},{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"Date_From__c"},"Name":"Certificate.Date_From__c"},{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"Date_To__c"},"Name":"Certificate.Date_To__c"},{"Column":{"Expression":{"SourceRef":{"Source":"o"}},"Property":"Organization Name"},"Name":"Organization.Name"},{"Column":{"Expression":{"SourceRef":{"Source":"o"}},"Property":"Type (groups)"},"Name":"Organization.Type (groups)"},{"Column":{"Expression":{"SourceRef":{"Source":"o"}},"Property":"Site status"},"Name":"Organization.Site status"},{"Column":{"Expression":{"SourceRef":{"Source":"o"}},"Property":"State_County__c"},"Name":"Organization.State_County__c"},{"Column":{"Expression":{"SourceRef":{"Source":"c1"}},"Property":"Country/Area"},"Name":"Country_R.Country/Region","NativeReferenceName":"Country/Region"}],"Where":[{"Condition":{"In":{"Expressions":[{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"status"}}],"Values":[[{"Literal":{"Value":"'Valid'"}}]]}}},{"Condition":{"Not":{"Expression":{"In":{"Expressions":[{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"Cert_Status__c"}}],"Values":[[{"Literal":{"Value":"null"}}],[{"Literal":{"Value":"'Applicant'"}}]]}}}}},{"Condition":{"In":{"Expressions":[{"Column":{"Expression":{"SourceRef":{"Source":"o"}},"Property":"Hide_Site__c"}}],"Values":[[{"Literal":{"Value":"false"}}]]}}}],"OrderBy":[{"Direction":2,"Expression":{"Column":{"Expression":{"SourceRef":{"Source":"c"}},"Property":"Date_From__c"}}}]},"Binding":{"Primary":{"Groupings":[{"Projections":[0,1,2,3,4,5,6,7,8,9,10],"ShowItemsWithNoData":[0,1,2,3,4,5,6,7,8,9,10]}]},"DataReduction":{"DataVolume":3,"Primary":{"Window":{"Count":500}}},"Version":1},"ExecutionMetricsKind":1}}]},"QueryId":"","ApplicationContext":{"DatasetId":"d94c3e96-8435-4af1-b2ba-21ef6defaa95","Sources":[{"ReportId":"8be6b660-e5ba-422d-951f-797c9e2a4af7","VisualId":"8eb8554a410440b2a897"}]}}],"cancelQueries":[],"modelId":542581}
""")

# Projection index -> output column name.
OUTPUT_FIELDS = [
    "Licence Code", "Certificate Code", "Status", "Controlled Wood",
    "Valid From", "Valid To", "Organization", "Role", "Site Status",
    "State/Province", "Country",
]
DATE_COLS = {4, 5}          # Valid From / Valid To are epoch-millis
FIELDNAMES = (OUTPUT_FIELDS[:2] + ["Certificate Type"] + OUTPUT_FIELDS[2:])
NCOL = len(OUTPUT_FIELDS)
CERT_TYPE_RE = re.compile(r"-((?:CW/FM|FM/COC|CFM|COC|FM|CW))-", re.I)

# The visual's default order (Date_From) has heavy ties, which makes restart-token
# paging overlap and stall. Re-order by a near-unique multi-column key so each
# window advances cleanly. (Rows are per certificate *site*, hence multiple keys.)
_c = {"SourceRef": {"Source": "c"}}
_o = {"SourceRef": {"Source": "o"}}
ORDER_BY = [
    {"Direction": 1, "Expression": {"Column": {"Expression": _c, "Property": "License"}}},
    {"Direction": 1, "Expression": {"Column": {"Expression": _c, "Property": "Full_Certificate_Code__c"}}},
    {"Direction": 1, "Expression": {"Column": {"Expression": _o, "Property": "Organization Name"}}},
    {"Direction": 1, "Expression": {"Column": {"Expression": _o, "Property": "State_County__c"}}},
]
QUERY_BODY["queries"][0]["Query"]["Commands"][0][
    "SemanticQueryDataShapeCommand"]["Query"]["OrderBy"] = ORDER_BY


def post_query(session, body):
    headers = {
        "x-powerbi-resourcekey": RESOURCE_KEY,
        "content-type": "application/json;charset=UTF-8",
        "accept": "application/json, text/plain, */*",
        "referer": "https://app.powerbi.com/",
        "activityid": str(uuid.uuid4()),
        "requestid": str(uuid.uuid4()),
        "user-agent": "Mozilla/5.0 FSC-cert-scraper/1.0",
    }
    r = session.post(QUERYDATA_URL, data=json.dumps(body).encode(), headers=headers, timeout=90)
    r.raise_for_status()
    return r.json()["results"][0]["result"]["data"]["dsr"]["DS"][0]


def decode_window(ds):
    """Decode Power BI's compressed DataShape (DM0) into a list of value rows.

    Each row after the first carries only changed columns: the `R` bitmask marks
    columns repeated from the previous row and `Ø` marks nulls; dictionary columns
    hold an index into `ValueDicts`."""
    dm = ds["PH"][0]["DM0"]
    vdicts = ds.get("ValueDicts", {})
    col_dict = {i: c.get("DN") for i, c in enumerate(dm[0]["S"])}  # col -> value-dict name

    rows, prev = [], [None] * NCOL
    for item in dm:
        repeat = item.get("R", 0)
        nulls = item.get("Ø", 0)
        values = item.get("C", [])
        vi = 0
        raw = []
        for col in range(NCOL):
            if nulls & (1 << col):
                raw.append(None)
            elif repeat & (1 << col):
                raw.append(prev[col])
            else:
                raw.append(values[vi])
                vi += 1
        prev = raw
        # resolve dictionary indices to strings
        resolved = []
        for col in range(NCOL):
            v = raw[col]
            dn = col_dict.get(col)
            if dn and isinstance(v, int):
                v = vdicts[dn][v]
            resolved.append(v)
        rows.append(resolved)
    return rows, ds.get("RT")


def to_date(ms):
    if not isinstance(ms, (int, float)):
        return ""
    return datetime.fromtimestamp(ms / 1000, timezone.utc).strftime("%Y-%m-%d")


def to_record(row):
    rec = {}
    for i, name in enumerate(OUTPUT_FIELDS):
        v = row[i]
        rec[name] = to_date(v) if i in DATE_COLS else ("" if v is None else v)
    m = CERT_TYPE_RE.search(rec["Certificate Code"] or "")
    rec["Certificate Type"] = m.group(1).upper() if m else ""
    return {k: rec[k] for k in FIELDNAMES}


def main():
    session = requests.Session()
    print("Querying FSC public dashboard …")
    records, seen = [], set()
    restart = prev_restart = None
    page = 0
    while True:
        page += 1
        body = copy.deepcopy(QUERY_BODY)
        window = body["queries"][0]["Query"]["Commands"][0][
            "SemanticQueryDataShapeCommand"]["Binding"]["DataReduction"]["Primary"]["Window"]
        if restart is not None:
            window["RestartTokens"] = restart

        ds = post_query(session, body)
        rows, restart = decode_window(ds)

        new = 0
        for row in rows:
            rec = to_record(row)
            key = tuple(rec[k] for k in FIELDNAMES)   # per-site rows: dedupe whole row
            if key in seen:              # restart tokens overlap by one row
                continue
            seen.add(key)
            records.append(rec)
            new += 1
        print(f"  Page {page}: {len(rows)} rows ({new} new) — total {len(records)}",
              end="\r", flush=True)

        # Stop at the genuine end: a short window, no restart token, or the token
        # stopped advancing (guards against an overlap loop).
        if len(rows) < WINDOW or restart is None or restart == prev_restart:
            break
        prev_restart = restart
        time.sleep(REQUEST_DELAY)

    print(f"\nParsed {len(records)} valid certificates.")

    date_str = datetime.now(timezone.utc).strftime("%Y.%m.%d")
    for path in (f"FSC certificates {date_str}.csv", "FSC certificates latest.csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(records)
        print(f"Saved → {path}")


if __name__ == "__main__":
    main()
