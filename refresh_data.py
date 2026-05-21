#!/usr/bin/env python3
"""
refresh_data.py — pull fresh Instagram data from Apify and rebuild the dashboard.

WHAT IT DOES
  1. Runs your Apify Instagram scrapers (your account + competitors)
  2. Saves the results into data/  (instagram_account.json, instagram_competitors.json)
  3. Re-runs build_dashboard.py so the dashboard reflects the new data
  Your dashboard edits, This Week checklist, collabs etc. are NOT touched —
  those live in the browser; only the analytics data is refreshed.

------------------------------------------------------------------------------
ONE-TIME SETUP
------------------------------------------------------------------------------
  1. Get your Apify API token:
       https://console.apify.com/account/integrations  ->  "Personal API tokens"
  2. Save the token in a file named `.apify_token` next to this script
     (one line, just the token). That file is git-ignored, so the token is
     never committed. Alternatively, set an APIFY_TOKEN environment variable.
  3. Check the JOBS list below — the usernames are already set to your handle
     and the two competitors. Adjust resultsLimit if you want more posts.

  TIP: if a job fails with an "input" error, open your last successful run in
  the Apify Console -> the run -> "Input" tab -> copy that JSON into "input"
  below. That guarantees the exact shape each actor expects.

------------------------------------------------------------------------------
RUN IT
------------------------------------------------------------------------------
  cd "/Users/shambhaviadhikari/Documents/Claude/Creator Insights"
  python3 refresh_data.py

------------------------------------------------------------------------------
SCHEDULE IT (run automatically, e.g. every morning at 7am)
------------------------------------------------------------------------------
  Make sure the `.apify_token` file exists (cron reads it automatically), then:

      crontab -e

  add this line (one line), then save:

      0 7 * * * cd "/Users/shambhaviadhikari/Documents/Claude/Creator Insights" && /usr/bin/python3 refresh_data.py >> refresh.log 2>&1

  That runs the refresh daily at 7:00am and logs to refresh.log.
  To open the dashboard from anywhere, host the generated HTML on a free
  static host (Netlify / GitHub Pages) and re-upload it after each refresh.
------------------------------------------------------------------------------
"""
import os
import sys
import json
import time
import subprocess
import urllib.request
import urllib.error

# ============================================================================
# CONFIG  — edit this section
# ============================================================================
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")


def _read_token():
    """Token resolution — kept OUT of this file so it never lands in git:
       1. the APIFY_TOKEN environment variable, or
       2. a `.apify_token` file next to this script (git-ignored)."""
    env = os.environ.get("APIFY_TOKEN", "").strip()
    if env:
        return env
    try:
        with open(os.path.join(HERE, ".apify_token")) as fh:
            return fh.read().strip()
    except OSError:
        return ""


APIFY_TOKEN = _read_token()

# Each job runs one Apify actor and saves its dataset to a file the dashboard
# reads. The dashboard expects: instagram_account.json (your posts) and
# instagram_competitors.json (competitor posts).
JOBS = [
    {
        "name": "your account (@journeysbysam)",
        "actor": "apify~instagram-scraper",
        "input": {
            "directUrls": ["https://www.instagram.com/journeysbysam/"],
            "resultsType": "posts",
            "resultsLimit": 50,
            "addParentData": False,
        },
        "out": os.path.join(DATA_DIR, "instagram_account.json"),
    },
    {
        "name": "competitors",
        "actor": "apify~instagram-post-scraper",
        "input": {
            "username": ["peekinourjournal", "tripuntraveled"],
            "resultsLimit": 30,
        },
        "out": os.path.join(DATA_DIR, "instagram_competitors.json"),
    },
]

API = "https://api.apify.com/v2"
POLL_SECONDS = 10          # how often to check if the run finished
MAX_WAIT_MINUTES = 15      # give up waiting after this long
# ============================================================================


def _req(url, method="GET", body=None):
    """Minimal JSON HTTP call using only the standard library."""
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(
        url, data=data, method=method,
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=90) as resp:
        raw = resp.read().decode()
    return json.loads(raw) if raw else {}


def run_actor(actor, inp):
    """Start an Apify actor run, wait for it to finish, return the dataset items."""
    start = _req(f"{API}/acts/{actor}/runs?token={APIFY_TOKEN}", "POST", inp)
    run = start["data"]
    run_id = run["id"]
    dataset_id = run["defaultDatasetId"]
    print(f"   run {run_id} started, waiting for it to finish...")

    deadline = time.time() + MAX_WAIT_MINUTES * 60
    while True:
        status = _req(f"{API}/actor-runs/{run_id}?token={APIFY_TOKEN}")["data"]["status"]
        if status == "SUCCEEDED":
            break
        if status in ("FAILED", "ABORTED", "TIMED-OUT"):
            raise RuntimeError(f"Apify run ended with status {status}")
        if time.time() > deadline:
            raise RuntimeError(f"gave up after {MAX_WAIT_MINUTES} min (run still {status})")
        time.sleep(POLL_SECONDS)

    items = _req(f"{API}/datasets/{dataset_id}/items?token={APIFY_TOKEN}&format=json&clean=true")
    return items


def main():
    if not APIFY_TOKEN:
        sys.exit("ERROR: no Apify token. Put it in a '.apify_token' file next to "
                 "refresh_data.py, or set the APIFY_TOKEN environment variable.")

    os.makedirs(DATA_DIR, exist_ok=True)
    ok = 0
    for job in JOBS:
        print(f"Pulling {job['name']} via {job['actor']} ...")
        try:
            items = run_actor(job["actor"], job["input"])
        except (urllib.error.HTTPError, urllib.error.URLError, RuntimeError, KeyError) as exc:
            print(f"   FAILED: {exc}")
            print("   (keeping the previous data file so the dashboard still works)")
            continue
        if not items:
            print("   WARNING: 0 records returned — keeping previous data file.")
            continue
        with open(job["out"], "w") as fh:
            json.dump(items, fh)
        print(f"   saved {len(items)} records -> {os.path.relpath(job['out'], HERE)}")
        ok += 1

    if ok == 0:
        sys.exit("No data refreshed — dashboard NOT rebuilt. Check your token / inputs above.")

    print("Rebuilding dashboard...")
    subprocess.run([sys.executable, os.path.join(HERE, "build_dashboard.py")], check=True)
    print("Done. Reload the dashboard in your browser to see the fresh data.")


if __name__ == "__main__":
    main()
