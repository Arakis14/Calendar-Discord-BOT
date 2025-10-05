import os
import math
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timezone
from dotenv import load_dotenv

load_dotenv(os.path.expanduser("~/.bot.env"))

SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
if not SPREADSHEET_ID:
    raise ValueError("SPREADSHEET_ID not set in ~/.bot.env")
RANGE = os.environ["RANGE"]
DISCORD_WEBHOOK_URL = os.environ["DISCORD_WEBHOOK_URL"]
NOW = datetime.now(timezone.utc)
WEEK_NUMBER = NOW.isocalendar()[1]  # ISO week number (1–53)

# --- color handling with fuzzy match ---
import math
REFERENCE_COLORS = {
    "available": (0.0, 1.0, 0.0),   # green
    "tentative": (1.0, 1.0, 0.0),   # yellow
    "unavailable": (1.0, 0.0, 0.0), # red
    "unknown": (1.0, 1.0, 1.0),     # white / empty
}
def normalize_color(bg):
    return (
        bg.get("red", 1.0),
        bg.get("green", 1.0),
        bg.get("blue", 1.0),
    )
def color_distance(c1, c2):
    return math.sqrt((c1[0]-c2[0])**2+(c1[1]-c2[1])**2+(c1[2]-c2[2])**2)
def color_to_status(bg):
    color = normalize_color(bg or {})
    best_match, min_dist = "unknown", float("inf")
    for status, ref in REFERENCE_COLORS.items():
        dist = color_distance(color, ref)
        if dist < min_dist:
            min_dist = dist
            best_match = status
    return best_match

# --- Google Sheets fetch ---
def fetch_griddata(spreadsheet_id, ranges):
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    resp = service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=[ranges], includeGridData=True).execute()
    return resp.get("sheets", [])[0]["data"][0].get("rowData", [])

# --- Build weekly message ---
def build_week_message(rowData):
    header_values = rowData[0].get("values", [])
    times = [v.get("formattedValue", "").strip() for v in header_values[2:]]
    week_summary, day_name, players = [], None, []

    for row in rowData[1:]:
        values = row.get("values", [])
        if not values:
            continue
        day_cell = values[0].get("formattedValue", "").strip() if len(values) > 0 else ""
        player_name = values[1].get("formattedValue", "").strip() if len(values) > 1 else ""

        if day_cell and day_cell.lower() not in ("dzień:", "kto:"):
            if day_name and players:
                day_summary = process_day(day_name, times, players)
                if day_summary:   # skip empty string
                    week_summary.append(day_summary)
            day_name, players = day_cell, []

        if not player_name or player_name.lower() in ("kto", "kto:"):
            continue

        player_statuses = {}
        for col_idx, time in enumerate(times, start=2):
            if col_idx < len(values):
                cell = values[col_idx]
                bg = cell.get("userEnteredFormat", {}).get("backgroundColor", {})
                status = color_to_status(bg)
                player_statuses[time] = status
        players.append((player_name, player_statuses))

    if day_name and players:
        day_summary = process_day(day_name, times, players)
        if day_summary:
            week_summary.append(day_summary)
    #debug
    #timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    if not week_summary:
        return ""  # always return a string

    return f"**Kalendarzyk grania na tydzień {WEEK_NUMBER}** \n\n" + "\n\n".join(week_summary)

# --- Compress day slots into ranges ---
def process_day(day_name, times, players):
    total_players = len(players)
    availability_by_time = {t: {} for t in times}
    for player_name, statuses in players:
        for t in times:
            availability_by_time[t][player_name] = statuses.get(t, "unknown")

    results = []
    for t in times:
        statuses = list(availability_by_time[t].values())
        counts = {
            "available": statuses.count("available"),
            "tentative": statuses.count("tentative"),
            "unavailable": statuses.count("unavailable"),
            "unknown": statuses.count("unknown"),
        }
        if counts["available"] == total_players:
            results.append((t, "✅ Wszyscy dostępni! Łałałiła"))
        elif counts["available"] == 2 and counts["tentative"] == 1:
            results.append((t, f"⚠️ Granie możliwe, dostępni (2/{total_players}, 1 niepewny)"))
        else:
            results.append((t, None))

    # compress ranges
    compressed, start, prev_summary = [], None, None
    for i, (time, summary) in enumerate(results):
        if not summary:
            if prev_summary:
                compressed.append((start, results[i-1][0], prev_summary))
                start, prev_summary = None, None
            continue
        if prev_summary == summary:
            pass
        else:
            if prev_summary:
                compressed.append((start, results[i-1][0], prev_summary))
            start, prev_summary = time, summary
    if prev_summary:
        compressed.append((start, results[-1][0], prev_summary))

    if not compressed:
        return ""  # always return a string

    messages = [f"__{day_name}__"]
    for start, end, summary in compressed:
        if start == end:
            messages.append(f"**{start}** — {summary}")
        else:
            messages.append(f"**{start}–{end}** — {summary}")

    return "\n".join(messages)



# --- Discord send ---
def send_to_discord(message):
    if not DISCORD_WEBHOOK_URL:
        print("No Discord webhook configured.")
        return
    resp = requests.post(DISCORD_WEBHOOK_URL, json={"content": message})
    if resp.status_code != 204:
        print(f"Discord error {resp.status_code}: {resp.text}")

# --- Main ---
def main():
    sheet_name = f"Week_{WEEK_NUMBER}"
    sheet_name_and_range = f"{sheet_name}!{RANGE}"
    data = fetch_griddata(SPREADSHEET_ID, sheet_name_and_range)
    msg = build_week_message(data)
    print(msg)
    send_to_discord(msg)

if __name__ == "__main__":
    main()
