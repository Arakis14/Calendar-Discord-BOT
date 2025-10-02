import os
import math
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timezone
from dotenv import load_dotenv

load_dotenv()

SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
RANGE = os.environ.get("RANGE", "Week_40!A6:Q39")  # bigger range to cover whole week

def fetch_griddata(spreadsheet_id, ranges):
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    resp = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[ranges],
        includeGridData=True
    ).execute()
    sheets = resp.get("sheets", [])
    data = sheets[0]["data"][0].get("rowData", [])
    return data

REFERENCE_COLORS = {
    "available": (0.0, 1.0, 0.0),   # bright green
    "tentative": (1.0, 1.0, 0.0),   # yellow
    "unavailable": (1.0, 0.0, 0.0), # red
    "unknown": (1.0, 1.0, 1.0),     # white / empty
}

def normalize_color(bg):
    return (
        bg.get("red", 1.0),
        bg.get("green", 1.0),
        bg.get("blue", 1.0)
    )

def color_distance(c1, c2):
    return math.sqrt(
        (c1[0] - c2[0])**2 +
        (c1[1] - c2[1])**2 +
        (c1[2] - c2[2])**2
    )

def color_to_status(bg):
    color = normalize_color(bg or {})
    best_match = "unknown"
    min_dist = float("inf")

    for status, ref_color in REFERENCE_COLORS.items():
        dist = color_distance(color, ref_color)
        if dist < min_dist:
            min_dist = dist
            best_match = status

    #print(f"DEBUG: Raw color={color} -> Closest={best_match} (dist={min_dist:.3f})")

    return best_match

def build_week_message(rowData):
    if not rowData or len(rowData) < 2:
        return "No timetable data found."

    # Row 1 (index 0) has times in C..end
    header_values = rowData[0].get("values", [])
    times = [v.get("formattedValue", "").strip() for v in header_values[2:]]  # skip A="dzień", B="kto:"
    #print("DEBUG: Times detected ->", times)

    week_summary = []
    day_name = None
    players = []

    for row in rowData[1:]:
        values = row.get("values", [])
        if not values:
            continue

        day_cell = values[0].get("formattedValue", "").strip() if len(values) > 0 else ""
        player_name = values[1].get("formattedValue", "").strip() if len(values) > 1 else ""

        # New day row
        if day_cell and day_cell.lower() not in ("dzień:", "kto:"):
            # process previous day if exists
            if day_name and players:
                week_summary.append(process_day(day_name, times, players))
            day_name = day_cell
            players = []  # reset player list for this new day

            # ⚠️ IMPORTANT: do NOT "continue" here
            # because this row ALSO has a player in col B

        # Skip headers or empty rows
        if not player_name or player_name.lower() in ("kto", "kto:"):
            continue

        # Collect player row
        player_statuses = {}
        for col_idx, time in enumerate(times, start=2):  # start at col C
            if col_idx < len(values):
                cell = values[col_idx]
                bg = cell.get("userEnteredFormat", {}).get("backgroundColor", {})
                status = color_to_status(bg)
                player_statuses[time] = status
        players.append((player_name, player_statuses))

    # Process last day
    if day_name and players:
        week_summary.append(process_day(day_name, times, players))

    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    return f"**Gaming Availability for the Week** (generated {timestamp})\n\n" + "\n\n".join(week_summary)

def process_day(day_name, times, players):
    #print(f"\n===== DEBUG: Processing {day_name} =====")
    total_players = len(players)
    availability_by_time = {t: {} for t in times}

    # Fill availability map
    for player_name, statuses in players:
        for t in times:
            availability_by_time[t][player_name] = statuses.get(t, "unknown")

    messages = [f"__{day_name}__"]
    for t in times:
        statuses = list(availability_by_time[t].values())
        counts = {
            "available": statuses.count("available"),
            "tentative": statuses.count("tentative"),
            "unavailable": statuses.count("unavailable"),
            "unknown": statuses.count("unknown"),
        }

        # Only keep important slots
        if counts["available"] == total_players:
            summary = "✅ Grane!"
            messages.append(f"**{t}** — {summary}")

        elif counts["available"] == 2 and counts["tentative"] == 1:
            summary = f"⚠️ Możliwe granie? (2/{total_players}, 1 być może)"
            messages.append(f"**{t}** — {summary}")

    # If no slots matched, skip the day completely
    if len(messages) == 1:
        return f"__{day_name}__\n(no good common slots)"

    return "\n".join(messages)


def main():
    data = fetch_griddata(SPREADSHEET_ID, RANGE)
    msg = build_week_message(data)
    print("\n===== FINAL SUMMARY =====")
    print(msg)

if __name__ == "__main__":
    main()