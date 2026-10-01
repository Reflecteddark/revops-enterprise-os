"""
Applies executive styling, colors, font styles, and cell merges
to '🎙️ ИИ_Аудит' (sheetId: 852624872) in Google Sheet:
https://docs.google.com/spreadsheets/d/1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc
"""

import gspread
import sys

sys.stdout.reconfigure(encoding="utf-8")

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key("1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc")
ws = sh.worksheet("🎙️ ИИ_Аудит")
sheet_id = ws.id

def color_rgb(r, g, b):
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}

c_navy_dark = color_rgb(15, 23, 42)      # #0F172A
c_navy_light = color_rgb(30, 41, 59)     # #1E293B
c_table_hdr = color_rgb(51, 65, 85)      # #334155
c_card_bg = color_rgb(248, 250, 252)     # #F8FAFC
c_blue_light = color_rgb(238, 242, 255)  # #EEF2FF
c_green_light = color_rgb(236, 253, 245) # #ECFDF5
c_red_light = color_rgb(254, 242, 242)   # #FEF2F2
c_amber_light = color_rgb(255, 251, 235) # #FFFBEB
c_border_gray = color_rgb(203, 213, 225) # #CBD5E1

c_text_white = color_rgb(255, 255, 255)
c_text_dark = color_rgb(15, 23, 42)
c_text_blue = color_rgb(37, 99, 235)
c_text_green = color_rgb(22, 163, 74)
c_text_red = color_rgb(220, 38, 38)
c_text_muted = color_rgb(100, 116, 139)

requests = []

# Helper to merge
def req_merge(r_start, r_end, c_start, c_end):
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r_start,
                "endRowIndex": r_end,
                "startColumnIndex": c_start,
                "endColumnIndex": c_end
            },
            "mergeType": "MERGE_ALL"
        }
    }

# Helper to format range
def req_format(r_start, r_end, c_start, c_end, bg_color=None, text_color=None, bold=False, font_size=10, halign="LEFT", valign="MIDDLE", wrap=False):
    cell_format = {
        "textFormat": {
            "fontFamily": "Segoe UI",
            "fontSize": font_size,
            "bold": bold
        },
        "horizontalAlignment": halign,
        "verticalAlignment": valign
    }
    if bg_color:
        cell_format["backgroundColor"] = bg_color
    if text_color:
        cell_format["textFormat"]["foregroundColor"] = text_color
    if wrap:
        cell_format["wrapStrategy"] = "WRAP"

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r_start,
                "endRowIndex": r_end,
                "startColumnIndex": c_start,
                "endColumnIndex": c_end
            },
            "cell": {"userEnteredFormat": cell_format},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"
        }
    }

# 1. Unmerge previous merges to avoid collision
requests.append({
    "unmergeCells": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 65,
            "startColumnIndex": 0,
            "endColumnIndex": 22
        }
    }
})

# 2. Merges
# Row 1: Title (A1:V1)
requests.append(req_merge(0, 1, 0, 22))

# Row 2: Nav Bar (6 tabs across 22 cols)
requests.append(req_merge(1, 2, 0, 3))    # A:C
requests.append(req_merge(1, 2, 3, 6))    # D:F
requests.append(req_merge(1, 2, 6, 10))   # G:J (Active!)
requests.append(req_merge(1, 2, 10, 13))  # K:M
requests.append(req_merge(1, 2, 13, 16))  # N:P
requests.append(req_merge(1, 2, 16, 22))  # Q:V

# Row 3-4 KPI Cards:
# Card 1: A:B (cols 0..2)
requests.append(req_merge(2, 3, 0, 2))
requests.append(req_merge(3, 4, 0, 2))
# Card 2: C:D (cols 2..4)
requests.append(req_merge(2, 3, 2, 4))
requests.append(req_merge(3, 4, 2, 4))
# Card 3: E:F (cols 4..6)
requests.append(req_merge(2, 3, 4, 6))
requests.append(req_merge(3, 4, 4, 6))
# Card 4: G:I (cols 6..9)
requests.append(req_merge(2, 3, 6, 9))
requests.append(req_merge(3, 4, 6, 9))
# Card 5: J:L (cols 9..12)
requests.append(req_merge(2, 3, 9, 12))
requests.append(req_merge(3, 4, 9, 12))
# Card 6: M:V (cols 12..22)
requests.append(req_merge(2, 3, 12, 22))
requests.append(req_merge(3, 4, 12, 22))

# Row 6: Section Leaderboard (A6:V6)
requests.append(req_merge(5, 6, 0, 22))
# Row 7: Leaderboard recommendation merge (J7:V7)
requests.append(req_merge(6, 7, 9, 22))
for r_idx in range(7, 12):
    requests.append(req_merge(r_idx, r_idx+1, 9, 22))
# Row 13: Leaderboard total recommendation
requests.append(req_merge(12, 13, 9, 22))

# Row 15: Section 13-criteria matrix (A15:V15)
requests.append(req_merge(14, 15, 0, 22))
# Row 27: Matrix total (A27:C27)
requests.append(req_merge(26, 27, 0, 3))

# Row 29: Section Quotes (A29:V29)
requests.append(req_merge(28, 29, 0, 22))
# Row 30: Quotes Headers
requests.append(req_merge(29, 30, 4, 6))    # E30:F30
requests.append(req_merge(29, 30, 6, 13))   # G30:M30
requests.append(req_merge(29, 30, 13, 22))  # N30:V30
for r_idx in range(30, 35):
    requests.append(req_merge(r_idx, r_idx+1, 4, 6))
    requests.append(req_merge(r_idx, r_idx+1, 6, 13))
    requests.append(req_merge(r_idx, r_idx+1, 13, 22))

# Row 37: Section 13 criteria reference (A37:V37)
requests.append(req_merge(36, 37, 0, 22))
# Row 38: Reference Headers
requests.append(req_merge(37, 38, 4, 7))    # E38:G38
requests.append(req_merge(37, 38, 7, 12))   # H38:L38
requests.append(req_merge(37, 38, 12, 22))  # M38:V38
for r_idx in range(38, 51):
    requests.append(req_merge(r_idx, r_idx+1, 4, 7))
    requests.append(req_merge(r_idx, r_idx+1, 7, 12))
    requests.append(req_merge(r_idx, r_idx+1, 12, 22))

# Row 53: Section ROI (A53:V53)
requests.append(req_merge(52, 53, 0, 22))
# Row 54: ROI Headers
requests.append(req_merge(53, 54, 0, 3))    # A54:C54
requests.append(req_merge(53, 54, 3, 5))    # D54:E54
requests.append(req_merge(53, 54, 5, 22))   # F54:V54
for r_idx in range(54, 60):
    requests.append(req_merge(r_idx, r_idx+1, 0, 3))
    requests.append(req_merge(r_idx, r_idx+1, 3, 5))
    requests.append(req_merge(r_idx, r_idx+1, 5, 22))

# 3. Styling
# Row 1: Main Title
requests.append(req_format(0, 1, 0, 22, bg_color=c_navy_dark, text_color=c_text_white, bold=True, font_size=13, halign="CENTER"))

# Row 2: Nav Bar
requests.append(req_format(1, 2, 0, 22, bg_color=c_navy_light, text_color=color_rgb(148, 163, 184), bold=True, font_size=9, halign="CENTER"))
# Highlight active nav tab (G:J, cols 6 to 10)
requests.append(req_format(1, 2, 6, 10, bg_color=color_rgb(37, 99, 235), text_color=c_text_white, bold=True, font_size=9, halign="CENTER"))

# Row 3: KPI Labels
requests.append(req_format(2, 3, 0, 2, bg_color=c_card_bg, text_color=c_text_muted, bold=True, font_size=8, halign="CENTER"))
requests.append(req_format(2, 3, 2, 4, bg_color=c_blue_light, text_color=c_text_blue, bold=True, font_size=8, halign="CENTER"))
requests.append(req_format(2, 3, 4, 6, bg_color=c_amber_light, text_color=color_rgb(180, 83, 9), bold=True, font_size=8, halign="CENTER"))
requests.append(req_format(2, 3, 6, 9, bg_color=c_red_light, text_color=c_text_red, bold=True, font_size=8, halign="CENTER"))
requests.append(req_format(2, 3, 9, 12, bg_color=c_red_light, text_color=c_text_red, bold=True, font_size=8, halign="CENTER"))
requests.append(req_format(2, 3, 12, 22, bg_color=c_red_light, text_color=c_text_red, bold=True, font_size=8, halign="CENTER"))

# Row 4: KPI Values
requests.append(req_format(3, 4, 0, 2, bg_color=c_card_bg, text_color=c_text_dark, bold=True, font_size=13, halign="CENTER"))
requests.append(req_format(3, 4, 2, 4, bg_color=c_blue_light, text_color=c_text_blue, bold=True, font_size=13, halign="CENTER"))
requests.append(req_format(3, 4, 4, 6, bg_color=c_amber_light, text_color=color_rgb(180, 83, 9), bold=True, font_size=13, halign="CENTER"))
requests.append(req_format(3, 4, 6, 9, bg_color=c_red_light, text_color=c_text_red, bold=True, font_size=13, halign="CENTER"))
requests.append(req_format(3, 4, 9, 12, bg_color=c_red_light, text_color=c_text_red, bold=True, font_size=13, halign="CENTER"))
requests.append(req_format(3, 4, 12, 22, bg_color=c_red_light, text_color=c_text_red, bold=True, font_size=11, halign="CENTER"))

# Section Headers (Rows 6, 15, 29, 37, 53)
for r_idx in [5, 14, 28, 36, 52]:
    requests.append(req_format(r_idx, r_idx+1, 0, 22, bg_color=c_blue_light, text_color=c_text_dark, bold=True, font_size=11, halign="LEFT"))

# Table Headers (Rows 7, 16, 30, 38, 54)
for r_idx in [6, 15, 29, 37, 53]:
    requests.append(req_format(r_idx, r_idx+1, 0, 22, bg_color=c_table_hdr, text_color=c_text_white, bold=True, font_size=9, halign="CENTER", wrap=True))

# Leaderboard Data Rows (8-12)
requests.append(req_format(7, 12, 0, 22, bg_color=color_rgb(255, 255, 255), text_color=c_text_dark, bold=False, font_size=9, halign="LEFT"))
requests.append(req_format(7, 12, 0, 1, halign="CENTER"))
requests.append(req_format(7, 12, 2, 7, halign="CENTER"))
requests.append(req_format(7, 12, 7, 8, halign="RIGHT", bold=True, text_color=c_text_red))
requests.append(req_format(7, 12, 8, 9, halign="CENTER", bold=True))
# Leaderboard Total (Row 13)
requests.append(req_format(12, 13, 0, 22, bg_color=c_card_bg, text_color=c_text_dark, bold=True, font_size=9, halign="LEFT"))
requests.append(req_format(12, 13, 2, 7, halign="CENTER", bold=True))
requests.append(req_format(12, 13, 7, 8, halign="RIGHT", bold=True, text_color=c_text_red))

# Matrix Data Rows (17-26)
requests.append(req_format(16, 26, 0, 22, bg_color=color_rgb(255, 255, 255), text_color=c_text_dark, bold=False, font_size=9, halign="LEFT"))
requests.append(req_format(16, 26, 0, 1, halign="CENTER"))
requests.append(req_format(16, 26, 3, 4, halign="RIGHT", bold=True))
requests.append(req_format(16, 26, 4, 6, halign="CENTER"))
requests.append(req_format(16, 26, 6, 19, halign="CENTER"))
requests.append(req_format(16, 26, 19, 20, halign="CENTER", bold=True))
requests.append(req_format(16, 26, 20, 21, halign="CENTER", bold=True))
requests.append(req_format(16, 26, 21, 22, halign="RIGHT", bold=True, text_color=c_text_red))
# Matrix Total (Row 27)
requests.append(req_format(26, 27, 0, 22, bg_color=c_card_bg, text_color=c_text_dark, bold=True, font_size=9, halign="LEFT"))
requests.append(req_format(26, 27, 3, 4, halign="RIGHT", bold=True))
requests.append(req_format(26, 27, 19, 21, halign="CENTER", bold=True))
requests.append(req_format(26, 27, 21, 22, halign="RIGHT", bold=True, text_color=c_text_red))

# Quotes Data Rows (31-35)
requests.append(req_format(30, 35, 0, 22, bg_color=color_rgb(255, 255, 255), text_color=c_text_dark, bold=False, font_size=9, halign="LEFT", wrap=True))
requests.append(req_format(30, 35, 0, 1, halign="CENTER"))
requests.append(req_format(30, 35, 3, 4, halign="RIGHT", bold=True, text_color=c_text_red))
requests.append(req_format(30, 35, 4, 6, bold=True, text_color=c_text_red))
requests.append(req_format(30, 35, 13, 22, bg_color=c_amber_light, bold=True))

# 13 Criteria Reference Rows (39-51)
requests.append(req_format(38, 51, 0, 22, bg_color=color_rgb(255, 255, 255), text_color=c_text_dark, bold=False, font_size=9, halign="LEFT", wrap=True))
requests.append(req_format(38, 51, 0, 1, halign="CENTER"))
requests.append(req_format(38, 51, 1, 2, bold=True))
requests.append(req_format(38, 51, 2, 4, halign="CENTER", bold=True))
requests.append(req_format(38, 51, 7, 12, text_color=c_text_red))
requests.append(req_format(38, 51, 12, 22, bold=True))

# ROI Rows (55-60)
requests.append(req_format(54, 60, 0, 22, bg_color=color_rgb(255, 255, 255), text_color=c_text_dark, bold=False, font_size=9, halign="LEFT", wrap=True))
requests.append(req_format(54, 60, 0, 3, bold=True))
requests.append(req_format(54, 60, 3, 5, halign="CENTER", bold=True, text_color=c_text_green))

# Send batch update
print(f"Sending batchUpdate with {len(requests)} format requests...")
res = sh.batch_update({"requests": requests})
print("Successfully applied styling to Google Sheet!")
