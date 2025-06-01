from datetime import datetime, timedelta

def get_all_days(year, month):
    first_day = datetime(year, month, 1).date()
    if month == 12:
        next_month = datetime(year + 1, 1, 1).date()
    else:
        next_month = datetime(year, month + 1, 1).date()
    delta = next_month - first_day
    return [first_day + timedelta(days=i) for i in range(delta.days)]

def parse_flexible_date(date_str):
    # If date_str is already a datetime object, convert it to a string
    if isinstance(date_str, datetime):
        return date_str

    # If it's a string, try to parse it
    formats = [
        "%d-%b-%y", "%d/%m/%Y", "%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d",
        "%d %B %Y", "%b %d, %Y", "%d.%m.%Y", "%B %d, %Y"
    ]
    for fmt in formats:
        try:
            cleaned = date_str.strip().replace('\ufeff', '')
            return datetime.strptime(cleaned, fmt)  # Return as datetime, not date
        except ValueError:
            continue
    raise ValueError(f"Unsupported date format: '{date_str}'")

def convert_to_12h(time_str):
    """
    Converts time to 12-hour format like '01:45 pm' regardless of input format.
    Accepts both 24-hour ('13:45') and 12-hour ('01:45 pm') formats.
    """
    time_str = time_str.strip().lower()

    # Try parsing as 24-hour time first
    try:
        t = datetime.strptime(time_str, "%H:%M")
        return t.strftime("%I:%M %p").lower()
    except ValueError:
        pass

    # Try parsing as 12-hour time (if it's already formatted)
    try:
        t = datetime.strptime(time_str, "%I:%M %p")
        return t.strftime("%I:%M %p").lower()
    except ValueError:
        raise ValueError(f"Unsupported time format: '{time_str}'")

from tkinter import messagebox
import os
import csv
import sys
import subprocess

def open_file_location(filepath):
    folder = os.path.dirname(filepath)
    if sys.platform == "win32":
        os.startfile(folder)
    elif sys.platform == "darwin":
        subprocess.run(["open", folder])
    else:
        subprocess.run(["xdg-open", folder])

def generate_template():
    filename = "logbook_template.csv"
    filepath = os.path.abspath(filename)
    example_data = [
        ["date", "activity", "clockin", "clockout"],
        ["2025-06-01", "Studying Python", "08:00", "10:00"],
        ["2025-06-02", "Workshop attendance", "13:00", "15:00"],
        ["2025-06-03", "OFF", "OFF", "OFF"]
    ]

    try:
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerows(example_data)

        result = messagebox.askyesno("Template Generated ✅", f"CSV template created:\n{filepath}\n\nOpen file location?")
        if result:
            open_file_location(filepath)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate template:\n{e}")