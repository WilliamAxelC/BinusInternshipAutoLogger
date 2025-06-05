import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext, ttk
import threading
import asyncio
import requests
import os
from playwright.async_api import async_playwright
import sys
import json
from collections import defaultdict
from utility import get_all_days, parse_flexible_date, convert_to_12h, generate_template
from datetime import datetime
import argparse
import pandas as pd

cookie = None
month_header_dict = None
last_browser = None  # Global variable to hold the previous browser instance

# Parse command-line arguments
parser = argparse.ArgumentParser(description="Run BINUS logbook automation.")
parser.add_argument("--debug", action="store_true", help="Enable debugging mode.")
args = parser.parse_args()

debugging_mode = args.debug

def get_header_id_for_date(month_header_dict, date_str):
    # If date_str is already a datetime object, use it directly
    if isinstance(date_str, datetime):
        date_obj = date_str
    else:
        # Otherwise, assume it's a string and parse it
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")

    # Extract the month from the date
    month_name = date_obj.strftime("%B")  # Get full month name (e.g., "January")
    month_name = month_name.upper()

    # Retrieve the corresponding headerID for the month
    header_id = month_header_dict.get(month_name)
    if not header_id:
        log_message(f"❌ No headerID found for month: {month_name}")
        return None
    return header_id


# For PyInstaller -- resolve the path correctly
def get_runtime_browser_path():
    if getattr(sys, 'frozen', False):
        # If running in a PyInstaller bundle
        return os.path.join(sys._MEIPASS, 'playwright', 'driver', 'package', '.local-browsers')
    else:
        # Running normally (dev environment)
        return None

browser_path = get_runtime_browser_path()
if browser_path:
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = browser_path

async def launch_and_get_cookie_and_header_async(email, password):
    log_message("Starting Playwright task...")
    global last_browser

    try:
        async with async_playwright() as p:
                # Close previous browser if it exists
            if last_browser:
                log_message("🔁 Closing previous browser instance...")
                try:
                    await last_browser.close()
                    log_message("✅ Previous browser closed.")
                except Exception as e:
                    log_message(f"⚠️ Error closing previous browser: {e}")
                last_browser = None
            log_message("   Launching Headless browser...")

            browser = await p.webkit.launch(headless=True)
            last_browser = browser  # Store the reference to this new browser
            context = await browser.new_context(viewport={"width": 1920, "height": 1080})
            page = await context.new_page()

            # Step 1: Go to the enrichment site
            log_message("   Navigating to the enrichment site...")
            await page.goto("https://enrichment.apps.binus.ac.id/Dashboard")
            await page.wait_for_load_state("domcontentloaded")
            log_message("   Enrichment site loaded.")

            # Step 2: Click the "Sign in with Microsoft" button
            log_message("   Clicking 'Sign in with Microsoft' button...")
            await page.click("button#btnLogin")
            log_message("   'Sign in with Microsoft' button clicked.")

            # Step 3: Insert Email and Click "Next"
            log_message("   Entering email...")
            await page.fill('input#i0116', email)
            log_message(f"  Email entered: {email}")
            await page.click('input#idSIButton9')  # Click "Next"
            log_message("   Clicked 'Next' after entering email.")

            # Step 4: Insert Password and Click "Sign in"
            log_message("   Entering password...")
            await page.fill('input#i0118', password)
            log_message("   Password entered.")
            await page.click('input#idSIButton9')  # Click "Sign in"
            log_message("   Clicked 'Sign in' after entering password.")

            # Step 5: Handle "No" button click
            log_message("   Checking for 'No' button...")
            try:
                await page.click('input#idBtn_Back')  # If "No" button exists, click it
                log_message("   Clicked 'No' button.")
            except:
                log_message("   'No' button not found or not clickable.")

            log_message("   Waiting for 'Go to Activity Enrichment Apps' button...")

            button_selector = 'a.button-orange[href*="/SSOToActivity"]'
            max_attempts = 2

            for attempt in range(1, max_attempts + 1):
                try:
                    log_message(f"   Attempt {attempt} to find and click the button...")
                    await page.wait_for_selector(button_selector, timeout=10000, state='visible')
                    log_message("   Button found, scrolling into view and clicking...")

                    # Scroll into view and click
                    await page.eval_on_selector(button_selector, "el => el.scrollIntoView({ behavior: 'smooth', block: 'center' })")
                    await page.wait_for_timeout(500)
                    await page.click(button_selector, force=True)
                    log_message("   Clicked 'Go to Activity Enrichment Apps' button.")
                    break  # Exit the loop if successful

                except Exception as e:
                    log_message(f"   ⚠️ Attempt {attempt} failed: {e}")
                    if attempt < max_attempts:
                        log_message("   Refreshing page and retrying...")
                        await page.reload()
                        await page.wait_for_timeout(2000)  # Optional wait to allow page to settle
                    else:
                        log_message("   ❌ Final attempt failed.")
                        log_message("❌❌❌ Something went wrong! Please click on 'Fetch Cookies & Header ID' again. 🛠️🔄")
                        return None, None



            # Step 7: Account selection
            log_message("   Waiting for account selection tile...")
            await page.wait_for_selector('//*[@id="tilesHolder"]/div[1]/div/div[1]/div/div[2]/div[1]', timeout=5000)
            log_message("   Account selection tile found, clicking it...")
            await page.click('//*[@id="tilesHolder"]/div[1]/div/div[1]/div/div[2]/div[1]')

            # Wait for page load and get LogBookHeaderID
            log_message("   Waiting for LogBook button...")
            await page.wait_for_selector("#btnLogBook")
            log_message("   LogBook button found, clicking it...")
            await page.click("#btnLogBook")

            # Step 6: Wait for month tabs to appear
            log_message("   Waiting for month tabs...")
            await page.wait_for_selector('#monthTab')

            # Extract the months and onclick header ids
            month_header_dict = {}
            
            # Wait for the overlay to disappear (use a timeout of 5 seconds)
            log_message("   Waiting for overlay to disappear...")
            await page.wait_for_selector('.fancybox-overlay', state='hidden', timeout=5000)

            # Loop through each month <a> tag and click it to extract the onclick id
            month_elements = await page.query_selector_all('#monthTab li a')
            log_message(f"   Found {len(month_elements)} month elements.")
            
            for month_element in month_elements:
                # Extract month name and strip extra spaces
                month_name = await month_element.inner_text()
                month_name = month_name.strip()

                month_name = month_name.replace(' ●', '').strip()  # This will remove the circle character


                # Click on the month element
                log_message(f"   Clicking on {month_name}...")
                await month_element.click()

                # Extract the onclick header ID after clicking
                header_id = await page.evaluate('''(el) => {
                    const onclick = el.getAttribute('onclick');
                    return onclick ? onclick.split("'")[1] : null;
                }''', month_element)

                if header_id:
                    month_header_dict[month_name] = header_id
                    log_message(f"   Mapped {month_name} to {header_id}")

            log_message(f"   Extracted month-header pairs: {month_header_dict}")

            if not month_header_dict:
                log_message("   No LogBookHeaderID received.")

            # Extract cookies
            log_message("   Extracting cookies...")
            cookies = await context.cookies()
            cookie_header = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
            log_message("   Cookies extracted.")

            await browser.close()
            log_message("   Browser closed.")
            return cookie_header, month_header_dict

    except Exception as e:
        log_message(f"  Error in Playwright task: {e}")
        raise

# File to store credentials
DATA_FILE = "data.json"

# Function to save credentials
def save_credentials(email, password):
    data = {
        "email": email,
        "password": password
    }
    with open(DATA_FILE, 'w') as file:
        json.dump(data, file)

def load_json():
    try:
        with open('data.json', 'r') as f:
            data = json.load(f)

            # Restore CSV path if available
            saved_path = data.get('csv_path')
            if saved_path and entry_file:
                entry_file.delete(0, tk.END)
                entry_file.insert(0, saved_path)

            # Return credentials if available
            return data.get("email"), data.get("password")

    except (FileNotFoundError, json.JSONDecodeError):
        return None, None

# Step 2: Tkinter Dialog to get credentials
class CustomDialog(simpledialog.Dialog):
    def __init__(self, parent, title, prompt1, prompt2, remember_me=False):
        self.prompt1 = prompt1
        self.prompt2 = prompt2
        self.remember_me = remember_me
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text=self.prompt1).grid(row=0, column=0, padx=5, pady=5)
        self.email_entry = tk.Entry(master, width=40)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(master, text=self.prompt2).grid(row=1, column=0, padx=5, pady=5)
        self.password_entry = tk.Entry(master, width=40, show="*")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        self.remember_me_var = tk.BooleanVar()
        self.remember_me_check = tk.Checkbutton(master, text="Remember Me", variable=self.remember_me_var)
        self.remember_me_check.grid(row=2, column=0, columnspan=2, pady=5)

        return self.email_entry  # Focus on email entry initially

    def apply(self):
        self.result = (self.email_entry.get(), self.password_entry.get(), self.remember_me_var.get())

def ask_for_credentials(parent, title="Login", prompt1="Enter your email:", prompt2="Enter your password:"):
    # Ensure dialog has time to finish before proceeding
    dialog = CustomDialog(parent, title, prompt1, prompt2)
    return dialog.result

def fetch_existing_entries(header_id, cookie):
    try:
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Cookie": cookie,
            "X-Requested-With": "XMLHttpRequest"
        }
        response = requests.post(
            "https://activity-enrichment.apps.binus.ac.id/LogBook/GetLogBook",
            headers=headers,
            data={"logBookHeaderID": header_id}
        )

        response.raise_for_status()
        data = response.json().get("data", [])
        return {entry["date"][:10]: entry["id"] for entry in data}
    except Exception as e:
        log_message(f"❌ Failed to fetch existing entries: {e}")
        return {}


def log_message(message):
    if debugging_mode == True:
        prefix = "[DEBUG]"
    else:
        prefix = "[INFO]"

    message = f"{prefix} {message}"

    if "✅" in message:
        color = "green"
    elif "❌" in message or "⚠️" in message:
        color = "red"
    elif "[DEBUG]" in message:
        color = "blue"
    else:
        color = "black"
        
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    full_message = f"{timestamp} {message}"
    print(full_message)

    with open("debug_log.txt", "a", encoding="utf-8") as f:
        f.write(full_message + "\n")
    
    output_box.insert(tk.END, message + "\n", color)
    output_box.tag_config("green", foreground="green")
    output_box.tag_config("red", foreground="red")
    output_box.tag_config("blue", foreground="blue")
    output_box.tag_config("black", foreground="black")
    output_box.see(tk.END)

def process_logbook(csv_path, cookie, edit=False, month_header_dict=None):
    # if debugging_mode == True:
    #     pdb.set_trace()
    log_message(f"Edit mode: {edit}")
    handled_dates = set()
    invalid_rows = []
    active_days = 0
    month_entries = defaultdict(list)

    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
    except Exception as e:
        log_message(f"❌ Failed to read CSV: {e}")
        return

    # Normalize column names
    df.columns = [col.strip().lower() for col in df.columns]
    required_cols = {"date", "activity", "clockin", "clockout"}

    if not required_cols.issubset(df.columns):
        log_message("❌ CSV missing required headers: date, activity, clockin, clockout")
        return

    for idx, row in df.iterrows():
        try:
            raw_date = str(row["date"]).strip()
            activity = str(row["activity"]).strip()
            clockin = str(row["clockin"]).strip()
            clockout = str(row["clockout"]).strip()

            if activity.lower() == 'off':
                clockin = clockout = 'off'

            if not raw_date or not activity:
                continue  # Skip empty or incomplete rows

            parsed_date = parse_flexible_date(raw_date)
            header_id = get_header_id_for_date(month_header_dict, parsed_date)

            
            clockin_12h = convert_to_12h(clockin)
            clockout_12h = convert_to_12h(clockout)

            is_off = parsed_date.weekday() >= 5 or activity.lower() == "off"
            date_key = parsed_date.strftime("%Y-%m-%d")

            if not is_off:
                active_days += 1
                handled_dates.add(parsed_date.date())

                entry = {
                    "model[ID]": None,
                    "model[LogBookHeaderID]": header_id,
                    "model[ClockIn]": clockin_12h,
                    "model[ClockOut]": clockout_12h,
                    "model[Date]": parsed_date.strftime("%Y-%m-%dT00:00:00"),
                    "model[Activity]": activity,
                    "model[Description]": activity
                }

                if edit:
                    existing_entries = fetch_existing_entries(header_id, cookie)
                    entry["model[ID]"] = existing_entries.get(date_key, "00000000-0000-0000-0000-000000000000")

                month_entries[parsed_date.month].append(entry)

        except Exception as e:
            invalid_rows.append(f"{row.to_dict()} ({e})")

    if invalid_rows:
        log_message("❌ Process aborted due to invalid row(s):")
        for err in invalid_rows:
            log_message(f"  - {err}")
        return

    if active_days < 5:
        log_message("❌ Fewer than 5 Active (non-off) days! Check your CSV file.")
        return

    # Submit active entries
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Cookie": cookie
    }
    url = "https://activity-enrichment.apps.binus.ac.id/LogBook/StudentSave"

    for month, entries in month_entries.items():
        for entry in entries:
            try:
                if not debugging_mode:
                    response = requests.post(url, headers=headers, data=entry)
                else:
                    class MockResponse:
                        ok = True
                        status_code = 200
                        text = 'Debug mode: simulated response'
                    response = MockResponse()

                date_display = entry['model[Date]'][:10]
                if response.ok:
                    log_message(f"✅ {date_display} submitted successfully.")
                else:
                    log_message(f"❌ Failed {date_display} - {response.status_code}: {response.text}")
            except Exception as e:
                log_message(f"❌ Network error on {entry['model[Date]'][:10]}: {e}")

    # Submit OFF entries for missing weekdays
    if not handled_dates:
        log_message("⚠️ No dates found in CSV to infer OFF days.")
        return

    month_year_pairs = set((d.year, d.month) for d in handled_dates)

    print(month_year_pairs)

    for year, month in month_year_pairs:
        for day in get_all_days(year, month):
            if day not in handled_dates:
                header_id = get_header_id_for_date(month_header_dict, day.isoformat())
                date_str = day.strftime("%Y-%m-%dT00:00:00")
                entry_id = None

                if edit:
                    existing = fetch_existing_entries(header_id, cookie)
                    entry_id = existing.get(day.isoformat(), "00000000-0000-0000-0000-000000000000")

                payload = {
                    "model[ID]": entry_id,
                    "model[LogBookHeaderID]": header_id,
                    "model[Date]": date_str,
                    "model[Activity]": "OFF",
                    "model[ClockIn]": "OFF",
                    "model[ClockOut]": "OFF",
                    "model[Description]": "OFF",
                    "model[flagjulyactive]": "false"
                }

                try:
                    if not debugging_mode:
                        response = requests.post(url, data=payload, headers=headers)
                    else:
                        class MockResponse:
                            ok = True
                            status_code = 200
                            text = 'Debug mode: simulated response'
                        response = MockResponse()

                    if response.ok:
                        log_message(f"🟡 OFF submitted for {day}")
                    else:
                        log_message(f"❌ Failed OFF for {day}: {response.status_code} - {response.text}")
                except Exception as e:
                    log_message(f"❌ Network error submitting OFF for {day}: {e}")

# === GUI SETUP ===
def browse_file():
    path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, path)

        # Save to data.json
        try:
            with open('data.json', 'r') as f:
                data = json.load(f)
        except FileNotFoundError:
            data = {}

        data['csv_path'] = path

        with open('data.json', 'w') as f:
            json.dump(data, f, indent=4)



def start_process():
    output_box.delete(1.0, tk.END)
    file_path = entry_file.get()
    # clock_in = entry_clockin.get().strip() or "09:00 am"
    # clock_out = entry_clockout.get().strip() or "06:00 pm"
    is_edit = edit_mode.get()

    if not (file_path):
        messagebox.showerror("Missing Info", "Please fill in all fields.")
        return
    
    if not (cookie and month_header_dict):
        if not (month_header_dict or cookie):
            messagebox.showerror("Missing Cookie and Header ID", "Please click on 'Fetch Cookie & Header ID'")
        elif cookie:
            messagebox.showerror("Missing Header ID", "Please click on 'Fetch Cookie & Header ID'")
        else:
            messagebox.showerror("Missing Cookie", "Please click on 'Fetch Cookie & Header ID'")

        return
    
    

    thread = threading.Thread(target=process_logbook, args=(file_path, cookie, is_edit, month_header_dict))
    thread.start()

def get_cookie_and_header():
    def on_credentials_gathered(email, password, remember_me):
        if not email or not password:
            messagebox.showerror("Credentials Missing", "Email and password are required.")
            return

        log_message("Launching browser for cookie/header retrieval...")

        # Save credentials if "Remember Me" is checked
        if remember_me:
            save_credentials(email, password)
        
        # Create a separate function to run the async task in a new thread
        def fetch_data():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                # Call the async function inside the thread
                global cookie, month_header_dict
                cookie, month_header_dict = loop.run_until_complete(launch_and_get_cookie_and_header_async(email, password))
                # Update the GUI fields with the fetched data
                if cookie is not None and month_header_dict is not None:
                    root.after(0, update_gui_fields)

            except Exception as e:
                log_message(f"❌ Failed to get cookie and header: {e}")
                log_message("❌❌❌ Something went wrong! Please click on 'Fetch Cookies & Header ID' again. 🛠️🔄")

        
        # Start the fetch_data function in a new thread
        threading.Thread(target=fetch_data, daemon=True).start()

    def update_gui_fields():
        log_message("✅ Cookie and Header ID retrieved.")

    # Try to load saved data first
    email, password = load_json()

    if email and password:
        log_message("✅ Using saved credentials.")
        on_credentials_gathered(email, password, remember_me=True)
    else:
        # Ask user if no saved credentials
        dialog = CustomDialog(root, "Login", "Enter your email:", "Enter your password:")
        if dialog.result:
            email, password, remember_me = dialog.result
            on_credentials_gathered(email, password, remember_me)

# === Help Popup ===
def show_help_popup():
    help_win = tk.Toplevel(root)
    help_win.title("How to Use the BINUS Logbook Submitter")
    help_win.geometry("500x400")

    instructions = (
        "📘 How to Use:\n\n"
        "1. 'Generate CSV Template' button\n"
        "    - Creates a blank CSV file to fill in your daily activities.\n\n"
        "2. Fill the CSV\n"
        "    - Enter your logbook entries. Include: date, activity, clock-in, and clock-out.\n\n"
        "3. 'Fetch Cookie & Header'\n"
        "    - Opens browser, logs in to BINUS Enrichment, and grabs session info needed for submission.\n\n"
        "4. 'Edit existing entries'\n"
        "    - Turn this ON if you want to update already-submitted entries.\n\n"
        "5. 'Submit Logbook'\n"
        "    - Submits your entries.\n"
    )

    text_box = tk.Text(help_win, wrap=tk.WORD, width=60, height=20)
    text_box.insert(tk.END, instructions)
    text_box.config(state='disabled')
    text_box.pack(padx=10, pady=10)

    tk.Button(help_win, text="Close", command=help_win.destroy).pack(pady=(0, 10))


# === Build GUI ===
root = tk.Tk()
root.title("BINUS Logbook Submitter")

# Get path relative to the executable (for PyInstaller compatibility)
base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
icon_path = os.path.join(base_path, 'logo.ico')

root.iconbitmap(icon_path)

edit_mode = tk.BooleanVar()

# === Row 0: File Selection ===
tk.Label(root, text="Logbook CSV File:").grid(row=0, column=0, sticky="e")
entry_file = tk.Entry(root, width=80)
entry_file.grid(row=0, column=1, sticky="w")
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, sticky="w")

# === Row 1: Generate CSV Template ===
tk.Button(root, text="Generate CSV Template", command=generate_template).grid(row=1, column=1, pady=4, sticky="w")

# === Row 2: Fetch Cookie and Header ID ===
tk.Button(root, text="Fetch Cookie & Header ID", command=get_cookie_and_header, bg="blue", fg="white").grid(row=2, column=1, pady=4, sticky="w")

# === Row 3: Edit Mode Checkbox ===
tk.Checkbutton(root, text="Edit existing entries (turn this ON to update)", variable=edit_mode).grid(row=3, column=1, pady=4, sticky="w")

# === Row 4: Submit Button ===
tk.Button(root, text="Submit Logbook", command=start_process, bg="green", fg="white").grid(row=4, column=1, pady=4, sticky="w")

# === Row 5: Help Button ===
tk.Button(root, text="❓ How to Use", command=show_help_popup).grid(row=5, column=1, pady=4, sticky="w")

# === Row 6: Output Box ===
output_box = scrolledtext.ScrolledText(root, width=80, height=20, state='normal')
output_box.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

# === Load Configs & Start GUI ===
load_json()
root.mainloop()
