import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
import threading
import asyncio
import requests
import csv
import os
from datetime import datetime, timedelta
from playwright.async_api import async_playwright
import sys
import json

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

    try:
        async with async_playwright() as p:
            log_message("   Launching Headless browser...")

            browser = await p.webkit.launch(headless=True)
            context = await browser.new_context()
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

            log_message("   Waiting for login confirmation button to appear...")

            try:
            # Wait up to 3 minutes for the button with href containing '/SSOToActivity'
                await page.wait_for_selector('a.button-orange[href*="/SSOToActivity"]', timeout=180000)
                log_message("   'Go to Activity Enrichment Apps' button found, scrolling into view...")

                # Scroll it into view
                await page.evaluate('''() => {
                    const el = document.querySelector('a.button-orange[href*="/SSOToActivity"]');
                    if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                }''')

                await page.wait_for_timeout(1000)  # allow time for scroll

                log_message("   Clicking the 'Go to Activity Enrichment Apps' button...")
                await page.click('a.button-orange[href*="/SSOToActivity"]')
            except:
                log_message("   ❌ 'Go to Activity Enrichment Apps' button not found within 3 minutes.")
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

            # Capture GetLogBook response
            header_id = None

            async def handle_response(response):
                nonlocal header_id
                if "GetLogBook" in response.url:
                    log_message("   GetLogBook response received.")
                    try:
                        # Await response.json() properly to get the data
                        json_data = await response.json()
                        if json_data["data"]:
                            header_id = json_data["data"][0]["logBookHeaderID"]
                            log_message(f"  LogBookHeaderID retrieved: {header_id}")
                    except Exception as e:
                        log_message(f"  Error parsing response: {e}")

            page.on("response", handle_response)
            await page.wait_for_timeout(3000)  # Wait for the response

            if not header_id:
                log_message("   No LogBookHeaderID received.")

            # Extract cookies
            log_message("   Extracting cookies...")
            cookies = await context.cookies()
            cookie_header = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
            log_message("   Cookies extracted.")

            await browser.close()
            log_message("   Browser closed.")
            return cookie_header, header_id

    except Exception as e:
        log_message(f"  Error in Playwright task: {e}")
        log_message("❌❌❌ Something went wrong! Please re-run the process. 🛠️🔄")
        raise


# File to store credentials
CREDENTIALS_FILE = "data.json"

# Function to save credentials
def save_credentials(email, password):
    data = {
        "email": email,
        "password": password
    }
    with open(CREDENTIALS_FILE, 'w') as file:
        json.dump(data, file)

# Function to load credentials if available
def load_credentials():
    if os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE, 'r') as file:
            data = json.load(file)
            return data.get("email"), data.get("password")
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


def get_all_days(year, month):
    first_day = datetime(year, month, 1).date()
    if month == 12:
        next_month = datetime(year + 1, 1, 1).date()
    else:
        next_month = datetime(year, month + 1, 1).date()
    delta = next_month - first_day
    return [first_day + timedelta(days=i) for i in range(delta.days)]

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
    print(message)
    output_box.insert(tk.END, message + '\n')
    output_box.see(tk.END)

def parse_flexible_date(date_str):
    formats = [
        "%d-%b-%y", "%d/%m/%Y", "%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d",
        "%d %B %Y", "%b %d, %Y", "%d.%m.%Y", "%B %d, %Y"
    ]
    for fmt in formats:
        try:
            cleaned = date_str.strip().replace('\ufeff', '')
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Unsupported date format: '{date_str}'")


def process_logbook(csv_path, header_id, cookie, clock_in, clock_out, edit=False):
    log_message(f"Edit mode: {edit}")
    existing_ids = fetch_existing_entries(header_id, cookie) if edit else {}
    logbook_entries = []
    handled_dates = set()
    invalid_dates = []
    active_days = 0

    try:
        with open(csv_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=',')
            rows = list(reader)

            if len(rows) < 2:
                log_message("❌ CSV file does not contain enough rows.")
                return

            date_strs = rows[0]
            activities = rows[1]

            if not any(cell.strip() for cell in date_strs) or not any(cell.strip() for cell in activities):
                log_message("❌ CSV appears to be empty or missing required data.")
                return

            for date_str, activity in zip(date_strs, activities):
                activity_str = activity.strip()
                if activity_str:
                    try:
                        parsed_date = parse_flexible_date(date_str)
                        iso_date = parsed_date.strftime('%Y-%m-%dT00:00:00')

                        is_off = parsed_date.weekday() >= 5 or activity_str.strip().lower() == "off"

                        if not is_off:
                            active_days += 1
                            handled_dates.add(parsed_date)
                            logbook_entries.append({
                                "model[ID]": existing_ids.get(parsed_date.isoformat(), "00000000-0000-0000-0000-000000000000"),
                                "model[LogBookHeaderID]": header_id,
                                "model[ClockIn]": clock_in,
                                "model[ClockOut]": clock_out,
                                "model[flagjulyactive]": "false",
                                'model[Date]': iso_date,
                                'model[Activity]': activity_str,
                                'model[Description]': activity_str
                            })


                    except ValueError as e:
                        invalid_dates.append(f"{date_str} ({e})")

    except Exception as e:
        log_message(f"❌ Failed to read CSV: {e}")
        return

    if invalid_dates:
        log_message("❌ Process aborted due to invalid date(s):")
        for err in invalid_dates:
            log_message(f"  - {err}")
        return

    if active_days < 10:
        log_message("❌ Fewer than 10 Active (non-off) days! Check your CSV file.")
        return

    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Cookie": cookie
    }

    url = "https://activity-enrichment.apps.binus.ac.id/LogBook/StudentSave"

    for data in logbook_entries:
        try:
            response = requests.post(url, headers=headers, data=data)
            date_display = data['model[Date]'][:10]
            if response.ok:
                log_message(f"✅ {date_display} submitted successfully.")
            else:
                log_message(f"❌ Failed {date_display} - {response.status_code}: {response.text}")
        except Exception as e:
            log_message(f"❌ Network error on {data['model[Date]'][:10]}: {e}")
    
    # Add "OFF" entries
    if not handled_dates:
        log_message("⚠️ No dates found in CSV to infer OFF days.")
        return
    
    print(handled_dates)

    # Assume one month only
    first_date = min(handled_dates)
    year, month = first_date.year, first_date.month
    all_dates = get_all_days(year, month)

    print(all_dates)


    # Add "OFF" entries
    all_dates = get_all_days(year, month)
    for date in all_dates:
        if date not in handled_dates:
            payload = {
                "model[ID]": existing_ids.get(date.isoformat(), "00000000-0000-0000-0000-000000000000"),
                "model[LogBookHeaderID]": header_id,
                "model[Date]": date.isoformat(),
                "model[Activity]": "OFF",
                "model[ClockIn]": "OFF",
                "model[ClockOut]": "OFF",
                "model[Description]": "OFF",
                "model[flagjulyactive]": "false"
            }

            try:
                response = requests.post(url, data=payload, headers=headers)
                if response.status_code == 200:
                    log_message(f"🟡 OFF submitted for {date}.")
                else:
                    log_message(f"❌ Failed OFF for {date}: {response.status_code} - {response.text}")
            except Exception as e:
                log_message(f"❌ Network error submitting OFF for {date}: {e}")

    log_message("✅ Process complete.")


# === GUI SETUP ===
def browse_file():
    path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    entry_file.delete(0, tk.END)
    entry_file.insert(0, path)

def start_process():
    output_box.delete(1.0, tk.END)
    file_path = entry_file.get()
    cookie = entry_cookie.get()
    header_id = entry_header.get()
    clock_in = entry_clockin.get().strip() or "09:00 am"
    clock_out = entry_clockout.get().strip() or "06:00 pm"
    is_edit = edit_mode.get()

    if not (file_path and cookie and header_id):
        messagebox.showerror("Missing Info", "Please fill in all fields.")
        return

    thread = threading.Thread(target=process_logbook, args=(file_path, header_id, cookie, clock_in, clock_out, is_edit))
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
                cookie, header = loop.run_until_complete(launch_and_get_cookie_and_header_async(email, password))
                # Update the GUI fields with the fetched data
                root.after(0, update_gui_fields, cookie, header)
            except Exception as e:
                log_message(f"❌ Failed to get cookie and header: {e}")
        
        # Start the fetch_data function in a new thread
        threading.Thread(target=fetch_data, daemon=True).start()

    def update_gui_fields(cookie, header):
        # Update the Tkinter fields after the task completes
        entry_cookie.delete(0, tk.END)
        entry_cookie.insert(0, cookie)
        entry_header.delete(0, tk.END)
        entry_header.insert(0, header or "")
        log_message("✅ Cookie and Header ID retrieved.")

    # Try to load saved credentials first
    email, password = load_credentials()

    if email and password:
        log_message("✅ Using saved credentials.")
        on_credentials_gathered(email, password, remember_me=True)
    else:
        # Ask user if no saved credentials
        dialog = CustomDialog(root, "Login", "Enter your email:", "Enter your password:")
        if dialog.result:
            email, password, remember_me = dialog.result
            on_credentials_gathered(email, password, remember_me)


# === Build GUI ===
root = tk.Tk()
root.title("BINUS Logbook Submitter")
edit_mode = tk.BooleanVar()

tk.Label(root, text="Logbook CSV File:").grid(row=0, column=0, sticky="e")
entry_file = tk.Entry(root, width=50)
entry_file.grid(row=0, column=1, sticky="w")
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, sticky="w")

tk.Button(root, text="Auto Fetch Cookie & Header ID", command=get_cookie_and_header, bg="blue", fg="white").grid(row=1, column=0, columnspan=3, pady=(2, 0))

tk.Label(root, text="Header ID:").grid(row=2, column=0, sticky="e")
entry_header = tk.Entry(root, width=50)
entry_header.grid(row=2, column=1, columnspan=2, sticky="w")

tk.Label(root, text="Cookie:").grid(row=3, column=0, sticky="e")
entry_cookie = tk.Entry(root, width=50)
entry_cookie.grid(row=3, column=1, columnspan=2, sticky="w")

tk.Label(root, text="Clock In (e.g. 09:00 am):").grid(row=4, column=0, sticky="e")
entry_clockin = tk.Entry(root, width=15)
entry_clockin.insert(0, "09:00 am")
entry_clockin.grid(row=4, column=1, sticky="w")

tk.Label(root, text="Clock Out (e.g. 06:00 pm):").grid(row=5, column=0, sticky="e")
entry_clockout = tk.Entry(root, width=15)
entry_clockout.insert(0, "06:00 pm")
entry_clockout.grid(row=5, column=1, sticky="w")

tk.Checkbutton(root, text="Edit existing entries", variable=edit_mode).grid(row=6, column=0, columnspan=3)

tk.Button(root, text="Submit Logbook", command=start_process, bg="green", fg="white").grid(row=7, column=0, columnspan=3, pady=10)

output_box = scrolledtext.ScrolledText(root, width=80, height=20, state='normal')
output_box.grid(row=8, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
