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
from collections import defaultdict

cookie = None
month_header_dict = None
last_browser = None  # Global variable to hold the previous browser instance
debugging_mode = os.environ.get("DEBUG_MODE", "false").lower() == "true"

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

            log_message("   Waiting for 'Go to Activity Enrichment Apps' button...")

            try:
                # Wait up to 3 minutes for the button to be visible
                button_selector = 'a.button-orange[href*="/SSOToActivity"]'
                await page.wait_for_selector(button_selector, timeout=20000, state='visible')
                log_message("   Button found, scrolling into view and clicking...")

                # Scroll the button into view properly
                await page.eval_on_selector(button_selector, "el => el.scrollIntoView({ behavior: 'smooth', block: 'center' })")
                await page.wait_for_timeout(500)  # Allow time for scroll animation

                # Click the button
                await page.click(button_selector, force=True)
                log_message("   Clicked 'Go to Activity Enrichment Apps' button.")
            except Exception as e:
                log_message(f"   ❌ Button click failed: {e}")
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

def process_logbook(csv_path, cookie, clock_in, clock_out, edit=False, month_header_dict=None):
    log_message(f"Edit mode: {edit}")
    handled_dates = set()
    invalid_dates = []
    active_days = 0

    month_entries = defaultdict(list)  # Track entries for each month

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
                        date_str = str(date_str)  # Ensure it's a string
                        parsed_date = parse_flexible_date(date_str)
                        iso_date = parsed_date.strftime('%Y-%m-%d')  # Format to YYYY-MM-DD

                        # Get LogBookHeaderID for the current date
                        header_id = get_header_id_for_date(month_header_dict, parsed_date)

                        is_off = parsed_date.weekday() >= 5 or activity_str.strip().lower() == "off"

                        if not is_off:
                            active_days += 1
                            handled_dates.add(parsed_date.date())
                            month_entries[parsed_date.month].append({
                                "model[ID]": None,  # Will be filled in later if edit=True
                                "model[LogBookHeaderID]": header_id,
                                "model[ClockIn]": clock_in,
                                "model[ClockOut]": clock_out,
                                'model[Date]': iso_date,  # Use only the date part
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

    # Process each month's data separately
    for month, entries in month_entries.items():
        header_id = None
        for data in entries:
            try:
                # Get header_id for the month if not done yet
                if header_id is None:
                    header_id = get_header_id_for_date(month_header_dict, data['model[Date]'])
                existing_ids = fetch_existing_entries(header_id, cookie) if edit else {}
                data['model[ID]'] = existing_ids.get(data['model[Date]'], "00000000-0000-0000-0000-000000000000")
                
                # Submit each entry
                if not debugging_mode:
                    response = requests.post(url, headers=headers, data=data)
                else:
                    # Create a mock response object with the .ok attribute
                    class MockResponse:
                        ok = True
                        status_code = 200
                        text = 'Debug mode: simulated response'

                    response = MockResponse()

                date_display = data['model[Date]'][:10]
                if response.ok:
                    log_message(f"✅ {date_display} submitted successfully.")
                else:
                    log_message(f"❌ Failed {date_display} - {response.status_code}: {response.text}")
            except Exception as e:
                log_message(f"❌ Network error on {data['model[Date]'][:10]}: {e}")

    # Add "OFF" entries for each month
    if len(handled_dates) == 0:
        log_message("⚠️ No dates found in CSV to infer OFF days.")
        return

    # Step 1: Extract month-year pairs from handled_dates
    month_year_pairs = set((date.year, date.month) for date in handled_dates)

    # for date in handled_dates:
    #     print(date)

    # Step 2: For each month-year, generate all days and submit OFF where needed
    for year, month in month_year_pairs:
        # Generate all days in the month
        all_days = get_all_days(year, month)  # Returns list of datetime.date

        for day in all_days:
            if day not in handled_dates:
                # print(f"{day}")
                header_id = get_header_id_for_date(month_header_dict, day.isoformat())

                payload = {
                    "model[ID]": None if not edit else fetch_existing_entries(header_id, cookie).get(day.isoformat(), "00000000-0000-0000-0000-000000000000"),
                    "model[LogBookHeaderID]": header_id,
                    "model[Date]": day.isoformat(),
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
                        # Create a mock response object with the .ok attribute
                        class MockResponse:
                            ok = True
                            status_code = 200
                            text = 'Debug mode: simulated response'

                        response = MockResponse()

                    if response.status_code == 200:
                        log_message(f"🟡 OFF submitted for {day}.")
                    else:
                        log_message(f"❌ Failed OFF for {day}: {response.status_code} - {response.text}")
                except Exception as e:
                    log_message(f"❌ Network error submitting OFF for {day}: {e}")



# === GUI SETUP ===
def browse_file():
    path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    entry_file.delete(0, tk.END)
    entry_file.insert(0, path)

def start_process():
    output_box.delete(1.0, tk.END)
    file_path = entry_file.get()
    clock_in = entry_clockin.get().strip() or "09:00 am"
    clock_out = entry_clockout.get().strip() or "06:00 pm"
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
    
    

    thread = threading.Thread(target=process_logbook, args=(file_path, cookie, clock_in, clock_out, is_edit, month_header_dict))
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
                root.after(0, update_gui_fields)
            except Exception as e:
                log_message(f"❌ Failed to get cookie and header: {e}")
                log_message("❌❌❌ Something went wrong! Please click on 'Fetch Cookies & Header ID' again. 🛠️🔄")

        
        # Start the fetch_data function in a new thread
        threading.Thread(target=fetch_data, daemon=True).start()

    def update_gui_fields():
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

tk.Button(root, text="Fetch Cookie & Header ID", command=get_cookie_and_header, bg="blue", fg="white").grid(row=1, column=0, columnspan=3, pady=(2, 0))

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
