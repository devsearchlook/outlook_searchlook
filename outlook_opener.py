"""
Automated Outlook Signup with Playwright + Groq LLM
Loads GROQ_API_KEY from .env file using python-dotenv.
Attempts to bypass automation detection to make "Press and Hold" captcha appear.
"""

import os
import random
import time
import requests
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# Load .env file
load_dotenv(override=True)
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

if not GROQ_API_KEY:
    raise RuntimeError("Missing GROQ_API_KEY. Please add it to your .env file.")

SIGNUP_URL = "https://go.microsoft.com/fwlink/p/?LinkID=2125440&clcid=0x409&culture=en-us&country=us"


# ------------------ Utility: Human-like Typing ------------------
def human_type(page, selector, text):
    """Type into a field character by character with random delay."""
    page.click(selector)
    for char in text:
        page.keyboard.insert_text(char)
        time.sleep(random.uniform(0.05, 0.2))  # slight variation


# ------------------ Email Generation ------------------
def generate_mexican_email():
    """Generate realistic Mexican name + email handle using Groq LLM."""
    prompt = (
        "Give me a realistic Mexican first name and last name. "
        "Make an email handle by joining them in lowercase with an underscore (_) between them "
        "and then append a random number between 100 and 999 at the end. "
        "Respond in plain text like: Carlos Ramirez -> carlos_ramirez_123"
    )

    resp = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": "llama-3.3-70b-versatile",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 64,
        },
        timeout=30,
    )

    if resp.status_code != 200:
        raise RuntimeError(f"Groq API error: {resp.status_code} {resp.text}")

    content = resp.json()["choices"][0]["message"]["content"].strip()
    if "->" in content:
        name_part, handle_part = content.split("->", 1)
        name = name_part.strip()
        handle = handle_part.strip()
    else:
        first, last = "Carlos", "Ramirez"
        name = f"{first} {last}"
        handle = f"{first.lower()}_{last.lower()}_{random.randint(100,999)}"

    handle = "".join(ch for ch in handle if ch.isalnum() or ch == "_")
    return name, handle


# ------------------ Form Fillers ------------------
def try_fill_email(page, handle):
    page.wait_for_selector('input[name="New email"]', timeout=15000)
    human_type(page, 'input[name="New email"]', handle)
    page.get_by_test_id("primaryButton").click()


def fill_password(page, password):
    page.wait_for_selector('input[type="password"]', timeout=15000)
    human_type(page, 'input[type="password"]', password)
    page.get_by_test_id("primaryButton").click()


def select_random_option_from_combobox(page, combobox_selector):
    try:
        page.locator(combobox_selector).scroll_into_view_if_needed()
        page.click(combobox_selector, timeout=5000)
    except PlaywrightTimeoutError:
        print(f"[warn] Could not click {combobox_selector}")
        return False

    time.sleep(0.3)
    page.wait_for_selector('[role="option"]', timeout=5000)
    options = page.locator('[role="option"]')
    total = options.count()
    if total == 0:
        return False

    visible_indices = [i for i in range(total) if options.nth(i).is_visible()]
    pick = random.choice(visible_indices or [0])
    text = options.nth(pick).inner_text()
    print(f"📅 Selected: {text}")
    options.nth(pick).click()
    return True


def select_month(page):
    """Special handling for BirthMonthDropdown since it's sometimes flaky."""
    try:
        expand_icon = page.locator("#BirthMonthDropdown svg")
        expand_icon.scroll_into_view_if_needed(timeout=5000)
        expand_icon.click(force=True)
    except PlaywrightTimeoutError:
        print("[warn] Could not click month expand icon")
        return False

    try:
        page.wait_for_selector('[role="option"]', timeout=5000)
    except PlaywrightTimeoutError:
        print("[warn] Month options did not appear, retrying click...")
        expand_icon.click(force=True)
        page.wait_for_selector('[role="option"]', timeout=5000)

    options = page.locator('[role="option"]')
    total = options.count()
    if total == 0:
        print("[error] No months available to select.")
        return False

    visible_indices = [i for i in range(total) if options.nth(i).is_visible()]
    pick = random.choice(visible_indices or [0])
    selected_text = options.nth(pick).inner_text()
    print(f"📅 Selected month: {selected_text}")
    options.nth(pick).click()
    return True


def fill_birthdate(page):
    print("Selecting birth month...")
    success = select_month(page)
    if not success:
        print("[warn] Month selection failed, retrying...")
        select_month(page)

    time.sleep(0.3)
    print("Selecting birth day...")
    select_random_option_from_combobox(page, "#BirthDayDropdown")

    time.sleep(0.3)
    year = random.randint(1990, 2002)
    print(f"Typing year: {year}")
    human_type(page, 'input[name="BirthYear"]', str(year))
    page.get_by_test_id("primaryButton").click()


def fill_first_last_name(page, first_name, last_name):
    """Fill first and last name fields on the next screen and press next."""
    try:
        page.wait_for_selector("#firstNameInput", timeout=15000)
        human_type(page, "#firstNameInput", first_name)
        human_type(page, "#lastNameInput", last_name)
        print(f"✅ Filled first name: {first_name}, last name: {last_name}")
        page.get_by_test_id("primaryButton").click()
    except PlaywrightTimeoutError:
        print("[error] Could not find first/last name inputs.")


# ------------------ Main ------------------
def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            slow_mo=120,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(viewport={"width": 1366, "height": 900})
        page = context.new_page()

        # Patch navigator.webdriver to avoid detection
        page.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )

        page.goto(SIGNUP_URL)

        # Simulate some natural mouse movement before interacting
        for _ in range(3):
            page.mouse.move(random.randint(0, 500), random.randint(0, 500))
            time.sleep(0.4)

        while True:
            name, handle = generate_mexican_email()
            print(f"Generated Name: {name}")
            print(f"Generated Handle: {handle}")
            try:
                try_fill_email(page, handle)
            except PlaywrightTimeoutError:
                print("Timeout, retrying...")
                page.goto(SIGNUP_URL)
                continue

            time.sleep(2.5)
            alerts = page.locator('[role="alert"]')
            alert_text = " ".join(
                alerts.nth(i).inner_text().strip()
                for i in range(alerts.count())
                if alerts.nth(i).is_visible()
            )
            if "already" in alert_text:
                print("❌ Email taken, retrying...")
                continue
            print("✅ Email accepted.")
            break

        first_name, last_name = name.split()[0], name.split()[-1]
        password = f"{first_name.lower()}@12345"
        print(f"Generated Password: {password}")
        fill_password(page, password)

        time.sleep(2)
        fill_birthdate(page)

        time.sleep(2)
        fill_first_last_name(page, first_name, last_name)

        # Wait to see if captcha appears
        try:
            page.wait_for_selector("p:has-text('Press and hold')", timeout=15000)
            print("✅ Captcha loaded! Solve manually.")
            page.pause()
        except:
            print("⚠️ Captcha not detected. Microsoft may still be blocking automation.")

        input("Press Enter to close...")
        browser.close()


if __name__ == "__main__":
    main()
