"""
Automated Outlook Signup with Playwright + Groq LLM
Loads GROQ_API_KEY from .env file using python-dotenv.
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


def try_fill_email(page, handle):
    page.wait_for_selector('input[name="New email"]', timeout=15000)
    page.fill('input[name="New email"]', handle)
    page.get_by_test_id("primaryButton").click()


def fill_password(page, password):
    page.wait_for_selector('input[type="password"]', timeout=15000)
    page.fill('input[type="password"]', password)
    page.get_by_test_id("primaryButton").click()


def select_random_option_from_combobox(page, combobox_selector):
    try:
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
    if not visible_indices:
        options.nth(0).click()
        return True

    pick = random.choice(visible_indices)
    options.nth(pick).click()
    return True


def fill_birthdate(page):
    select_random_option_from_combobox(page, "#BirthMonthDropdown")
    time.sleep(0.3)
    select_random_option_from_combobox(page, "#BirthDayDropdown")
    time.sleep(0.3)
    year = random.randint(1990, 2002)
    page.fill('input[name="BirthYear"]', str(year))
    page.get_by_test_id("primaryButton").click()


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=150)
        context = browser.new_context()
        page = context.new_page()
        page.goto(SIGNUP_URL)

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

        first_name = name.split()[0].lower()
        password = f"{first_name}@12345"
        print(f"Generated Password: {password}")
        fill_password(page, password)

        time.sleep(2)
        fill_birthdate(page)
        print("✅ Birthdate submitted. Browser will stay open.")
        input("Press Enter to close...")
        browser.close()


if __name__ == "__main__":
    main()
