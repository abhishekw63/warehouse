"""
=============================================================================
D365 Business Central - Staging M-Assist Sales Order Automation
=============================================================================
Company     : Renee Cosmetics Pvt Ltd
File        : d365_automation.py
Version     : 1.0.0

HOW TO USE:
-----------
FIRST TIME (save login session):
    python d365_automation.py --save-login

DAILY RUN (today's date):
    python d365_automation.py

SPECIFIC DATE:
    python d365_automation.py --date 2026-05-10

VISIBLE BROWSER (default):
    python d365_automation.py --visible

BACKGROUND (no window):
    python d365_automation.py --headless

INSTALL DEPENDENCIES (run once):
    pip install playwright
    playwright install chromium
=============================================================================
"""

import argparse
import io
import logging
import sys
import time
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


# =============================================================================
# =============================================================================

# -- D365 URL ------------------------------------------------------------------
D365_URL = (
    "https://businesscentral.dynamics.com/"
    "914df76c-6ead-4a02-a626-12d149abb825/"
    "Production?company=Renee%20Cosmetics%20Pvt%20Ltd-Final&dc=0"
)

# -- Browser -------------------------------------------------------------------
DEFAULT_HEADLESS = False   # False = visible browser window
SLOW_MO_MS       = 400    # ms pause between each browser action
TIMEOUT_MS       = 60_000 # max wait per step in ms (60 seconds)

# -- Real Chrome (so Windows Hello / Passkey works during --save-login) --------
import os as _os
CHROME_EXE         = _os.path.expandvars(r'%PROGRAMFILES%\Google\Chrome\Application\chrome.exe')
CHROME_PROFILE_DIR = _os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\User Data')
_CHROME_FALLBACKS  = [
    _os.path.expandvars(r'%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe'),
    _os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
]

def _find_chrome():
    for path in [CHROME_EXE] + _CHROME_FALLBACKS:
        if _os.path.exists(path):
            return path
    return CHROME_EXE

# -- Paths ---------------------------------------------------------------------
BASE_DIR        = Path(__file__).parent
AUTH_FILE       = BASE_DIR / "auth_state.json"   # saved login session
LOG_DIR         = BASE_DIR / "logs"
SCREENSHOT_DIR  = BASE_DIR / "screenshots"

# -- Error handling ------------------------------------------------------------
SCREENSHOT_ON_ERROR = True   # save screenshot automatically when a step fails


# =============================================================================
# LOGGING
# =============================================================================

LOG_DIR.mkdir(parents=True, exist_ok=True)
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

log_filename = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

# Force stdout to UTF-8 on Windows (fixes UnicodeEncodeError with cp1252)
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),  # UTF-8 in file
        logging.StreamHandler(sys.stdout),                    # UTF-8 in terminal
    ],
)
log = logging.getLogger("d365")


# =============================================================================
# HELPERS
# =============================================================================

def take_take_screenshot(page, label: str):
    """Save a screenshot to the screenshots folder."""
    if SCREENSHOT_ON_ERROR:
        path = SCREENSHOT_DIR / f"{label}_{datetime.now().strftime('%H%M%S')}.png"
        try:
            page.screenshot(path=str(path))
            log.info(f"[SCREENSHOT] Saved -> {path.name}")
        except Exception:
            pass


def banner(msg: str):
    log.info("-" * 55)
    log.info(f"  {msg}")
    log.info("-" * 55)


# =============================================================================
# =============================================================================

def step1_open_d365(page) -> bool:
    """Open D365 and confirm page loaded by checking the URL."""
    log.info("[STEP 1] Opening D365 Business Central...")
    try:
        page.goto(D365_URL, wait_until="domcontentloaded", timeout=TIMEOUT_MS)
        time.sleep(4)  # allow SPA to hydrate

        current_url = page.url
        log.info(f"[STEP 1] Current URL: {current_url}")

        if "businesscentral.dynamics.com" in current_url:
            log.info("[STEP 1] OK - D365 loaded successfully.")
            return True

        if "login.microsoftonline.com" in current_url:
            log.error("[STEP 1] FAIL - Redirected to login. Session expired. Run --save-login again.")
            take_take_screenshot(page, "step1_login_required")
            return False

        log.error(f"[STEP 1] FAIL - Unexpected URL: {current_url}")
        take_take_screenshot(page, "step1_unexpected_url")
        return False

    except Exception as e:
        log.error(f"[STEP 1] FAIL - {e}")
        take_take_screenshot(page, "step1_error")
        return False


def step2_search_staging(page) -> bool:
    """Open D365 search and type 'STAGING'."""
    try:
        # Open search with keyboard shortcut
        page.keyboard.press("Alt+Q")
        time.sleep(1.2)

        # Find the search input
        search_input = page.wait_for_selector(
            'input[placeholder*="Tell me"], input[aria-label*="Search"], '
            'input[aria-label*="search"], [role="combobox"]',
            timeout=TIMEOUT_MS,
        )
        search_input.fill("")
        search_input.type("STAGING", delay=80)
        time.sleep(1.5)  # wait for autocomplete to populate

        log.info("[STEP 2] OK - Typed 'STAGING' in search.")
        return True
    except Exception as e:
        log.error(f"[STEP 2] FAIL - Search failed: {e}")
        take_screenshot(page, "step2_error")
        return False


def step3_click_staging_page(page) -> bool:
    """Click 'Staging M-Assist Sales Order' from search results."""
    log.info("[STEP 3] Clicking 'Staging M-Assist Sales Order'...")
    try:
        page.click("text=Staging M-Assist Sales Order", timeout=TIMEOUT_MS)
        page.wait_for_load_state("networkidle", timeout=TIMEOUT_MS)
        log.info("[STEP 3] OK - Staging page opened.")
        return True
    except Exception as e:
        log.error(f"[STEP 3] FAIL - Could not click Staging page: {e}")
        take_screenshot(page, "step3_error")
        return False


def step4_click_integration_button(page) -> bool:
    """Click the 'M-Assist Sales Order Integration' button in the toolbar."""
    log.info("[STEP 4] Clicking 'M-Assist Sales Order Integration'...")
    try:
        page.click("text=M-Assist Sales Order Integration", timeout=TIMEOUT_MS)
        time.sleep(1.5)
        log.info("[STEP 4] OK - Integration button clicked.")
        return True
    except Exception as e:
        log.error(f"[STEP 4] FAIL - Button not found: {e}")
        take_screenshot(page, "step4_error")
        return False


def step5_set_date(page, order_date: date) -> bool:
    """Set the date field in the dialog (if present)."""
    log.info(f"[STEP 5] Setting date to {order_date}...")
    try:
        # Indian locale format DD/MM/YYYY
        date_str = order_date.strftime("%d/%m/%Y")

        # Try multiple selectors D365 uses for date fields
        selectors = [
            'input[class*="date"]',
            'input[type="date"]',
            '[aria-label*="Date"]',
            '[aria-label*="date"]',
            '[data-testid*="date"]',
        ]
        date_field = None
        for sel in selectors:
            try:
                date_field = page.locator(sel).first
                if date_field.count() > 0:
                    break
            except Exception:
                continue

        if date_field and date_field.count() > 0:
            date_field.triple_click()
            date_field.type(date_str, delay=60)
            page.keyboard.press("Tab")
            time.sleep(0.5)
            log.info(f"[STEP 5] OK - Date set to {date_str}.")
        else:
            log.info("[STEP 5] OK - No date field found - skipping (may not be required).")

        return True
    except Exception as e:
        log.warning(f"[STEP 5] WARNING - Date field issue (non-fatal): {e}")
        return True   # Non-fatal - some flows don't need a date


def step6_click_import(page) -> bool:
    """Click the Import button."""
    log.info("[STEP 6] Clicking 'Import'...")
    try:
        page.click("text=Import", timeout=TIMEOUT_MS)
        log.info("[STEP 6] OK - Import triggered.")
        return True
    except Exception as e:
        log.error(f"[STEP 6] FAIL - Import button not found: {e}")
        take_screenshot(page, "step6_error")
        return False


def step7_wait_for_completion(page) -> bool:
    """Wait for 'Working on it...' spinner to disappear."""
    log.info("[STEP 7] Waiting for import to complete...")
    try:
        # Wait for spinner to appear (it may take a moment)
        try:
            page.wait_for_selector("text=Working on it", timeout=10_000)
            log.info("    Import in progress ('Working on it...' detected)...")
        except PlaywrightTimeout:
            log.info("    No loading dialog detected - import may be quick.")

        # Wait for spinner to disappear (up to 5 minutes)
        page.wait_for_selector(
            "text=Working on it",
            state="hidden",
            timeout=300_000,
        )
        log.info("[STEP 7] OK - Import completed successfully!")
        return True
    except Exception as e:
        log.error(f"[STEP 7] FAIL - Timed out waiting for import: {e}")
        take_screenshot(page, "step7_timeout")
        return False


# =============================================================================
# =============================================================================

def run_automation(order_date: date = None, headless: bool = DEFAULT_HEADLESS) -> bool:
    """
    Run the complete Staging M-Assist import automation.

    Args:
        order_date : The date to use for the import. Defaults to today.
        headless   : If True, browser runs invisibly in the background.

    Returns:
        True if all steps succeeded, False otherwise.
    """
    if order_date is None:
        order_date = date.today()

    banner("D365 Staging M-Assist Automation")
    log.info(f"  Date      : {order_date.strftime('%d %B %Y')}")
    log.info(f"  Headless  : {headless}")
    log.info(f"  Auth File : {'Found' if AUTH_FILE.exists() else 'NOT FOUND - run --save-login'}")
    log.info(f"  Log File  : {log_filename.name}")

    steps = [
        ("Open D365",                  step1_open_d365),
        ("Search Staging",             step2_search_staging),
        ("Click Staging Page",         step3_click_staging_page),
        ("Click Integration Button",   step4_click_integration_button),
        ("Set Date",                   lambda p: step5_set_date(p, order_date)),
        ("Click Import",               step6_click_import),
        ("Wait for Completion",        step7_wait_for_completion),
    ]

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=headless,
            slow_mo=SLOW_MO_MS,
            args=["--start-maximized"],
        )
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            # Reuse saved login session if available (created by --save-login)
            storage_state=str(AUTH_FILE) if AUTH_FILE.exists() else None,
        )
        page = context.new_page()
        page.set_default_timeout(TIMEOUT_MS)

        success = True
        for step_num, (step_name, step_fn) in enumerate(steps, 1):
            ok = step_fn(page)
            if not ok:
                log.error(f"[STOPPED] Step {step_num}: {step_name}")
                success = False
                break
            time.sleep(0.8)  # small buffer between steps

        if success:
            banner("SUCCESS - All steps completed!")
            # Save updated session for next run
            context.storage_state(path=str(AUTH_FILE))
            log.info("[INFO] Login session refreshed and saved.")
        else:
            banner("FAILED - Check logs and screenshots folder")

        browser.close()

    return success


# =============================================================================
# LOGIN SESSION SAVER
# =============================================================================

def save_login():
    """
    Opens your REAL installed Chrome (not Playwright Chromium) so that
    Windows Hello / Passkey works exactly as it does when you log in manually.
    After login, saves session cookies to auth_state.json.
    All future automation runs reuse this session - no login needed.
    """
    chrome_exe  = _find_chrome()
    profile_dir = CHROME_PROFILE_DIR

    print()
    print("=" * 60)
    print("  D365 Login Session Saver")
    print("=" * 60)
    print()
    print("  Opening YOUR real Chrome so Windows Hello works.")
    print()
    print("  Steps:")
    print("  1. Your Chrome browser will open now.")
    print("  2. Click:  abhishek.wagh@reneecosmetics.in")
    print("  3. Complete Windows Hello PIN as normal.")
    print("  4. Click YES on Stay signed in? prompt.")
    print("  5. Wait for D365 HOME PAGE to fully load.")
    print("  6. Come back here and press ENTER.")
    print()
    print("  WARNING: Close all other Chrome windows first!")
    print("  WARNING: Do NOT close the browser before pressing ENTER!")
    print()
    print(f"  Chrome  : {chrome_exe}")
    print(f"  Profile : {profile_dir}")
    print()

    if not _os.path.exists(chrome_exe):
        print("  ERROR: Chrome not found!")
        print("  Update CHROME_EXE in the CONFIG section at the top of the file.")
        return

    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            executable_path=chrome_exe,
            headless=False,
            slow_mo=200,
            args=[
                "--start-maximized",
                "--profile-directory=Default",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-blink-features=AutomationControlled",
            ],
        )

        page = context.pages[0] if context.pages else context.new_page()
        page.goto(D365_URL)

        print("  Chrome opened. Log in now...")
        print()
        input("  >> D365 home page fully loaded? Press ENTER to save session: ")

        context.storage_state(path=str(AUTH_FILE))
        context.close()

    print()
    print(f"  Session saved -> {AUTH_FILE.name}")
    print("  Run  python d365_automation.py  - no login needed from now on.")
    print("  Session lasts ~30-90 days before renewal needed.")
    print()


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="D365 Staging M-Assist Sales Order Automation - Renee Cosmetics",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="""
Examples:
  python d365_automation.py                        # run for today
  python d365_automation.py --date 2026-05-10      # run for specific date
  python d365_automation.py --headless             # run without browser window
  python d365_automation.py --save-login           # first-time login setup
        """,
    )
    parser.add_argument(
        "--save-login",
        action="store_true",
        help="Open browser for manual login and save session (run this first time)",
    )
    parser.add_argument(
        "--date",
        type=str,
        default=None,
        metavar="YYYY-MM-DD",
        help="Order date (default: today)",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        default=DEFAULT_HEADLESS,
        help="Run browser in background (no visible window)",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Force visible browser window",
    )

    args = parser.parse_args()

    # -- Save login mode -------------------------------------------------------
    if args.save_login:
        save_login()
        sys.exit(0)

    # -- Parse date ------------------------------------------------------------
    run_date = date.today()
    if args.date:
        try:
            run_date = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print("ERROR: Invalid date. Use YYYY-MM-DD e.g. --date 2026-05-10")
            sys.exit(1)

    # -- Headless flag ---------------------------------------------------------
    headless_mode = DEFAULT_HEADLESS
    if args.headless:
        headless_mode = True
    if args.visible:
        headless_mode = False

    # -- Warn if no auth file --------------------------------------------------
    if not AUTH_FILE.exists():
        print()
        print("WARNING: No saved login session found!")
        print("   Run this first:  python d365_automation.py --save-login")
        print()

    # -- Run -------------------------------------------------------------------
    success = run_automation(order_date=run_date, headless=headless_mode)
    sys.exit(0 if success else 1)