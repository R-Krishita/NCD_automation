"""
NCD Portal Auto Data Entry Tool
================================
This script reads an Excel file and automatically fills
the NCD portal "NCD ENROLLMENT" form row by row.

BEFORE RUNNING:
1. Add your username/password in config.py
2. Set EXCEL_FILE_PATH and SHEET_NAME correctly in config.py
3. Run once with SHOW_BROWSER = True

NOTE:
- Explicit selectors in FIELD_MAP and login selectors are optional overrides.
- If placeholders are left as-is, the script uses smart locator fallbacks.
"""

import re
import os
import sys
import time

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options

import config


# ============================================================
# OPTIONAL EXPLICIT FIELD MAP (override smart locator behavior)
# ============================================================
# Excel column name : (selector_type, selector_value)
# selector_type can be "id", "name", "css", "xpath", or "class"
FIELD_MAP = {
    "District": ("id", "PLACEHOLDER_ID_district"),
    "Taluka": ("id", "PLACEHOLDER_ID_taluka"),
    "PHC": ("id", "PLACEHOLDER_ID_phc"),
    "SHC": ("id", "PLACEHOLDER_ID_shc"),
    "Village": ("id", "PLACEHOLDER_ID_village"),
}

LOCATION_FIELDS = ("District", "Taluka", "PHC", "SHC", "Village")
SECTION_FIELDS = {
    "Subcenter Information": LOCATION_FIELDS,
    "Family Information": ("Individual ID", "Family ID", "Family Contact #"),
    "Fill Individual Details": (
        "First Name",
        "Middle Name",
        "Last Name",
        "Age",
        "Sex",
        "Mobile #",
        "Address",
    ),
    "ID Information": ("ABHA Number",),
}

# Optional login selectors (fallbacks are auto-generated if placeholders remain)
LOGIN_USERNAME_SELECTOR = ("id", "PLACEHOLDER_ID_username")
LOGIN_PASSWORD_SELECTOR = ("id", "PLACEHOLDER_ID_password")
LOGIN_BUTTON_SELECTOR = ("xpath", "//button[contains(text(),'Login')]")

# Optional explicit navigation selectors
ENROLLMENT_BUTTON_SELECTOR = (
    "xpath",
    "//button[contains(text(),'NCD ENROLLMENT')] | //a[contains(text(),'NCD ENROLLMENT')]",
)
SUBMIT_BUTTON_SELECTOR = ("xpath", "//button[contains(text(),'Submit') or contains(text(),'Save')]")


FIELD_SYNONYMS = {
    "district": ["district"],
    "taluka": ["taluka", "block"],
    "phc": ["phc", "primary health centre", "primary health center"],
    "shc": ["shc", "sub health centre", "sub health center"],
    "village": ["village", "hamlet"],
    "individual": ["individual", "individual id", "number", "id"],
    "name": ["name", "full name", "patient name"],
    "age": ["age"],
    "gender": ["gender", "sex"],
    "mobile": ["mobile", "phone", "contact", "mobile no", "phone number"],
    "aadhar": ["aadhar", "aadhaar", "uid"],
}


def by_type(selector_type):
    mapping = {
        "id": By.ID,
        "name": By.NAME,
        "css": By.CSS_SELECTOR,
        "xpath": By.XPATH,
        "class": By.CLASS_NAME,
    }
    if selector_type not in mapping:
        raise ValueError(f"Unsupported selector type: {selector_type}")
    return mapping[selector_type]


def is_placeholder_selector(selector):
    if not selector:
        return True
    if not isinstance(selector, tuple) or len(selector) != 2:
        return False
    return "PLACEHOLDER" in str(selector[1]).upper()


def normalize_words(text):
    return re.findall(r"[a-z0-9]+", text.lower())


def xpath_literal(value):
    if "'" not in value:
        return f"'{value}'"
    if '"' not in value:
        return f'"{value}"'
    parts = value.split("'")
    quoted = ", \"'\", ".join(f"'{part}'" for part in parts)
    return f"concat({quoted})"


def dedupe_selectors(selectors):
    seen = set()
    ordered = []
    for sel in selectors:
        if not isinstance(sel, tuple) or len(sel) != 2:
            continue
        key = (sel[0], sel[1])
        if key in seen:
            continue
        seen.add(key)
        ordered.append(sel)
    return ordered


def find_visible_element(driver, selector):
    sel_type, sel_value = selector
    elements = driver.find_elements(by_type(sel_type), sel_value)
    for el in elements:
        try:
            if el.is_displayed():
                return el
        except StaleElementReferenceException:
            continue
    return None


def wait_and_find(driver, selector, timeout=15):
    sel_type, sel_value = selector
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by_type(sel_type), sel_value))
    )


def wait_and_find_any(driver, selectors, timeout=15):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for selector in selectors:
            el = find_visible_element(driver, selector)
            if el is not None:
                return el
        time.sleep(0.2)
    raise TimeoutException(f"No element found for selectors: {selectors}")


def wait_and_click(driver, selector, timeout=15):
    sel_type, sel_value = selector
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by_type(sel_type), sel_value))
    )
    el.click()
    return el


def wait_and_click_any(driver, selectors, timeout=15):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for selector in selectors:
            try:
                el = find_visible_element(driver, selector)
                if el is not None and el.is_enabled():
                    el.click()
                    return el
            except StaleElementReferenceException:
                continue
        time.sleep(0.2)
    raise TimeoutException(f"No clickable element found for selectors: {selectors}")


def dismiss_full_screen_backdrop(driver, timeout=15):
    """Wait for Angular modal/loading backdrops to disappear."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        visible_backdrops = []
        for backdrop in driver.find_elements(By.CSS_SELECTOR, ".backdrop.full-screen"):
            try:
                if backdrop.is_displayed():
                    visible_backdrops.append(backdrop)
            except StaleElementReferenceException:
                continue

        if not visible_backdrops:
            return

        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        except Exception:
            pass
        time.sleep(0.25)

    raise TimeoutException("A full-screen portal backdrop did not disappear")


def field_keywords(column_name):
    words = normalize_words(column_name)
    phrases = {" ".join(words), "".join(words), *words}

    for key, synonyms in FIELD_SYNONYMS.items():
        if key in phrases:
            phrases.update(synonyms)

    return [p.strip() for p in phrases if p.strip()]


def build_smart_field_selectors(column_name):
    selectors = []
    keywords = field_keywords(column_name)

    for keyword in keywords:
        lit = xpath_literal(keyword.lower())
        contains_attributes = (
            "placeholder",
            "aria-label",
            "formcontrolname",
        )
        selectors.extend(
            [
                ("name", keyword),
                ("id", keyword),
                ("xpath", f"//*[@name={lit}]"),
                ("xpath", f"//*[@id={lit}]"),
                (
                    "xpath",
                    f"//label[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})]"
                    "//*[self::input or self::textarea or self::select]",
                ),
                (
                    "xpath",
                    f"//label[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})]"
                    "/following::*[self::input or self::textarea or self::select][1]",
                ),
            ]
        )
        if keyword not in ("name", "user", "login"):
            selectors.extend(
                [
                    (
                        "xpath",
                        f"//*[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})]",
                    ),
                    (
                        "xpath",
                        f"//*[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})]",
                    ),
                ]
            )
        if keyword != "name":
            for attribute in contains_attributes:
                selectors.append(
                    (
                        "xpath",
                        f"//*[contains(translate(@{attribute},'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})]",
                    )
                )

    return dedupe_selectors(selectors)


def build_username_selectors():
    keywords = ["username", "user name", "user", "login", "mobile", "email"]
    selectors = [LOGIN_USERNAME_SELECTOR]
    for keyword in keywords:
        lit = xpath_literal(keyword.lower())
        selectors.extend(
            [
                ("name", keyword),
                ("id", keyword),
                (
                    "xpath",
                    f"//input[(contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})"
                    f" or contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})"
                    f" or contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})"
                    f" or contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit}))"
                    " and @type!='password']",
                ),
                (
                    "xpath",
                    f"//label[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})]"
                    "/following::input[1]",
                ),
            ]
        )
    return dedupe_selectors([s for s in selectors if not is_placeholder_selector(s)])


def build_password_selectors():
    selectors = [LOGIN_PASSWORD_SELECTOR]
    selectors.extend(
        [
            ("xpath", "//input[@type='password']"),
            ("name", "password"),
            ("id", "password"),
            (
                "xpath",
                "//*[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password')]",
            ),
            (
                "xpath",
                "//*[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password')]",
            ),
        ]
    )
    return dedupe_selectors([s for s in selectors if not is_placeholder_selector(s)])


def build_login_button_selectors():
    selectors = [LOGIN_BUTTON_SELECTOR]
    selectors.extend(
        [
            (
                "xpath",
                "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]",
            ),
            (
                "xpath",
                "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'sign in')]",
            ),
            (
                "xpath",
                "//input[@type='submit' and contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]",
            ),
        ]
    )
    return dedupe_selectors([s for s in selectors if not is_placeholder_selector(s)])


def build_ncd_login_selectors():
    return dedupe_selectors(
        [
            (
                "xpath",
                "//*[self::button or self::a or self::input]"
                "[contains(translate(normalize-space(.)"
                ",'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ncd login')"
                " or contains(translate(@value"
                ",'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ncd login')]",
            ),
            (
                "xpath",
                "//*[self::button or self::a or self::input]"
                "[contains(translate(normalize-space(.)"
                ",'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ncd')]",
            ),
        ]
    )


def build_enrollment_button_selectors():
    selectors = [ENROLLMENT_BUTTON_SELECTOR]
    selectors.extend(
        [
            (
                "xpath",
                "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ncd enrollment')]",
            ),
            (
                "xpath",
                "//a[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ncd enrollment')]",
            ),
            (
                "xpath",
                "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'enrollment')]",
            ),
            (
                "xpath",
                "//a[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'enrollment')]",
            ),
        ]
    )
    return dedupe_selectors([s for s in selectors if not is_placeholder_selector(s)])


def build_submit_button_selectors():
    selectors = [SUBMIT_BUTTON_SELECTOR]
    selectors.extend(
        [
            (
                "xpath",
                "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'submit')]",
            ),
            (
                "xpath",
                "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'save')]",
            ),
            (
                "xpath",
                "//input[@type='submit']",
            ),
        ]
    )
    return dedupe_selectors([s for s in selectors if not is_placeholder_selector(s)])


def login(driver):
    print("-> Opening login page...")
    driver.get(config.PORTAL_URL)
    time.sleep(max(config.WAIT_TIME, 2))

    username_selectors = build_username_selectors()
    password_selectors = build_password_selectors()
    login_button_selectors = build_login_button_selectors()

    # The portal may first show a landing page with an NCD Login entry point.
    # Do not click it when the credential form is already present.
    if find_visible_element(driver, password_selectors[0]) is None:
        print("-> Opening NCD Login...")
        wait_and_click_any(driver, build_ncd_login_selectors(), timeout=20)
        time.sleep(max(config.WAIT_TIME, 2))

    dismiss_full_screen_backdrop(driver)
    user_field = wait_and_find_any(driver, username_selectors, timeout=20)
    user_field.clear()
    user_field.send_keys(config.USERNAME)

    pass_field = wait_and_find_any(driver, password_selectors, timeout=20)
    pass_field.clear()
    pass_field.send_keys(config.PASSWORD)

    wait_and_click_any(driver, login_button_selectors, timeout=15)
    time.sleep(max(config.WAIT_TIME, 2))
    dismiss_full_screen_backdrop(driver)

    still_on_login = any(
        find_visible_element(driver, selector) is not None
        for selector in password_selectors
    ) and find_visible_element(
        driver, ("xpath", "//*[contains(@id,'enrollment_menu')]")
    ) is None
    if still_on_login:
        print("[WARN] Password field is still visible after login click.")
        print("       Check credentials, captcha/OTP, or login selectors.")
    else:
        print("-> Login click submitted and login form is no longer visible.")


def open_enrollment_form(driver):
    print("-> Opening NCD Enrollment form...")
    dismiss_full_screen_backdrop(driver)
    wait_and_click_any(driver, build_enrollment_button_selectors(), timeout=20)
    time.sleep(config.WAIT_TIME)


def select_dropdown_value(field, value):
    select = Select(field)
    try:
        select.select_by_visible_text(value)
        return
    except NoSuchElementException:
        pass

    value_norm = value.strip().lower()
    for option in select.options:
        if option.text.strip().lower() == value_norm:
            option.click()
            return
    raise NoSuchElementException(f"No dropdown option matched '{value}'")


def wait_for_suggestion_and_select(driver, value, timeout=5):
    """Wait for common suggestion containers and click a matching item."""
    lit = xpath_literal(value.lower())
    option_xpath = (
        "//*[self::li or self::a or @role='option' or "
        "contains(@class,'ng-option') or contains(@class,'select2-results__option') or "
        "contains(@class,'autocomplete-suggestion')]"
    )
    match_xpaths = [
        f"{option_xpath}[normalize-space(.)={lit}]",
        f"{option_xpath}[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), {lit})]",
    ]

    for xp in match_xpaths:
        try:
            el = wait_and_find_any(driver, [("xpath", xp)], timeout=timeout)
            el.click()
            return True
        except TimeoutException:
            continue
    return False


def field_value(field):
    """Return the value currently displayed by an input or custom widget."""
    tag = field.tag_name.lower()
    if tag in ("input", "textarea", "select"):
        return (field.get_attribute("value") or "").strip()
    return (field.text or "").strip()


def value_was_selected(field, expected):
    actual = field_value(field)
    return actual.casefold() == expected.strip().casefold()


def handle_autocomplete_input(driver, field, value):
    """Handle non-<select> dropdowns/autocomplete widgets.

    Strategy:
    - click the field
    - send the text (with small pauses)
    - try to pick a suggestion matching text
    - if no suggestion is clicked, use keyboard selection and verify the value
    """
    try:
        field.click()
    except Exception:
        pass

    try:
        field.clear()
    except Exception:
        pass

    # type slowly to trigger client-side autocomplete
    for ch in value:
        field.send_keys(ch)
        time.sleep(0.05)
    time.sleep(0.3)

    # try to select a visible suggestion
    picked = wait_for_suggestion_and_select(driver, value, timeout=3)
    if picked:
        return True

    # Keyboard selection is useful for Angular widgets whose option list is
    # rendered outside the input. It is only accepted when the value remains.
    try:
        field.send_keys(Keys.ARROW_DOWN, Keys.ENTER)
        time.sleep(0.3)
        return value_was_selected(field, value)
    except Exception:
        return False


def resolve_field_selectors(column_name, explicit_selector):
    selectors = []
    if explicit_selector and not is_placeholder_selector(explicit_selector):
        selectors.append(explicit_selector)
    selectors.extend(build_smart_field_selectors(column_name))
    return dedupe_selectors(selectors)


def handle_ng_select_widget(driver, field, value):
    """Handle Angular/ng-select-like widgets where the visible element is a div.

    Strategy:
    - Try to find a descendant input and type into it
    - If no input, click the container to open the options and try to select
    - Search globally for option elements matching value (handles dropdowns appended to body)
    """
    # Try to find an input inside the widget
    input_selectors = ['input[type="text"]', 'input[role="combobox"]', 'input', '.ng-input input']
    for sel in input_selectors:
        try:
            inputs = field.find_elements(By.CSS_SELECTOR, sel)
            if inputs:
                inp = inputs[0]
                return handle_autocomplete_input(driver, inp, value)
        except Exception:
            continue

    # No input found: click to open, then try to select suggestion by text
    try:
        field.click()
        time.sleep(0.2)
    except Exception:
        pass

    picked = wait_for_suggestion_and_select(driver, value, timeout=3)
    if picked:
        return True

    # fallback: try pressing Enter on the container
    try:
        field.send_keys(Keys.ARROW_DOWN, Keys.ENTER)
        time.sleep(0.2)
        return value_was_selected(field, value)
    except Exception:
        return False


def open_information_section(driver, section_name):
    """Open one of the portal's expandable information sections."""
    lit = xpath_literal(section_name.lower())
    selectors = [
        (
            "xpath",
            f"//*[self::button or self::a or @role='button']"
            f"[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),{lit})]",
        ),
        (
            "xpath",
            f"//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),{lit})]"
            f"/ancestor::*[self::button or @role='button' or self::a][1]",
        ),
    ]
    section = wait_and_find_any(driver, selectors, timeout=20)
    try:
        section.click()
    except Exception:
        driver.execute_script("arguments[0].click();", section)
    time.sleep(0.5)


def fill_field(driver, column_name, value):
    """Locate and fill one field, selecting an option when it is a dropdown."""
    selectors = resolve_field_selectors(column_name, None)
    field = wait_and_find_any(driver, selectors, timeout=10)
    tag = field.tag_name.lower()
    classes = (field.get_attribute("class") or "").lower()

    if tag == "select":
        select_dropdown_value(field, value)
        return

    ng_like = "ng-select" in classes or "ng-value" in classes
    if ng_like or column_name in LOCATION_FIELDS:
        if not handle_ng_select_widget(driver, field, value):
            raise RuntimeError(f"Option '{value}' was not selected for '{column_name}'")
        return

    target = field
    if tag not in ("input", "textarea"):
        target = field.find_element(By.CSS_SELECTOR, "input, textarea")
    target.clear()
    target.send_keys(value)
    time.sleep(0.2)


def fill_one_entry(driver, row):
    """Open each portal section and fill its corresponding workbook columns."""
    for section_name, columns in SECTION_FIELDS.items():
        available = [
            column
            for column in columns
            if column in row and pd.notna(row[column]) and str(row[column]).strip()
        ]
        if not available:
            continue

        print(f"   -> Opening {section_name}...")
        open_information_section(driver, section_name)
        for column_name in available:
            value = str(row[column_name]).strip()
            print(f"      Filling {column_name}")
            fill_field(driver, column_name, value)

    time.sleep(1)
    wait_and_click_any(driver, build_submit_button_selectors(), timeout=20)
    time.sleep(config.WAIT_TIME)


def main():
    print("=" * 55)
    print("NCD Portal Auto Data Entry Tool")
    print("=" * 55)

    try:
        df = pd.read_excel(config.EXCEL_FILE_PATH, sheet_name=config.SHEET_NAME)
    except Exception as e:
        print(f"[ERROR] Could not read Excel file: {e}")
        sys.exit(1)

    max_rows = os.environ.get("NCD_MAX_ROWS")
    if max_rows:
        try:
            df = df.head(int(max_rows))
        except ValueError:
            print(f"[ERROR] NCD_MAX_ROWS must be an integer, got '{max_rows}'.")
            sys.exit(1)

    print(f"-> Found {len(df)} rows in Excel.")

    options = Options()
    if not config.SHOW_BROWSER:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)

    success_count = 0
    fail_count = 0

    try:
        login(driver)

        for idx, row in df.iterrows():
            display_name = " ".join(
                str(row.get(column, "")).strip()
                for column in ("First Name", "Middle Name", "Last Name")
                if pd.notna(row.get(column)) and str(row.get(column)).strip()
            )
            print(f"\n-> Filling entry {idx + 1}/{len(df)}: {display_name}")
            try:
                open_enrollment_form(driver)
                fill_one_entry(driver, row)
                success_count += 1
                print(f"   [OK] Entry {idx + 1} saved.")
            except Exception as e:
                fail_count += 1
                print(f"   [FAIL] Error in entry {idx + 1}: {e}")
                continue

    finally:
        print("\n" + "=" * 55)
        print(f"Completed. Success: {success_count} | Fail: {fail_count}")
        print("=" * 55)
        if sys.stdin and sys.stdin.isatty():
            try:
                input("Press Enter to close the browser...")
            except EOFError:
                pass
        driver.quit()


if __name__ == "__main__":
    main()
