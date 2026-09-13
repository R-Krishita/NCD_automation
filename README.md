# NCD Portal Auto Data Entry Tool — Step-by-Step Guide

## ⚠️ Read this first
This tool is intended to use **your own** login for **your own** NCD data entry work.
Please confirm your department and portal policy before using automation. Some workflows
may require manual verification, so check with your supervisor/IT team if needed.

---

## Step 1: Install Python
1. Go to https://www.python.org/downloads/
2. Download and install the latest version
3. During install, make sure **Add Python to PATH** is checked

To verify, open Command Prompt and run:
```
python --version
```

## Step 2: Ensure Google Chrome is installed
If you already use Chrome, you can skip this step.

## Step 3: Extract this project folder
Extract the zip to a location such as:
`C:\Users\YourName\ncd_automation\`

This folder includes:
- `auto_entry.py` — main script
- `config.py` — login details and settings
- `requirements.txt` — required libraries
- `data_template.csv` — sample data format

## Step 4: Install required libraries
Open Command Prompt and run:
```
cd C:\Users\YourName\ncd_automation
pip install -r requirements.txt
```

## Step 5: Add your login in config.py
Open `config.py` and update:
```python
USERNAME = "124-cho196941"
PASSWORD = "your_actual_password"
```

## Step 6: Prepare your Excel file
Create an Excel file with columns matching `data_template.csv`,
save it as `data.xlsx`, and place it in the same folder.
(Or update `EXCEL_FILE_PATH` in `config.py` to your file path.)

---

## Step 7 (Most Important): Find real form field selectors

The script now has **smart locators** and first tries to find fields by:
- field `name` / `id`
- placeholder text
- label text near the input
- common attribute patterns

So in many cases, you can run without manually filling all selectors.

Use manual selectors only if a field is not detected correctly. The script includes placeholders like `PLACEHOLDER_ID_xxx` for optional override:

1. Open the portal login page in Chrome
2. Right-click the username field → click **Inspect**
3. In HTML, find something like:
   ```html
   <input id="username" name="loginUsername" type="text">
   ```
4. Copy the `id` value (for example `"username"`)
5. In `auto_entry.py`, replace:
   ```python
   LOGIN_USERNAME_SELECTOR = ("id", "PLACEHOLDER_ID_username")
   ```
   with:
   ```python
   LOGIN_USERNAME_SELECTOR = ("id", "username")
   ```
6. Do the same for **password**, **login button**, **NCD ENROLLMENT button**,
   and only those enrollment fields that fail in automatic detection

If a field has no `id`, use `name`:
```python
"Name": ("name", "patientName"),
```

For dropdowns (Gender, Village, etc.), the script already handles `<select>` fields.
Add manual selector only when smart detection fails.

---

## Step 8: Run the script
```
python auto_entry.py
```

Chrome will open, login will be attempted, and each Excel row will be submitted.
For first-time runs, keep `SHOW_BROWSER = True` in `config.py` so you can observe behavior.

---

## Troubleshooting
- **"Could not map field" / "Element not found"** → add manual selector override for that field (Step 7)
- **Login fails** → verify username/password and check if portal requires captcha/OTP
- **Slow portal/network** → increase `WAIT_TIME` in `config.py` (example: `2` to `5`)
