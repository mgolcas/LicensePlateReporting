# License Plate Reporting – Windows Setup Guide

Straightforward instructions for teammates who just need the monthly parking report to run on a Windows PC.

---

## 1. Install prerequisites (one time)

1. **Install Python**
   - Browse to https://www.python.org/downloads/windows/.
   - Download the latest "Windows installer (64-bit)".
   - Run the installer and tick **"Add Python to PATH"** before pressing **Install**.

2. **Copy the project folder**
   - Transfer the entire `LicensePlateReporting` folder (USB, OneDrive, etc.) to the Windows machine, e.g. `C:\LicensePlateReporting`.
   - The folder must contain files such as `aggregate.py`, `config.json`, and `requirements.txt`.

---

## 2. Prepare the application

1. **Open Windows PowerShell**
   - Press **Start**, type **PowerShell**, hit **Enter**.

2. **Create a virtual environment and install dependencies**
   ```powershell
   cd C:\LicensePlateReporting
   python -m venv .venv
   .\.venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. **Update configuration file**
   ```powershell
   notepad config.json
   ```
   Inside Notepad, change:
   - `"source_folder"`: folder that holds the camera Excel exports on this PC.
   - `"output_file"`: desired output location (default `output\parking_durations.xlsx`).
   - Save and close Notepad.

4. **Run a test once**
   ```powershell
   python aggregate.py --config config.json
   ```
   - Open the Excel file referenced by `output_file`. You should see `monthly_totals`, `intervals`, and `issues` (only if problems were detected).

5. **Deactivate the virtual environment (optional)**
   ```powershell
   deactivate
   ```

---

## 3. Schedule automatic monthly runs

1. Open **Task Scheduler** (Start → search "Task Scheduler").
2. Click **Create Basic Task...**
3. Name it "License Plate Monthly Report".
4. Trigger: choose **Monthly**, select the day/time you want.
5. Action: choose **Start a program**.
6. **Program/script**:
   ```
   C:\LicensePlateReporting\.venv\Scripts\python.exe
   ```
7. **Add arguments**:
   ```
   aggregate.py --config C:\LicensePlateReporting\config.json
   ```
8. **Start in (optional but recommended)**:
   ```
   C:\LicensePlateReporting
   ```
9. Finish the wizard.
10. In Task Scheduler Library, right-click the new task → **Run** to verify it executes successfully.

---

## 4. Updating later

When you receive a newer version of the project:
1. Disable the Task Scheduler task (right-click → **Disable**).
2. Replace the contents of `C:\LicensePlateReporting` with the new version, keeping your existing `config.json`.
3. Re-run:
   ```powershell
   cd C:\LicensePlateReporting
   .\.venv\Scripts\activate
   pip install -r requirements.txt
   deactivate
   ```
4. Re-enable the scheduled task.

You now have a fully automated monthly reporting setup on Windows.