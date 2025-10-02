# 📊 Alarm & TT Data Processing (2G Miss)

## 📋 Overview

This Python script processes alarm, cell, and TT (Trouble Ticket) data from Excel files to generate a comprehensive **Outage & TT Analysis Report**. The script performs data cleaning, overlap checking, time calculations, and TT status verification.

The final output helps network teams identify alarms needing TT, calculate downtime, and monitor recovery.

---

## ⚙️ Prerequisites

* **Python 3.8+**
* **Libraries**: Install via pip

  ```bash
  pip install pandas numpy sqlalchemy sympy joblib
  ```
* Input Excel files must be in **same folder as the script** or provide full paths.

---

## 📂 Input Files

The script expects three Excel files:

1. **Cell List** (`WB Name`, `WS Name` sheet)

   * Contains all cell information for sites.
   * Columns: `Cid` (Cell ID including Band), etc.

2. **Alarm List** (`WB Name`, `WS Name` sheet)

   * Columns: `NE`, `event time`, `cease time`, `NE Name`, etc.

3. **TT List** (`WB Name`, `WS Name` sheet)

   * Columns: `faultFirstOccurTime(CF_TT_Create (Create TT))`, `ResolvedFault Recovery Time(CF_TT_Process (Process TT))`, `ProblemDescription(CF_TT_Create (Create TT))`, etc.

⚠️ Make sure date/time columns are in proper datetime format in Excel.

---

## 🛠️ How It Works

1. **Data Cleaning**

   * Removes unwanted NE patterns.
   * Splits NE and Cell info into separate columns.
   * Fills missing cease times with current datetime.

2. **Overlap Handling**

   * Detects overlapping alarms for same NE/Site.
   * Merges overlapping events to a single period.

3. **Cell Count & Down Sites**

   * Calculates number of affected cells per site.
   * Checks percentage of down cells vs total cells.

4. **TT Status Calculation**

   * Calculates durations: FFOT (First Fault Occur Time) vs event, FRT (Recovery Time) vs cease.
   * Determines TT status: `Has TT`, `FFOT need to check`, `FRT need to check`, `FFOT and FRT need to check`, `Check TT`, or `Need TT`.

5. **Filtering**

   * Allows user to filter alarms by duration: enter **minutes** and choose **over/less than input**.

6. **Final Processing**

   * Removes duplicate NE/event/cease combinations.
   * Calculates percentage of down cells per site.
   * Prepares final DataFrame for reporting.

---

## ▶️ How to Run

1. Place your Excel files in the **same folder as the script** or update the paths in the code.
2. Run the script from terminal or IDE:

   ```bash
   python main.py
   ```
3. Follow the prompt to enter **number of minutes** and choose **over/less than filter**:

   * `o` → Filter alarms over the entered minutes.
   * `l` → Filter alarms less than the entered minutes.

---

## 📊 Output

The script generates an Excel file: `Resultss.xlsx` with the following sheets:

1. **Alarms** – Cleaned and processed alarms data.
2. **TT** – Processed Trouble Ticket information.
3. **Cell** – Cleaned cell list with site and band info.
4. **Result** – Final merged report including:

   * Duration of downtime per alarm
   * Down cell count
   * TT Status
   * Percentage of down cells per site

All output is saved in the **same folder as the script** by default.

---

## ⚠️ Notes

* Ensure **datetime columns in Excel** are correctly formatted.
* Script fills missing cease/recovery times with the **current datetime**.
* TT Status calculation depends on FFOT/FRT timings and threshold checks.
* Filtering is **interactive**, user must provide input in console.
* Duplicate alarms are merged using overlap logic.
* Percentage of down cells is calculated per site for monitoring.

---

## 🖊️ Author

Mohammad Rahmani
📧 [mhrs1995@gmail.com](mailto:mhrs1995@gmail.com)
