# VBA Work Schedule – Private Leave & Overtime Tracking

## 📌 Project Overview

This project is a VBA automation created for an Excel-based employee work schedule.

The workbook allows employees to record their **working hours**, as well as **private leave** and **overtime worked**.

The main goal of the VBA macro is to automatically calculate the total time spent on private leave and the total overtime worked by each employee within a given month.

The solution was designed to work with a **dynamic number of employees and days**, without relying on a fixed number of columns or rows.

---

## 🎯 Project Goal

The VBA macro automatically adds two summary columns for each employee:

* **Private Leave** – total duration of private leave during the month
* **Overtime** – total overtime worked during the month

The calculation is performed separately for each employee and covers the entire month.

The workbook assumes that working-time settlement is performed on a **monthly basis**.

---

## ⏱️ Input Format

The expected format for private leave or overtime is:

```text
15:45-16:45
```

The macro also supports **two private leave / overtime entries within the same day**.

For example:

```text
12:05-12:50, 16:15-16:50
```

or:

```text
13:15-14:15 / 16:15-16:45
```

The macro recognizes both comma `,` and slash `/` as separators.

---

## 🛡️ Error Handling

The macro was designed to handle common data-entry mistakes that can occur when employees enter their working hours manually.

### Supported input errors

For example:

```text
16:00:16:45
```

instead of the correct:

```text
16:00-16:45
```

The macro can identify and process this type of common formatting mistake.

Other invalid or unrecognized entries are treated as an **Excel error value**, rather than causing the macro to stop working.

---

## 🔄 Dynamic Workbook Structure

One of the main requirements was that the VBA code should **not depend on a fixed workbook structure**.

The macro works correctly when:

* 👥 the number of employees changes,
* 📅 the number of days in the month changes,
* 📊 the number of employee-related columns changes,
* 🗓️ the month has a different number of days.

This means the code does not assume a predefined number of rows or columns.

---

## 📅 Example Workbook

The example workbook contains four monthly sheets:

* January
* February
* March
* April

Each sheet contains a sample employee schedule that can be used to test the macro.

A button for running the macro is located in **cell A1**.

Simply open a monthly sheet and click the button to run the VBA automation.

---

## 📂 Project Files

The repository contains several files for different purposes:

### `Grafik_VBA.xlsm`

Example workbook with the VBA macro included.

It can be used to **test the automation** and see how the macro works with the example schedule.

### `vba_grafik.txt`

The VBA code used to create the automation.

The code is provided separately so that it can be reviewed without opening the macro-enabled workbook.

### `Grafik.xlsx`

The same example schedule **without the VBA macro**.

This version is provided for reference and allows the workbook structure and example data to be reviewed without running any code.

---

## 🧪 How to Test

1. Open the macro-enabled Excel workbook -> Grafik_VBA.xlsm.
2. Select one of the monthly sheets.
3. Review the example employee schedule.
4. Click the macro button located in **cell A1**.
5. Check the automatically calculated **Private Leave** and **Overtime** columns.

You can also modify the example entries to test different scenarios, including:

* standard private leave,
* standard overtime,
* two entries in one day,
* comma-separated entries,
* slash-separated entries,
* incorrectly formatted time ranges,
* different numbers of employees,
* different numbers of days.

---

## ⚠️ Important

**The macro action cannot be undone using Excel's standard Undo function.**

For safe testing, it is recommended to:

> **Make a copy of the worksheet before running the macro.**

This allows the original data to be preserved and the macro to be tested multiple times.

---

## 💡 VBA Skills Demonstrated

This project demonstrates practical VBA skills including:

* Excel VBA automation
* Dynamic range detection
* Working with variable numbers of rows and columns
* Processing and parsing text values
* Working with time values and durations
* Handling multiple entries within a single cell
* Input validation
* Error handling
* Automating Excel worksheet modifications
* Working with multiple worksheets
* Creating and using macro buttons
* Designing VBA code around a real-world Excel workflow

## Author

**Dominika Schabowska**
