# Point Click - RPA Automation

VBA-based RPA that automates data entry in a Citrix application by simulating mouse clicks and keyboard input from Excel data.

## Overview

This project uses Windows API calls (`SetCursorPos` and `mouse_event`) to control the mouse and perform a complete data entry flow inside a Citrix environment.

## Problem

Manually entering data from Excel into a Citrix system is repetitive, slow and prone to human error.

## Solution

The macro:

- Reads data from Excel columns
- Simulates mouse clicks on specific screen coordinates
- Types the required values
- Copies the generated process number back to Excel
- Repeats the process for all rows

## Important Notes

- This automation depends on **exact screen coordinates**
- Resolution and window position must be the same every time
- Do not touch the mouse or keyboard while the macro is running
- The Citrix window must be in the foreground

## How to Use

1. Open the Excel file with the data
2. Press `Alt + F11` and import `src/Point_Click.bas`
3. Make sure the Citrix application is open and visible
4. Run the macro `Point_Click`
5. Do not touch the mouse/keyboard until it finishes

## Data Structure

| Column | Content              |
|--------|----------------------|
| B      | HAWB                 |
| E      | DR                   |
| G      | ETD                  |
| H      | ETA                  |
| I      | Ref Cliente          |
| J      | Agente               |
| C      | Process Number (output) |

## Warning

This type of automation is **fragile**. Any change in screen resolution, Citrix window size or application layout will break the coordinates.
