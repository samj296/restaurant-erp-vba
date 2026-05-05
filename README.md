# Overview

This project is a complete restaurant operations system built using:

* Excel VBA as the front‑end application
* Three Access databases as the backend
* ADODB for data connectivity
* UserForms for UI
* Automated billing, income, expense, salary, and reporting workflows
<br><br>
This system was created during the COVID‑19 lockdown to run a real restaurant (Ala Rasi Foods & Beverages Pvt. Ltd.). It handled daily operations for over a year, including billing, menu management, staff attendance, income/expense tracking, and multi‑brand menu support.
<br><br>
Although built while learning VBA, the system evolved into a fully functional ERP‑style application with real business impact.

# Application Startup Behavior (Important)
When the workbook opens, the Excel window is automatically hidden and the application launches in App Mode. This is intentional and part of the original design.
<br><br>
    * The startup sequence works like this:
<br>
    * The Excel workbook hides itself
<br>
    * The login form appears immediately
<br>
    * The user interacts only with the application UI (UserForms), not the Excel interface
<br>
    * This behavior is controlled by the following logic inside the workbook:

```vb
Private Sub Workbook_Open()
    ThisWorkbook.Application.Visible = False
    Frm_LogIn.Show
End Sub
```
Because of this:
<br>
    * You will not see the Excel grid when the file opens
<br>
    * The system behaves like a standalone application
<br>
    * Closing the app triggers a controlled shutdown via Quit_App
<br>
This is expected and part of the intended user experience.

# Key Features
1. Multi‑Database Architecture
The system uses three separate Access databases, each with a dedicated purpose:

     1. User & Configuration Database
<br>
        * Usernames & passwords
<br>
        * Salary tables
<br>
        * Master configuration values
<br>
        * Login/logout tracking
<br><br>
     2. Menu Database
<br>
        * Arabian Hub franchise menu
<br>
        * Chinese cuisine menu
<br>
        * Multi‑brand support
<br>
        * Item categories & pricing
<br>
<br>
    3. Billing & Transactions Database
<br>
        * Bill records
<br>
        * Restaurant‑specific separation
<br>
        * Daily income updates
<br>
        * Expense entries
<br>
        * Sales analytics
<br><br>

# Architecture Diagram
Code
```text


+-------------------------+           +-------------------------+
|     Excel Front-End     | -------→  |   Access DB: Menu       |
|  (VBA UserForms + UI)   |  ADODB    | (Multi-brand Menu Data) |
+-----------+-------------+           +-------------------------+
            |               \
            | ADODB          \ ADODB
            ↓                  ↘
+-------------------------+     +-------------------------+
|   Access DB: Users      |     |   Access DB: Billing    |
|  (Auth + Salary + Config)     | (Bills + Income + Exp)  |
+-------------------------+     +-------------------------+

```

# Major Modules
🔹 Authentication
<br>
    * Login form
<br>
    * Logout timestamp tracking
<br>
    * Admin vs non‑admin permissions
<br>
    * User switching
<br><br>
🔹 Billing System
<br>
    * Menu selection (multi‑brand)
<br>
    * Auto‑calculated totals
<br>
    * Bill saving to database
<br>
    * Automatic daily income update
<br><br>
🔹 Income & Expense Tracking
<br>
    * Daily income dashboard
<br>
    * Expense entry form
<br>
    * Category‑based reporting
<br><br>
🔹 Salary Management
<br>
    * Staff salary setup
<br>
    * Monthly salary calculations
<br>
    * Attendance integration
<br><br>
🔹 Dashboard
    * Total sales chart
    * Attendance list
    * Quick navigation buttons
    * Admin‑only controls

# Tech Stack

|Layer	|Technology |
|-------|-----------|
|Front-End|	Excel VBA (UserForms, Modules, Charts) |
|Backend|	Microsoft Access (3 databases) |
|Connectivity|	ADODB (OLEDB 12.0 Provider)|
|Reporting |	Excel Charts + Custom Dashboards|
|File Handling |	Dynamic folder selection, path validation |

# How to Run the Project
This project uses Excel VBA as the front‑end and three Access databases as the backend.
Because the system was originally built for a single‑machine environment, the first‑time setup requires selecting the folder where the databases are stored.
<br>
Follow the steps below to run the demo version.
<br>
1️) Download the Required Files
<br>
Download the following from this repository:
<br>
    * Excel Application
<br>
        `/dist/RestaurantSystem_Demo.xlsm`
<br>
    * Demo Databases
```shell
Located in /database/:
Users_Demo.accdb
Menu_Demo.accdb
Billing_Demo.accdb
```

2️)  Place All Files in the Same Folder
Create a folder anywhere on your computer and place:
<br>
* the .xlsm file
<br>
* all three .accdb files
<br>
in that same folder.
The system relies on this folder structure.

3️) Open the .xlsm File
When you open the workbook:
* Excel will prompt you to enable macros
* The application will launch in App Mode
* The login form will appear automatically

4️) First‑Time Setup (Important)
When you attempt to log in for the first time, the system will:
<br>
    1) Check the stored file path inside the workbook
<br>
    2) Detect that the path is missing or invalid (expected in the demo)
<br>
    3) Prompt you to select the folder where the databases are located

⚠️ You MUST select a folder
Canceling is **not allowed** due to the original design.
<br>
The dialog will continue to appear until a folder is selected.
<br><br>
This only happens once.
<br>
After selecting the folder, the system updates the paths automatically.

5️) Login Credentials (Demo)
Use the following credentials to log in:

Code
```shell
Username: admin
Password: admin123
```

These credentials are for the demo version only.
<br>

6️) Database Password (Demo)
<br>
All three Access databases use the same demo password:
<br><br>

```shell
admin123
```

This password is intentionally simple for testing and demonstration.

7️) VBA Project Password (Demo)
The VBA project inside the .xlsm file is unlocked using:

```shell
admin123
```

This password is included intentionally so reviewers can:
<br>
    * inspect the source code
<br>
    * explore the UserForms
<br>
    * review the ADODB logic
<br>
    * understand the architecture

All sensitive business data has been removed.

8️) Explore the System
Once logged in, you can explore:
    * Billing module
<br>
    * Menu management
<br>
    * Income tracking
<br>
    * Expense tracking
<br>
    * Salary module
<br>
    * Dashboard & charts
<br>
    * Attendance list
<br>
    * Admin‑only features