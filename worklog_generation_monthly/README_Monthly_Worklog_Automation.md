# Monthly Worklog Automation Workflow

## Overview
This n8n workflow automates the complete monthly worklog reporting process for:
- Technical Leads (TLs)
- Non-TL Employees

The workflow:
- Fetches worklog data from API
- Processes employee work hours
- Calculates leave and missing hours
- Generates formatted Excel reports
- Applies Python-based Excel styling
- Sends automated email reports

---

# Main Workflow Architecture

```mermaid
flowchart LR

A[Webhook Trigger]
--> B[Set Date Range]
--> C[Fetch Worklog API Data]

C --> D[TL Processing Branch]
C --> E[Employee Processing Branch]

D --> F[Filter Technical Leads]
F --> G[Generate TL Excel]
G --> H[Apply Python Formatting]
H --> I[Send TL Email]

E --> J[Filter Employees]
J --> K[Generate Employee Excel]
K --> L[Apply Python Formatting]
L --> M[Send Employee Email]

style A fill:#f9f,stroke:#333
style C fill:#bbf,stroke:#333
style H fill:#bfb,stroke:#333
style L fill:#bfb,stroke:#333
```

---

# Workflow Nodes

## 1. Webhook Trigger
### Node
`Webhook1`

### Purpose
Triggers the workflow through an API endpoint.

---

## 2. Set Date Range
### Node
`Edit Fields1`

### Purpose
Defines monthly reporting date range.

### Example
```json
{
  "FromDate": "01-Mar-2026",
  "ToDate": "31-Mar-2026"
}
```

---

## 3. Fetch Worklog Data
### Node
`HTTP Request1`

### Purpose
Retrieves worklog data from external API.

### API Endpoint
```plaintext
https://upaygoa.com/JiraAPILIVE/API/FetchData/GetReportData
```

---

# API Parameters

| Parameter | Value |
|---|---|
| ReportId | 0 |
| SubReportId | 19 |
| ProjId | 0 |
| UserKey | U2755 |

---

# Processing Branches

The workflow splits into:
1. TL Monthly Worklog Branch
2. Employee Monthly Worklog Branch

---

# Technical Lead Processing Flow

```mermaid
flowchart TD

A[API Response]
--> B[Process Monthly Data]

B --> C[Identify TL Employees]
C --> D[Calculate Working Days]

D --> E[Calculate Total Logged Hours]
E --> F[Calculate Leave Hours]

F --> G[Calculate Actual Hours]
G --> H[Calculate Hours Not Logged]

H --> I[Generate TL Excel]
I --> J[Apply Excel Formatting]

J --> K[Send Email]
```

---

# Employee Processing Flow

```mermaid
flowchart TD

A[API Response]
--> B[Process Monthly Data]

B --> C[Exclude TL Employees]
C --> D[Calculate Working Days]

D --> E[Calculate Total Logged Hours]
E --> F[Calculate Leave Hours]

F --> G[Calculate Actual Hours]
G --> H[Calculate Hours Not Logged]

H --> I[Generate Employee Excel]
I --> J[Apply Excel Formatting]

J --> K[Send Email]
```

---

# Core Business Logic

## Employee Grouping

The workflow groups records using:

```plaintext
EmployeeName + Designation
```

---

# Dynamic Date Generation

## Daily Columns

Each unique worklog date becomes a separate report column.

### Example
```plaintext
01/03/2026
02/03/2026
03/03/2026
```

---

# Technical Lead Logic

The workflow contains a predefined list of Technical Leads.

### TL Weekend Rules
Technical Leads receive:
- Saturday Off
- Sunday Off

---

## Non-TL Weekend Rules

Employees receive:
- Sunday Off only

---

# Holiday Handling

Predefined holidays are excluded from working day calculations.

### Example Holidays
```plaintext
2025-10-02
2025-10-20
```

---

# Worklog Calculation Logic

## Total Logged Hours

```math
TotalLogged = \sum DailyLoggedHours
```

---

## Working Hours Formula

```math
WorkingHours = WorkingDays \times 7
```

---

# Leave Calculation

The workflow checks:

```plaintext
LeaveTakenYn
```

Accepted leave indicators:
- 1
- true
- y
- yes

---

# Leave Hours Formula

```math
LeaveHours = LeaveDays \times 8
```

---

# Actual Hours Formula

```math
ActualHours = WorkingHours - LeaveHours
```

---

# Hours Not Logged Formula

```math
HoursNotLogged = ActualHours - TotalLogged
```

---

# Days Not Logged Formula

```math
DaysNotLogged = \frac{HoursNotLogged}{8}
```

Rounded to nearest integer.

---

# Unicode & Data Cleanup

The workflow:
- Removes hidden characters
- Cleans invalid keys
- Normalizes employee names

---

# Employee Report Structure

Generated report fields:

| Field | Description |
|---|---|
| EmployeeName | Employee name |
| Designation | Employee role |
| Daily Columns | Logged daily hours |
| Total Logged | Total logged hours |
| Working Hours | Expected hours |
| Leave | Leave count |
| Leave Hours | Leave deduction |
| Actual Hours | Expected working hours |
| Hours Not Logged | Missing hours |
| Days Not Logged | Missing work days |

---

# Excel Generation Flow

## Convert JSON to XLSX
### Nodes
- `Convert to File6`
- `Convert to File7`

### Purpose
Converts processed JSON into Excel files.

---

# Python Excel Formatting

## Formatter Script
```plaintext
color_excel.py
```

### Responsibilities
- Header coloring
- Conditional formatting
- Borders
- Cell styling
- Excel beautification

---

# Generated Report Paths

## TL Report
```plaintext
C:/Users/Shreya gavli/Desktop/worklog/Tl_monthly/formatted_worklog.xlsx
```

---

## Employee Report
```plaintext
C:/Users/Shreya gavli/Desktop/worklog/Employee_monthly/formatted_worklog_emp.xlsx
```

---

# Email Automation

## TL Report Email

### Subject
```plaintext
TL monthly worklog
```

---

## Employee Report Email

### Subject
```plaintext
Employee monthly worklog
```

---

# Gmail Delivery Flow

```mermaid
flowchart TD

A[Formatted Excel Report]
--> B[Read File from Disk]

B --> C[Attach File]
C --> D[Compose Email]

D --> E[Send Gmail Message]
```

---

# Complete End-to-End Workflow

```mermaid
flowchart TD

A[Webhook Trigger]
--> B[Set Date Range]

B --> C[Fetch API Worklog Data]

C --> D[Extract Employee Records]
D --> E[Generate Dynamic Date Columns]

E --> F[Group Employees]
F --> G[Calculate Logged Hours]

G --> H[Calculate Working Hours]
H --> I[Calculate Leave]

I --> J[Calculate Actual Hours]
J --> K[Calculate Hours Not Logged]

K --> L{Is Technical Lead?}

L -- Yes --> M[Generate TL Report]
L -- No --> N[Generate Employee Report]

M --> O[Convert TL JSON to Excel]
N --> P[Convert Employee JSON to Excel]

O --> Q[Apply Python Formatting]
P --> R[Apply Python Formatting]

Q --> S[Send TL Email]
R --> T[Send Employee Email]
```

---

# Key Features

## Automated Monthly Reporting
Generates:
- TL reports
- Employee reports

Automatically.

---

## Smart Worklog Analysis
Calculates:
- Logged hours
- Leave hours
- Missing hours
- Missing work days

---

## Dynamic Date Handling
Automatically creates:
- Daily worklog columns
- Monthly structures

---

## Intelligent Employee Classification
Separates:
- Technical Leads
- Employees

Using predefined logic.

---

## Excel Beautification
Automatically applies:
- Formatting
- Styling
- Highlighting
- Conditional colors

Using Python automation.

---

# Technologies Used

| Technology | Purpose |
|---|---|
| n8n | Workflow orchestration |
| JavaScript | Data transformation |
| Python | Excel formatting |
| Gmail API | Automated email delivery |
| HTTP API | Worklog data retrieval |
| XLSX Conversion | Excel generation |

---

# Workflow Components

## Input
```plaintext
Worklog API Data
```

---

## Processing
```plaintext
JavaScript Logic
```

---

## Formatting
```plaintext
Python Excel Formatter
```

---

## Delivery
```plaintext
Gmail Automation
```

---

# Expected Final Outputs

The workflow produces:

## 1. Technical Lead Monthly Report
Contains:
- Daily logged hours
- Leave calculations
- Missing work hours
- Missing work days

---

## 2. Employee Monthly Report
Contains:
- Worklog summaries
- Leave tracking
- Actual working hours
- Productivity analysis

---

## 3. Formatted Excel Files
Professionally formatted reports with:
- Color coding
- Highlighting
- Structured layouts

---

## 4. Automated Email Reports
Monthly reports automatically emailed with Excel attachments.
