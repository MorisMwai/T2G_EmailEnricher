# 📬 T2G Email Enricher (UiPath REFramework)

This is the **Queue Builder** component of the Talents2Germany email automation project. It extracts relevant metadata from incoming Outlook emails, enriches the data with parsed details, writes the information into an Excel report, and pushes each record into an Orchestrator queue for downstream handling.

Built using UiPath's Robotic Enterprise Framework (REFramework) in C#.

---

### 📌 What It Does

- ✅ Reads unread Outlook emails from the last 7 days
- ✅ Filters emails containing:
  - “I am interested in your program” *(must have CV attached)*
  - “partnership offer”
  - “join our company as an investor”
- ✅ Extracts structured metadata from the email signature:
  - Sender name
  - Company
  - Address
  - VAT ID
  - Sender email
- ✅ Appends data to an Excel report
- ✅ Adds each structured record to an Orchestrator queue (`FilteredEmailsQueue`)

---

### 🛠 Tech Stack

- UiPath Studio (REFramework – C#)
- Outlook Integration
- Regex for parsing signatures
- Excel Activities
- Orchestrator Queues & Assets

---

### 🔄 REFramework Logic

#### 1. **Init**
- Load Config.xlsx and Orchestrator assets

#### 2. **Get Transaction Data**
- Retrieve unread Gmail emails (last 7 days)
- Filter relevant emails based on keywords
- Return the filtered list for processing

#### 3. **Process Transaction**
- Extract metadata from email signature
- Append to Excel report (`PartnershipEmails_Report.xlsx`)
- Add metadata to Orchestrator Queue (`FilteredEmailsQueue`)

#### 4. **End Process**
- Close applications and dispose all open resources

---

### 📂 Orchestrator Setup

#### Queue
| Name               | Purpose                                |
|--------------------|----------------------------------------|
| `FilteredEmailsQueue` | Holds one queue item per parsed email |

#### Assets
| Name              | Type | Description                                  |
|-------------------|------|----------------------------------------------|
| `ExcelReportPath` | Text | Full path to Excel file in `Data\Output`     |
| `NotificationEmail` | Text | Summary recipient (used by Handler)         |

---

### 🧾 Output File

- **Location**: `Data\Output\PartnershipEmails_Report.xlsx`
- **Columns**: `SenderName`, `Company`, `Address`, `VatId`, `SenderEmail`

---

### 📁 Running the Bot

1. Open this project in UiPath Studio (C#)
2. Ensure Outlook is configured
3. Update `Config.xlsx` with the correct paths and values
4. Ensure assets and queue are created in Orchestrator
5. Run in attended or unattended mode
