# 📬 T2G Email Enricher (UiPath REFramework)

This is the **Queue Builder (Dispatcher)** component of the Talents2Germany email automation project.  
It retrieves relevant Gmail emails, extracts structured metadata from senders, writes the information to an Excel report, and queues the data in Orchestrator for further handling by the **Handler (Performer)**.  
Additionally, it moves successfully processed emails into designated folders for better mailbox organization.

Built using UiPath's Robotic Enterprise Framework (REFramework) in C#.

---

### 📌 What It Does

- ✅ Reads unread Gmail emails from the last 7 days
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
- ✅ Appends extracted data to an Excel report
- ✅ Saves email attachments (e.g., CVs) to:
  `Data\Output\Attachments
- ✅ Adds the following to **one Orchestrator queue (`FilteredEmailsQueue`)**:
  - **One item per processed email** containing all metadata
  - **Summary data** in a separate queue item:
    - `ProcessedEmails`: total count of partnership/investor emails
    - `ProcessedEmailsWithAttachments`: total count of CV emails
- ✅ Moves processed emails into configured folders:
  - **ProcessedEmails**
  - **ProcessedEmailsWithAttachments**

---

### 🛠 Tech Stack

- **UiPath Studio** (REFramework – C#)
- Gmail Integration
- Regex for parsing email signatures
- Excel Activities
- Orchestrator Queues & Assets
  
---

### 🔄 REFramework Logic

#### 1. **Init**
- Load `Config.xlsx` and Orchestrator assets

#### 2. **Get Transaction Data**
- Retrieve unread Gmail emails from the last 7 days
- Filter based on keywords for:
  - Partnership/Investor emails
  - Program interest emails (with CV)
- Return filtered email lists for processing

#### 3. **Process Transaction**
- For each filtered email:
  - Extract metadata (name, company, VAT, etc.)
  - Save attachments (if any) to `Data\Output\Attachments`
  - Append to Excel report (`PartnershipEmails_Report.xlsx`)
- After loop:
  - Add all processed records to `FilteredEmailsQueue`
  - Add a summary queue item containing:
    - `ProcessedEmails`
    - `ProcessedEmailsWithAttachments`
  - Move processed emails to designated folders

#### 4. **End Process**
- Close all resources and log summary

---

### 📂 Orchestrator Setup

#### Queue
| Name                 | Purpose                                         |
|----------------------|-------------------------------------------------|
| `FilteredEmailsQueue` | Stores processed email records and summary data |

#### Assets
| Name                      | Type | Description                                   |
|---------------------------|------|-----------------------------------------------|
| `ExcelReportPath`         | Text | Path to Excel file in `Data\Output`          |
| `NotificationEmail`       | Text | Email for sending summary (used by Handler)  |
| `ProcessedEmailsFolder`   | Text | Gmail folder for processed partnership emails |
| `ProcessedAttachmentsFolder` | Text | Gmail folder for emails with CV attachments  |

---

### 🧾 Output File

- **Path**: `Data\Output\PartnershipEmails_Report.xlsx`
- **Sheet Name**: `Partnership_Emails`
- **Columns**: `DateProcessed`, `SenderName`, `Company`, `Address`, `VatId`, `SenderEmail`, `EmailSubject`

- **Attachments**:  
  Path: `Data\Output\Attachments`

---

### 📁 Running the Bot

1. Open this project in UiPath Studio (C#)
2. Ensure Gmail integration is configured
3. Update `Config.xlsx` and Orchestrator assets with correct paths and credentials
4. Set up queues in Orchestrator as outlined above
5. Run the bot in attended or unattended mode
   
---

✅ Next: The **Performer** (Handler) will:
- Read `FilteredEmailsQueue`
- Send summary email using summary queue item
- Update Orchestrator transaction statuses
