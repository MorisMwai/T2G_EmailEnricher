# 📬 T2G Email Automation Coordinator (UiPath REFramework)

This is the **Dispatcher** component of the Talents2Germany email automation solution. It retrieves and filters Outlook emails based on business rules, extracts relevant metadata from partnership/investor inquiries, writes to Excel, and queues the structured data in Orchestrator for downstream processing.

Built using UiPath's REFramework in C#.

---

### 📌 What It Does

- ✅ Reads Outlook emails from the last 7 days
- ✅ Filters emails containing phrases like:
  - “I am interested in your program” *(must have CV attached)*
  - “partnership offer”
  - “join our company as an investor”
- ✅ Extracts metadata from email signatures (for partner/investor inquiries only):
  - Sender name
  - Company
  - Address
  - VAT ID
  - Email
- ✅ Appends extracted info to an Excel report
- ✅ Adds each valid record as a queue item in Orchestrator

---

### 🛠️ Stack

- UiPath Studio (C# – REFramework)
- Outlook Integration
- Regex for parsing signatures
- Excel file updates
- Orchestrator Queues & Assets

---

### 📂 Orchestrator Setup

#### Queue
| Name               | Purpose                                |
|--------------------|----------------------------------------|
| `FilteredEmailsQueue` | Stores filtered & parsed email metadata |

#### Assets
| Name              | Type | Purpose                                  |
|-------------------|------|------------------------------------------|
| `ExcelReportPath` | Text | Full path to Excel file in `Data\Output` |
| `NotificationEmail` | Text | Email address to send summary notification (used by Handler) |

---

### 📁 Output File

- **Location**: `Data\Output\PartnershipEmails_Report.xlsx`
- **Columns**: `SenderName`, `Company`, `Address`, `VatId`, `SenderEmail`

---

### 🔄 REFramework Flow Summary

1. **Init**  
   Load configuration and open Outlook

2. **GetTransactionData**  
   - Get recent emails  
   - Filter them  

3. **Process**  
   - Parse metadata  
   - Append to Excel  
   - Add queue items for valid records

4. **End Process**  
   Closes applications and finalizes logs

---

### 📦 Documentation

Included in the `Documentation/` folder.

---

### 🏁 Running the Bot

1. Open project in UiPath Studio (C#)
2. Ensure assets are configured in Orchestrator
3. Update `Config.xlsx` with correct paths & values
4. Run in attended or unattended mode
