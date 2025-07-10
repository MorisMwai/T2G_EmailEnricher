# 🤖 T2G Email Automation Bot (UiPath REFramework)

This is a test RPA project built using UiPath's Robotic Enterprise Framework (REFramework) in C#. It automates the extraction of specific emails from Outlook, parses sender details from their email signature, logs the structured data into an Excel report, and sends a summary notification email.

This project is designed to showcase best practices in transactional processing, queue integration, exception handling, and automation design using UiPath Studio.

---

### 📌 Use Case Summary

- Retrieve Outlook emails from the last 7 days.
- Filter emails that contain specific phrases like:
  - “I am interested in your program” *(with CV attached)*
  - “partnership offer”
  - “join our company as an investor”
- For partnership/investor emails:
  - Extract sender name, company, address, VAT ID, and email from the signature.
  - Append data to an Excel report in `Data\Output`.
- Send a notification email summarizing how many messages were processed.

---

### 🛠 Tech Stack

- UiPath Studio (REFramework in C#)
- Outlook Activities
- Excel Automation
- Regex for signature parsing
- Orchestrator Queues and Assets

---

### Documentation is included in the `Documentation` folder ###

---

### REFramework Template ###
**Robotic Enterprise Framework**

* Built on top of *Transactional Business Process* template  
* Uses *State Machine* layout for the phases of automation project  
* Offers high-level logging, exception handling and recovery  
* Keeps external settings in *Config.xlsx* file and Orchestrator assets  
* Pulls credentials from Orchestrator assets and *Windows Credential Manager*  
* Gets transaction data from Orchestrator queue and updates back status  
* Takes screenshots in case of system exceptions  

---

### How It Works ###

1. **INITIALIZE PROCESS**  
   - `./Framework/InitAllSettings` – Load configuration data from `Config.xlsx` and assets  
   - `./Framework/InitAllApplications` – Open and log in to required applications (Outlook in this case)

2. **GET TRANSACTION DATA**  
   - `./Framework/GetTransactionData` – Fetch emails using Outlook activities and filter based on keywords  
   - Filtered emails are pushed to the `FilteredEmailsQueue` for processing

3. **PROCESS TRANSACTION**  
   - `Process.xaml` – Extract structured data from email signature using Regex  
   - Append the data to `PartnershipEmails_Report.xlsx`  
   - `./Framework/SetTransactionStatus` – Mark queue items as Success or Failed

4. **END PROCESS**  
   - `./Framework/CloseAllApplications` – Close Outlook and other open apps

---

### Orchestrator Setup

#### Queue:
- **Name:** `FilteredEmailsQueue`
- **Purpose:** Store emails that meet the filter criteria (1 queue item = 1 email)

#### Assets in Use:
| Name              | Type | Description                              |
|-------------------|------|------------------------------------------|
| NotificationEmail | Text | Recipient of the summary notification    |
| ExcelReportPath   | Text | Path to the Excel file for report output |

---

### For New Project

1. Check the `Config.xlsx` file and add/customize any required fields and values.
2. Implement `InitAllApplications.xaml` and `CloseAllApplications.xaml`, and reference them in `Config.xlsx` if needed.
3. Modify `GetTransactionData.xaml` and `SetTransactionStatus.xaml` to align with your Outlook-based filtering and queue logic.
4. Implement logic in `Process.xaml` to parse email signatures and append the data to Excel.

---

### Running the Bot

1. Open the solution in UiPath Studio (C#).
2. Configure your Outlook and Excel environment.
3. Set up assets and queue in Orchestrator.
4. Run the project via Studio or Orchestrator.
5. Review logs and Excel outputs in the `Data\Output` folder.

---
