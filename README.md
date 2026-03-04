# OfficeManagementTools

PowerShell and Google Apps Script utilities for managing bulk email in Outlook (via Microsoft Graph) and Gmail.

## Prerequisites

### Outlook Scripts (PowerShell)
- **PowerShell 5.1+** (Windows) or **PowerShell 7+** (cross-platform)
- **Microsoft.Graph.Authentication** module (auto-installed by scripts if missing)
- A Microsoft 365 / Outlook.com account
- Scripts will prompt for interactive sign-in on first run

### Gmail Script (Google Apps Script)
- A Google account with Gmail
- Access to [Google Apps Script](https://script.google.com)

---

## outlookonline/

### Find-BulkSenders.ps1

Discovers bulk/unsolicited senders in your mailbox by combining two methods: Exchange's Focused Inbox "Other" classification and keyword search for "unsubscribe". Outputs a ranked list grouped by sender, flagging which ones are already handled by `Sort-BulkMail.ps1`.

```powershell
.\Find-BulkSenders.ps1            # show top 100 bulk senders
.\Find-BulkSenders.ps1 -Top 50    # limit to top 50
```

- **Requires:** `Mail.Read`
- **Output:** `~/Desktop/bulk_senders_new.txt` (one address per line, only senders not yet in Sort-BulkMail)

### Sort-BulkMail.ps1

Sorts bulk mail into three Outlook folders -- **Bulk Newsletters**, **Bulk Ads**, and **Bulk Other** -- based on a categorized sender list. Creates the folders automatically if they don't exist.

```powershell
.\Sort-BulkMail.ps1               # move messages into folders
.\Sort-BulkMail.ps1 -DryRun       # count matches without moving anything
```

- **Requires:** `Mail.ReadWrite`
- Prompts for confirmation before moving messages
- Edit the `$categories` hashtable at the top of the script to add/remove senders

### Delete-Newsletters.ps1

Bulk-deletes newsletter and marketing emails from a flat list of sender addresses. Messages are moved to Deleted Items (recoverable), not permanently deleted.

```powershell
.\Delete-Newsletters.ps1           # delete matching messages
.\Delete-Newsletters.ps1 -DryRun   # list matches without deleting
```

- **Requires:** `Mail.ReadWrite`
- Prompts for `YES` confirmation before deletion
- Edit the `$newsletterSenders` array to customize which senders to target

### Export-Joboffers.ps1

Searches your mailbox for job-offer-related emails (offer letters, compensation details, e-sign completions, your acceptance replies) since 2019 and exports them to CSV.

```powershell
.\Export-Joboffers.ps1
```

- **Requires:** `Mail.Read`
- **Output:** `~/Desktop/job_offers_since_2019.csv`
- Runs three AQS queries (offers/agreements, your acceptances, e-sign completions), deduplicates, and includes attachment names

### Read-Emails.ps1

Fetches a hardcoded list of emails by OWA Item ID using the Graph batch API and outputs a chronological timeline file. Useful for pulling specific known messages into a readable format.

```powershell
.\Read-Emails.ps1
```

- **Requires:** `Mail.Read`
- **Output:** `~/Desktop/email_timeline.txt`
- Edit the `$encodedIds` array to specify which message IDs to fetch

---

## gmailonline/

### Sort-BulkMail-Gmail.gs

Google Apps Script equivalent of `Sort-BulkMail.ps1` for Gmail. Labels bulk mail into **Bulk Newsletters**, **Bulk Ads**, and **Bulk Other**, with optional archiving (removal from Inbox).

**Setup:**
1. Go to [script.google.com](https://script.google.com) and create a new project
2. Paste the contents of `Sort-BulkMail-Gmail.gs` into the editor
3. Select a function from the dropdown and click Run:
   - **`dryRun`** -- count matches without changing anything
   - **`sortBulkMail`** -- label and optionally archive messages
4. Approve Gmail permissions on first run
5. Check results in **View > Execution log**

- Set `ARCHIVE_AFTER_LABEL = false` at the top of the script to label without archiving
- Edit the `categories` object to add/remove senders

---

## Typical Workflow

1. **Discover** bulk senders with `Find-BulkSenders.ps1`
2. **Add** new senders to the appropriate category in `Sort-BulkMail.ps1` (and/or `Delete-Newsletters.ps1`)
3. **Sort** or **delete** the messages by running the updated script
4. Repeat periodically as new bulk senders appear
