# Outlook VBA Modules for Financial Planning Emails

This folder contains the VBA modules that generate your standardized
financial-planning drafts and push qualifying messages to Asana. Import the
files into Outlook so each workstation has the same behavior.

## Files

| File | Type | Purpose |
| --- | --- | --- |
| `FinancialEmailConfig.bas` | Standard Module | Single location for every value you might need to edit (CC/BCC, attachment folder, worker URL, Asana IDs, task name, etc.). Update this once per workstation; other modules read from it. |
| `financial_email_macros.bas` | Standard Module | Builds the *Financial Overview* and *Information for Planning* drafts. Shared helpers assemble the greeting, signature, body copy, and the three standard attachments stored under `ATTACHMENT_ROOT`. |
| `AsanaTaskSync.bas` | Standard Module | Determines whether a sent message should create a follow-up task and either posts to your Cloudflare Worker (`WORKER_ENDPOINT`) or directly to Asana if you supply a PAT/project/custom-field IDs—values come from `FinancialEmailConfig`. |
| *(rule-driven)* | *(n/a)* | Instead of a dedicated module, create an Outlook **after sending** rule that runs `AsanaTaskSync.RunAsanaSyncFromRule`. This replaces the older `ThisOutlookSession` event handler so you don't have to edit the special class module. |

## Importing into Outlook

1. In Outlook, open the VBA editor with **Alt+F11**.
2. Select **File → Import File…** and import the standard modules:
   - Import `FinancialEmailConfig.bas` **first** so the other modules can read
     its values.
   - Import `financial_email_macros.bas` and `AsanaTaskSync.bas` next.
3. Open `FinancialEmailConfig.bas` and fill in the constants for your CC/BCC,
   attachment folder, and either the Cloudflare Worker endpoint **or** the
   Asana PAT/project/custom-field IDs (plus the task name if you want to rename
   it).
4. Save the VBA project (**Ctrl+S**) and close the editor.
5. In Outlook, create an **After sending** rule that targets the planning
   subjects (e.g. “Financial Overview” and “Information for Planning”), choose
   **Run a script**, and select `AsanaTaskSync.RunAsanaSyncFromRule`. The rule
   fires immediately after Outlook hands the message to Exchange, sending the
   recipient and subject to your Cloudflare Worker without touching
   `ThisOutlookSession`.

### Configuration cheat sheet

| Constant | Meaning |
| --- | --- |
| `CC_RECIPIENT` | Default teammate copied on every draft. Leave blank if you do not need an automatic CC. |
| `TRACKING_BCC` | Hidden address used both for Power Automate/worker tracking and as a safeguard before syncing to Asana. |
| `ATTACHMENT_ROOT` | Folder path that stores the three planning templates. Keep the trailing `\`. |
| `WORKER_ENDPOINT` | HTTPS URL for your Cloudflare Worker. Populate this when the worker will call Asana with your credentials. |
| `ASANA_PAT` | Personal Access Token for direct API calls (leave empty when using the worker). |
| `ASANA_PROJECT_GID` | Target project ID for the follow-up task (direct API flow only). |
| `CF_EMAIL_ADDRESS_GID` | Custom field ID for “Email Address” in Asana (direct API flow only). |
| `CF_SUBJECT_GID` | Custom field ID for “Subject” in Asana (direct API flow only). |
| `TASK_NAME` | Text used for the Asana task title. |

## Daily Usage

1. Run `DraftFinancialOverviewEmail` or `DraftInformationForPlanningEmail`
   (via the VBA editor, Quick Access Toolbar, or a custom ribbon button). The
   draft opens with CC/BCC, subject, formatted HTML body, and the standard
   attachments already populated. The **To** field remains blank so you can
   enter the recipient manually.
2. When you press **Send**, Outlook evaluates the message with
   `ShouldSyncMailToAsana`. If the subject matches and the tracking BCC is still
   present, it posts the recipient and subject to your Cloudflare Worker (or
   directly to Asana if configured) so the task is created with the "Email
   Address" and "Subject" custom fields populated.
3. Review the Immediate window (`Ctrl+G` in the VBA editor) if you need to
   troubleshoot missing attachments or sync errors—log entries are prefixed with
   `[FinancialEmails]` or `[AsanaSync]`.

Repeat the same steps on every workstation: enable Outlook macros, import the
modules (config first), adjust the constants in `FinancialEmailConfig.bas`, and
verify that the three planning documents live under `ATTACHMENT_ROOT`.

## What changed compared with the earlier snippets?

You originally shared two standalone macros. The repository now includes:

- A **configuration module** so every editable value (addresses, attachment
  folder, Asana/worker settings) lives in one place instead of being scattered
  through the procedures.
- **Refactored draft builders** that reference the shared configuration and
  helper routines, but they generate the exact same Financial Overview and
  Information for Planning email bodies you provided.
- An **optional Asana sync helper** that can either call your Cloudflare
  Worker or the Asana API directly to populate the "Email Address" and
  "Subject" custom fields on the follow-up task.
- This README, which explains how to import everything into Outlook and how
  to configure the pieces per machine.

If you only wanted setup guidance, focus on the README's import steps and the
`FinancialEmailConfig.bas` constants—you do not need to modify the core VBA
logic unless your content changes. The refactor simply packages the macros and
sync routine so they are easier to reuse across multiple computers.
