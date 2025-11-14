Attribute VB_Name = "FinancialEmailConfig"
Option Explicit

' Central place to store all of the editable values for the Outlook macros and
' the Asana sync helper. Update these constants once on each workstation after
' importing the modules, then leave the other code files untouched.

' --- Email addressing ---
Public Const CC_RECIPIENT As String = "abyers@reliantwealthstrategies.com"
Public Const TRACKING_BCC As String = "CT-20250601-A@hidden.local"

' --- Attachments ---
' Folder that holds Networth_Template.xlsx, Financial_Planning_Monthly_Budget.xlsx,
' and ES.DBA.FINANCIALPLANNING.pdf. Keep the trailing backslash so filenames
' concatenate correctly.
Public Const ATTACHMENT_ROOT As String = _
    "C:\\Users\\alysi\\Documents\\Financial Planning Docs\\"

' --- Asana/Worker integration ---
' If you are posting to a Cloudflare Worker, paste the HTTPS endpoint below and
' leave the Asana values blank. The worker is responsible for calling the Asana
' API with your PAT and mapping the custom fields.
Public Const WORKER_ENDPOINT As String = ""

' Optional direct-Asana fallback. Only fill these in if you plan to post to the
' Asana API straight from Outlook (the worker-first approach is safer because it
' keeps credentials out of VBA).
Public Const ASANA_PAT As String = ""
Public Const ASANA_PROJECT_GID As String = ""
Public Const CF_EMAIL_ADDRESS_GID As String = ""
Public Const CF_SUBJECT_GID As String = ""

' Task metadata applied regardless of which sync path you choose.
Public Const TASK_NAME As String = "Follow-Up: Financial Planning Email Sent"

