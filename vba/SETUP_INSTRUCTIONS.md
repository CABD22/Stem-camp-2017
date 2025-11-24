# Financial Planning Email Automation Setup Guide

This guide walks through every step required to deploy the Outlook VBA macros and
Cloudflare Worker integration that creates Asana follow-up tasks when you send
your "Financial Overview" or "Information for Planning" emails. Follow each
section in order on every workstation that should have the fast-button
experience.

## 1. Gather the project files

1. Download or clone this repository onto the workstation.
2. Locate the VBA modules under the `vba/` folder:
   - `FinancialEmailConfig.bas`
   - `financial_email_macros.bas`
   - `AsanaTaskSync.bas`
3. Copy these files to a convenient local folder. You can import them directly
   from this location inside Outlook.

> **Tip:** If you cannot download files, open each `.bas` file, copy its
> contents, and paste them into a new module in Outlook. Name the modules
> exactly as shown above.

## 2. Enable macros and programmatic access in Outlook

1. In Outlook, go to **File → Options → Trust Center → Trust Center Settings**.
2. Under **Macro Settings**, select **Notifications for all macros** or the
   lowest setting your policy allows.
3. Select **Programmatic Access** and allow programmatic access so Outlook does
   not prompt you each time the macro prepares or sends an email.
4. Click **OK** to apply the changes.

## 3. Import the VBA modules

1. Press **Alt+F11** to open the Outlook VBA editor.
2. In the Project pane, highlight `Project1 (VbaProject.OTM)`.
3. Choose **File → Import File…** and import the modules **in this order**:
   1. `FinancialEmailConfig.bas` – provides shared configuration values.
   2. `financial_email_macros.bas` – builds the email drafts.
   3. `AsanaTaskSync.bas` – posts qualifying sends to your Cloudflare Worker or
      directly to Asana.
4. Press **Ctrl+S** to save the VBA project when finished.

## 4. Configure shared settings

1. In the VBA editor, open the `FinancialEmailConfig` module.
2. Update each constant so it matches your environment:
   - `CC_RECIPIENT` – teammate copied on every draft (leave blank if not used).
   - `TRACKING_BCC` – hidden address already used for automation tracking.
   - `ATTACHMENT_ROOT` – folder that stores `Networth_Template.xlsx`,
     `Financial_Planning_Monthly_Budget.xlsx`, and `ES.DBA.FINANCIALPLANNING.pdf`.
     Keep the trailing backslash (e.g., `"C:\\Users\\alysi\\Documents\\Financial Planning Docs\\"`).
   - `WORKER_ENDPOINT` – HTTPS URL for your Cloudflare Worker. Leave blank if
     you plan to call Asana directly from Outlook.
   - `TASK_NAME` – defaults to `"Follow-Up: Financial Planning Email Sent"`.
   - If you **will not** use the worker, fill in `ASANA_PAT`,
     `ASANA_PROJECT_GID`, `CF_EMAIL_ADDRESS_GID`, and `CF_SUBJECT_GID` for the
     direct Asana API flow instead.
3. Confirm the attachment files live under `ATTACHMENT_ROOT` on this
   workstation. Copy them if necessary.

## 5. (Optional) Configure the Cloudflare Worker

Skip this section if you already deployed the worker and only need to reuse it.

1. Open the worker source you provided (for example,
   `fp-asana-worker.alysiamarie2010.workers.dev`).
2. Ensure it expects a JSON payload with `email`, `name`, `subject`, and
   `date_sent` fields—your existing script already does.
3. Use `wrangler secret put ASANA_PAT` (or the Cloudflare dashboard) to store
   your Asana Personal Access Token securely.
4. Deploy the worker with `wrangler deploy` and note the public HTTPS URL. This
   is the value you place in `WORKER_ENDPOINT` inside Outlook.
5. If you have not already configured SendGrid Inbound Parse, follow the
   earlier instructions from our conversation to route BCC copies to the worker.

## 6. Create the Outlook after-sending rule

1. In Outlook, open **File → Manage Rules & Alerts**.
2. Click **New Rule…** → **Apply rule on messages I send**.
3. Select **with specific words in the subject** and add the exact subjects:
   - `Financial Overview`
   - `Information for Planning`
4. Optionally add additional checks (e.g., stop processing more rules).
5. In the actions list, check **run a script** and choose
   `AsanaTaskSync.RunAsanaSyncFromRule`.
6. Finish the wizard and apply the rule.

When the rule fires, it sends the message metadata to your Cloudflare Worker (or
calls the Asana API directly if the worker URL is blank and the PAT is present).

## 7. Test end-to-end

1. In the VBA editor (or via a Quick Access Toolbar button), run
   `DraftFinancialOverviewEmail`.
2. Enter a test recipient in the **To** field and click **Send**.
3. Confirm the three standard attachments are included and the CC/BCC values are
   correct.
4. Watch the Immediate window (`Ctrl+G` in the VBA editor) for `[AsanaSync]`
   logs—successful runs show the payload and response status.
5. Verify that a task named `Follow-Up: Financial Planning Email Sent` appears
   in Asana with the recipient email stored in the "Email Address" custom field
   and the subject stored in the "Subject" custom field. The worker also stores
   the date in your custom date field.

## 8. Roll out to additional workstations

Repeat Sections 1–7 on every machine that needs the macros:

- Copy the three `.bas` files locally.
- Enable macros/programmatic access.
- Import the modules in the same order.
- Update `FinancialEmailConfig` to match that computer's attachment path.
- Ensure the after-sending rule is active.

Once complete, all computers will produce identical drafts and will create the
matching Asana follow-up task whenever one of the fast-button emails is sent.

## 9. (Optional) Add Quick Access Toolbar buttons

If you want visible buttons in Outlook for the two drafts:

1. Go to **File → Options → Quick Access Toolbar**.
2. Set “Choose commands from” to **Macros**.
3. Add `Project1.DraftFinancialOverviewEmail` and/or
   `Project1.DraftInformationForPlanningEmail` to the toolbar.
4. Click **Modify…** to pick an icon and friendly display name, then press
   **OK** to save.

You can now generate each draft with a single click from the Outlook window.
