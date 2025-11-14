Attribute VB_Name = "FinancialEmailMacros"
Option Explicit

' Centralised helper routines for generating planning-related Outlook drafts.
' These procedures mirror the snippets shared in chat but factor the repeated
' logic (greeting, signature, attachment handling) into reusable functions so
' the code is easier to maintain and extend.
'
' Addresses, attachment folders, and integration settings now live in
' FinancialEmailConfig.bas so you only have to edit them in one place after
' importing the modules on a new workstation. No default "To" recipient is
' enforced so the draft opens with an empty address box unless the caller
' provides one explicitly.

Private Function BuildGreeting() As String
    Dim currentHour As Integer
    currentHour = Hour(Now)

    Select Case currentHour
        Case Is < 12
            BuildGreeting = "Good morning,"
        Case Is < 17
            BuildGreeting = "Good afternoon,"
        Case Else
            BuildGreeting = "Good evening,"
    End Select
End Function

Private Function BuildSignatureBlock() As String
    BuildSignatureBlock = "<div style='margin-top: 25px;'><img src='https://i.imgur.com/ZV5VpdX.png' alt='Reliant Business & Wealth Strategies Signature' width='400' style='float: left; margin-right: 20px; margin-bottom: 20px;'></div>" & _
        "<div style='clear: both; margin-top: 0px; margin-bottom: 15px; margin-left: 50px;'><p style='font-family: ""Aptos (Body)"", Arial, sans-serif; font-size: 14pt; font-weight: bold; margin-bottom: 5px; text-align: left;'>" & _
        "<a href='https://calendly.com/prestonbyers/financial-topic-discussion' style='color: #0066cc; text-decoration: underline;'>BOOK A CONSULTATION</a></p></div>" & _
        "<hr style='width: 100%; margin-top: 10px; margin-bottom: 5px; clear: both;'>" & _
        "<div style='font-family: ""Aptos (Body)"", Arial, sans-serif; font-size: 10pt; line-height: 1.4; margin-bottom: 15px;'>" & _
        "<p style='margin-bottom: 5px;'><strong>Preston M. Byers</strong><br>Principal & Advisor<br>Reliant Business & Wealth Strategies</p>" & _
        "<p style='margin-bottom: 10px; font-size: 9pt;'><strong>Financial Advisor</strong> offering investment advisory services through <strong>Eagle Strategies LLC</strong>, a Registered Investment Adviser.<br>" & _
        "<strong>Registered Representative</strong> offering securities through <strong>NYLIFE Securities LLC</strong> (member <strong>FINRA/SIPC</strong>, A Licensed Insurance Agency).<br>" & _
        "<strong>Eagle Strategies</strong> and <strong>NYLIFE Securities</strong> are New York Life Companies.<br>" & _
        "<strong>Reliant Business & Wealth Strategies</strong> is not owned or operated by Eagle Strategies LLC or its affiliates.</p>" & _
        "<p style='margin-bottom: 10px; font-size: 9pt;'><strong>Phone (Eastern States):</strong> 601-622-7155<br>" & _
        "<strong>Phone (Western States):</strong> 469-294-3812<br>" & _
        "<strong>Website:</strong> <a href='http://www.reliantwealthstrategies.com' style='color: #0066cc;'>www.reliantwealthstrategies.com</a><br>" & _
        "<strong>General Office:</strong> 1052 Highland Colony Parkway, Suite 101, Ridgeland, MS 39157</p></div>" & _
        "<div style='font-family: ""Aptos (Body)"", Arial, sans-serif; font-size: 8pt; line-height: 1.4; margin-bottom: 2px;'>" & _
        "<span style='font-size: 7pt;'>Opt-Out Notice: If you do not wish to receive email communications from New York Life and/or NYLIFE Securities LLC, please reply to this email using the words ""opt out"" in the subject line.<br>" & _
        "Please copy <a href='mailto:email_optout@newyorklife.com'>email_optout@newyorklife.com</a>.<br>" & _
        "New York Life Insurance Company, 51 Madison Ave., New York, NY 10010</span></div>"
End Function

Private Function ComposeFinancialOverviewBody() As String
    Dim sections As Collection
    Set sections = New Collection

    sections.Add "<p><strong><span style='font-size: 14pt;'>" & BuildGreeting() & "</span></strong></p>"
    sections.Add "<p>Below, you will find an outline of the information I will use to assess your financial picture and provide clarity on the best path forward. This outline and the attached templates are comprehensive; however, if any section is not relevant to your situation, you may skip it and concentrate on the sections that apply.</p>"
    sections.Add "<p><strong>What to expect with our financial planning process, as we move forward:</strong></p><ol>" & _
        "<li><strong>Data Collection & Initial Consultation:</strong> My team and I will compile the required information into reports for our initial consultation. During this meeting, we will review your current position, discuss your goals, and outline the steps to move forward.</li>" & _
        "<li><strong>Development:</strong> If we proceed with planning, we will create and stress-test hypothetical models to evaluate different strategies.</li>" & _
        "<li><strong>Plan Overview Meeting:</strong> We will review the proposed models with you, explain their advantages, and help you choose the best path forward.</li>" & _
        "<li><strong>Implementation:</strong> We will prioritize each action step, implementing the plan in phases. While some actions will be executed immediately, others will take more time. We will monitor your progress and provide guidance throughout each stage.</li>" & _
        "<li><strong>Ongoing Advice & Monitoring:</strong> We will meet periodically to assess progress, offer advice on financial decisions, and help you adapt to any changes or challenges.</li></ol>"
    sections.Add "<p><strong>Outline of Information for your Financial Overview:</strong></p><ul>" & _
        "<li><strong>Personal Expenses and Balance Sheet:</strong> I have attached a spreadsheet for your expenses ('Financial Planning Monthly Budget') and one for your assets and debts ('Networth Template').</li></ul>" & _
        "<p><strong>Taxes:</strong></p><ul>" & _
        "<li><strong>Tax Returns:</strong> Include the most recent tax returns, as well as the returns for the prior year if you can access them easily. Your accountant may have emailed you a PDF copy.</li></ul>"
    sections.Add "<p><strong>Current Assets and Investments (Personal & Business if applicable):</strong></p><ul>" & _
        "<li><strong>Retirement and Investment Accounts:</strong> Provide statements or screenshots for any retirement or investment accounts you currently have. This will help us assess their approximate value and how they are invested.</li>" & _
        "<li><strong>Banking Assets:</strong> Describe how you manage your banking assets and provide approximate values for these accounts (for example, 'I keep $50,000 in savings, and my checking account usually has a balance of around $5,000').</li>" & _
        "<li><strong>IMPORTANT:</strong> If you have been contributing to savings, retirement accounts, or other savings vehicles, please let me know the amount you contribute and how often.</li></ul>" & _
        "<p><strong>Insurance</strong></p><ul>" & _
        "<li><strong>Current Insurance policies:</strong> Include all types you have, including Life Insurance, Disability Insurance, and Long-Term Care Insurance.</li>" & _
        "<li>Please send a scan and/or picture of the pages of the policy that shows the following:</li>" & _
        "<ul><li>Name of company, type of policy, policy number, the benefit amount, cost, and the page that shows how it grows and/or changes over time.</li></ul></ul>"
    sections.Add "<p><strong>Goals & Issues on the Horizon:</strong></p><ul>" & _
        "<li>Share any projects, financial concerns, significant purchases, or investments you are considering. Provide details such as:</li>" & _
        "<ul><li>A description of the project, requirement, or goal, along with an estimated budget. If it is an expense, provide the estimated cost.</li>" & _
        "<li>The price and loan terms for any real estate loans or other potential loans.</li></ul></ul>" & _
        "<p><strong>Any alternate scenarios or ideas you would like to explore:</strong></p><ul>" & _
        "<li>Include any alternate scenarios you would like to explore within your plans, such as investing in different types of investments or retiring at an earlier age, etc.</li>" & _
        "<li>Note any specific plans or preferences for passing your estate to your heirs. This will be essential in developing effective transfer strategies.</li></ul>" & _
        "<p>I will be on the lookout for your email. Please let me know if you have any questions.</p>"

    ComposeFinancialOverviewBody = "<html><body>" & JoinCollection(sections) & BuildSignatureBlock() & "</body></html>"
End Function

Private Function ComposeInformationForPlanningBody() As String
    Dim sections As Collection
    Set sections = New Collection

    sections.Add "<p><strong><span style='font-size: 14pt;'>" & BuildGreeting() & "</span></strong></p>"
    sections.Add "<p>Below is an outline of the information I will use to assess your business and financial picture and begin planning with you. This outline is comprehensive. However, if any section is not relevant to your situation don't stress; just skip over it and concentrate on the sections that apply.</p>"
    sections.Add "<p><strong>What to expect:</strong></p><ol>" & _
        "<li><strong>Data Collection & Initial Consultation:</strong> My team and I will gather the necessary information to prepare comprehensive reports for our initial consultation. In this session, we will assess your current situation, explore your goals, and establish a roadmap for achieving them.</li>" & _
        "<li><strong>Development:</strong> If we proceed with planning, we will create and stress-test hypothetical models to evaluate different strategies.</li>" & _
        "<li><strong>Plan Overview Meeting:</strong> We will walk you through the proposed models, highlighting their benefits to help you determine the optimal path forward.</li>" & _
        "<li><strong>Implementation:</strong> We will prioritize each action step, implementing the plan in phases. While some actions will be executed immediately, others will take more time. We will monitor your progress and provide guidance throughout each stage.</li>" & _
        "<li><strong>Ongoing Advice & Monitoring:</strong> We will meet periodically to assess progress, offer advice on financial decisions, and help you adapt to any changes or challenges.</li></ol>"
    sections.Add "<p><strong>INFORMATION TO SEND US:</strong></p>" & _
        "<p><strong>Personal Expenses and Balance Sheet</strong></p><ul>" & _
        "<li>I have attached a spreadsheet for your expenses (""Financial Planning Monthly Budget"") and one for your assets and debts (""Networth Template"").</li></ul>" & _
        "<p><strong>Business & Tax</strong></p><ul>" & _
        "<li><strong>Profit & Loss Statements:</strong> Please provide 3-5 years of Profit & Loss Statements for your business(es). *We like to see the P&L by month for each year so we can analyze trends. If any business has less than 3yrs of data, just send what you have so far. If you use QuickBooks, you'll be able to access these from the reports section and download them as a PDF.</li>" & _
        "<li><strong>Tax Returns:</strong> Include the most recent tax returns for both you and your business, as well as the returns for the prior year if you can access them easily. Your accountant may have emailed you a PDF copy.</li>" & _
        "<li><strong>Balance Sheet:</strong> Please provide a balance sheet for your business that details its assets and liabilities.</li></ul>" & _
        "<p><strong>Current Assets and Investments (Personal & Business) – If applicable</strong></p><ul>" & _
        "<li><strong>Retirement and Investment Accounts:</strong> Provide statements or screenshots for any retirement or investment accounts you currently have. This will help us assess their approximate value and how they are invested.</li>" & _
        "<li><strong>Banking Assets:</strong> Describe how you manage your banking assets and provide approximate values for these accounts (for example, 'I keep $50,000 in savings, and my checking account usually has a balance of around $5,000').</li>" & _
        "<li><strong>Business Accounts:</strong> Please provide similar information for your business accounts, especially if the business maintains operating, tax, and reserve accounts.</li></ul>" & _
        "<p><strong>IMPORTANT:</strong> If you have been contributing to savings, retirement accounts, or other savings vehicles, please let me know the amount you contribute and how often.</p>"
    sections.Add "<p><strong>Insurance</strong></p><ul>" & _
        "<li><strong>Current Insurance policies:</strong> Include all types you have, such as, Life Insurance, Disability Insurance, and Long-Term Care Insurance. (DREW… I obviously have some of this already, so just let me know about anything I don't already know about)</li>" & _
        "<li>Please send a scan and/or picture of the pages of the policy that shows the following:</li>" & _
        "<ul><li>Name of company, type of policy, policy number, the benefit amount, cost, and the page that show how it grows and/or changes over time.</li></ul></ul>"
    sections.Add "<p><strong>Goals & Issues on the Horizon</strong></p><ul>" & _
        "<li><strong>Future Projects and Investments:</strong> Share any projects, major purchases, or investments you are considering. Provide details such as:</li>" & _
        "<ul><li>P&L information on a business you wish to purchase and the purchase methods you are contemplating.</li>" & _
        "<li>The price and loan terms for any real estate property.</li>" & _
        "<li>Details regarding a construction project or new location or branch you plan to open.</li>" & _
        "<li>Information about a new business or product/service line you might add in the future.</li></ul></ul>"
    sections.Add "<p><strong>Any alternate scenarios or ideas you would like to explore:</strong></p><ul>" & _
        "<li>Include any alternate scenarios you would like to explore within your plans, such as selling or purchasing a business via one method or another, buying a rental house outright versus using a loan, retiring at an earlier age, etc.</li>" & _
        "<li>Note any specific plans or preferences for passing your estate to your heirs. This will be essential in developing effective transfer strategies.</li></ul>" & _
        "<p>I will be on the lookout for your email. Please let me know if you have any questions.</p>"

    ComposeInformationForPlanningBody = "<html><body>" & JoinCollection(sections) & BuildSignatureBlock() & "</body></html>"
End Function

Private Sub AddAttachments(ByVal mail As Outlook.MailItem, ByVal attachmentNames As Variant)
    Dim index As Long
    Dim filePath As String

    For index = LBound(attachmentNames) To UBound(attachmentNames)
        filePath = FinancialEmailConfig.ATTACHMENT_ROOT & attachmentNames(index)

        If Len(Dir$(filePath)) > 0 Then
            mail.Attachments.Add filePath
        Else
            Debug.Print "Attachment missing: " & filePath
        End If
    Next index
End Sub

Private Function StandardAttachments() As Variant
    ' Always include the three planning documents referenced in both templates:
    '   - Networth_Template.xlsx
    '   - Financial_Planning_Monthly_Budget.xlsx
    '   - ES.DBA.FINANCIALPLANNING.pdf
    StandardAttachments = Array("Networth_Template.xlsx", _
                                "Financial_Planning_Monthly_Budget.xlsx", _
                                "ES.DBA.FINANCIALPLANNING.pdf")
End Function

Private Function JoinCollection(ByVal sections As Collection) As String
    Dim buffer() As String
    Dim index As Long

    ReDim buffer(1 To sections.Count)
    For index = 1 To sections.Count
        buffer(index) = sections(index)
    Next index

    JoinCollection = Join(buffer, vbNullString)
End Function

Private Function InitialiseMail(ByVal recipient As String, _
                                ByVal subjectLine As String, _
                                ByVal htmlBody As String, _
                                ByVal attachments As Variant) As Outlook.MailItem

    Dim mail As Outlook.MailItem
    Set mail = Application.CreateItem(olMailItem)

    With mail
        If Len(Trim$(recipient)) > 0 Then
            .To = recipient
        End If
        .CC = FinancialEmailConfig.CC_RECIPIENT
        .BCC = FinancialEmailConfig.TRACKING_BCC
        .Subject = subjectLine
        .BodyFormat = olFormatHTML
        .HTMLBody = htmlBody
        AddAttachments mail, attachments
    End With

    Set InitialiseMail = mail
End Function

Public Sub DraftFinancialOverviewEmail(Optional ByVal recipient As String)
    With InitialiseMail(recipient, _
                        "Financial Overview", _
                        ComposeFinancialOverviewBody(), _
                        StandardAttachments())
        .Display
    End With
End Sub

Public Sub DraftInformationForPlanningEmail(Optional ByVal recipient As String)
    With InitialiseMail(recipient, _
                        "Information for Planning", _
                        ComposeInformationForPlanningBody(), _
                        StandardAttachments())
        .Display
    End With
End Sub

