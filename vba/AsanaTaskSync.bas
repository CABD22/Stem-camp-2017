Attribute VB_Name = "AsanaTaskSync"
Option Explicit

' Helpers that push qualifying Outlook messages to Asana once they are sent.
' In practice you will BCC your Cloudflare Worker so the Outlook macro can
' send it the recipient + subject immediately after you hit Send; the worker
' then creates the Asana task (and fills the custom fields) with your token,
' keeping credentials out of VBA.  A direct Asana API option is included for
' completeness, but the worker-first pattern is the safest and easiest to
' reuse across computers. Configure the constants below to match the method
' you choose.

' Direct API calls map the primary recipient into the "Email Address"
' custom field and the Outlook subject into the "Subject" custom field.
'
' All configurable values (worker endpoint, PAT, project/custom-field IDs,
' and task name) are stored in FinancialEmailConfig.bas so you only edit them
' once.

Public Function ShouldSyncMailToAsana(ByVal mail As Outlook.MailItem) As Boolean
    Dim subjectText As String
    subjectText = LCase$(Trim$(mail.Subject))

    Select Case subjectText
        Case "financial overview", "information for planning"
            ShouldSyncMailToAsana = (InStr(1, LCase$(mail.BCC & vbNullString), _
                                           LCase$(FinancialEmailConfig.TRACKING_BCC), vbBinaryCompare) > 0)
    End Select
End Function

Public Sub SyncMailToAsana(ByVal mail As Outlook.MailItem)
    Dim primaryRecipient As String
    primaryRecipient = ExtractPrimaryRecipient(mail)

    If Len(primaryRecipient) = 0 Then
        Debug.Print "[AsanaSync] No primary recipient detected; skipping Asana sync."
        Exit Sub
    End If

    If Len(Trim$(FinancialEmailConfig.WORKER_ENDPOINT)) > 0 Then
        PostViaWorker primaryRecipient, mail.Subject
    Else
        PostDirectlyToAsana primaryRecipient, mail.Subject
    End If
End Sub

Private Sub PostViaWorker(ByVal recipient As String, _
                          ByVal subjectLine As String)
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    On Error GoTo WorkerError

    ' Same idea as the quick XMLHTTP snippet shared earlier: push the
    ' essentials to your Cloudflare Worker so it can finish talking to Asana
    ' (and populate the custom fields server-side where credentials live
    ' safely). Keep WORKER_ENDPOINT empty if you prefer the direct API route.
    http.Open "POST", FinancialEmailConfig.WORKER_ENDPOINT, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send BuildWorkerPayload(recipient, subjectLine)

    If http.Status < 200 Or http.Status >= 300 Then
        Debug.Print "[AsanaSync] Worker error (" & http.Status & "): " & http.responseText
    Else
        Debug.Print "[AsanaSync] Worker acknowledged task for " & recipient
    End If

    Exit Sub

WorkerError:
    Debug.Print "[AsanaSync] Error calling worker: " & Err.Description
End Sub

Private Sub PostDirectlyToAsana(ByVal recipient As String, _
                                 ByVal subjectLine As String)
    ' Optional fallback for environments where you would rather post straight
    ' to Asana. You must store the PAT + project/custom-field IDs locally,
    ' which is why the worker-first flow is generally preferred.
    If Len(Trim$(FinancialEmailConfig.ASANA_PAT)) = 0 _
            Or Len(Trim$(FinancialEmailConfig.ASANA_PROJECT_GID)) = 0 _
            Or Len(Trim$(FinancialEmailConfig.CF_EMAIL_ADDRESS_GID)) = 0 _
            Or Len(Trim$(FinancialEmailConfig.CF_SUBJECT_GID)) = 0 Then
        Debug.Print "[AsanaSync] Direct Asana credentials incomplete; skipping."
        Exit Sub
    End If

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo RequestError

    http.Open "POST", "https://app.asana.com/api/1.0/tasks", False
    http.setRequestHeader "Authorization", "Bearer " & FinancialEmailConfig.ASANA_PAT
    http.setRequestHeader "Content-Type", "application/json"
    http.Send BuildDirectPayload(recipient, subjectLine)

    If http.Status < 200 Or http.Status >= 300 Then
        Debug.Print "[AsanaSync] Asana API error (" & http.Status & "): " & http.responseText
    Else
        Debug.Print "[AsanaSync] Task created successfully for " & recipient
    End If

    Exit Sub

RequestError:
    Debug.Print "[AsanaSync] Error posting to Asana: " & Err.Description
End Sub

Private Function BuildWorkerPayload(ByVal recipient As String, _
                                    ByVal subjectLine As String) As String
    ' Minimal JSON the worker needs so it can call the Asana API with your PAT
    ' and write the "Email Address" + "Subject" custom fields at task creation
    ' time. Posting straight from Outlook would expose the PAT, so the worker
    ' stays in the middle as the secure bridge.
    BuildWorkerPayload = _
        "{""email"":""" & JsonEscape(recipient) & """," & _
        ""subject"":""" & JsonEscape(subjectLine) & """," & _
        ""date_sent"":""" & JsonEscape(Format$(Now, "yyyy-mm-dd\THH:NN:SS\Z")) & """}"
End Function

Private Function BuildDirectPayload(ByVal recipient As String, _
                                    ByVal subjectLine As String) As String
    Dim q As String
    q = Chr$(34)

    BuildDirectPayload = _
        "{" & q & "data" & q & ":{" & _
        q & "name" & q & ":" & q & JsonEscape(FinancialEmailConfig.TASK_NAME) & q & "," & _
        q & "projects" & q & ":[" & q & JsonEscape(FinancialEmailConfig.ASANA_PROJECT_GID) & q & "]," & _
        q & "custom_fields" & q & ":{" & _
        q & JsonEscape(FinancialEmailConfig.CF_EMAIL_ADDRESS_GID) & q & ":" & q & JsonEscape(recipient) & q & "," & _
        q & JsonEscape(FinancialEmailConfig.CF_SUBJECT_GID) & q & ":" & q & JsonEscape(subjectLine) & q & _
        "}}}"
End Function

Private Function ExtractPrimaryRecipient(ByVal mail As Outlook.MailItem) As String
    Dim recipient As Outlook.Recipient
    Dim addressValue As String

    If mail.Recipients.Count = 0 Then Exit Function

    Call mail.Recipients.ResolveAll

    Set recipient = mail.Recipients(1)
    addressValue = GetSmtpAddress(recipient)

    If Len(addressValue) = 0 Then
        addressValue = Trim$(recipient.Address)
    End If

    If Len(addressValue) = 0 Then
        addressValue = ExtractAddressFromDisplay(mail.To)
    End If

    ExtractPrimaryRecipient = addressValue
End Function

Private Function GetSmtpAddress(ByVal recipient As Outlook.Recipient) As String
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    On Error Resume Next

    Dim accessor As Outlook.PropertyAccessor
    Set accessor = recipient.PropertyAccessor

    GetSmtpAddress = Trim$(accessor.GetProperty(PR_SMTP_ADDRESS))

    On Error GoTo 0
End Function

Private Function ExtractAddressFromDisplay(ByVal addressList As String) As String
    Dim parts As Variant
    Dim index As Long
    Dim candidate As String

    parts = Split(addressList, ";")

    For index = LBound(parts) To UBound(parts)
        candidate = StripAddress(parts(index))
        If Len(candidate) > 0 Then
            ExtractAddressFromDisplay = candidate
            Exit Function
        End If
    Next index
End Function

Private Function StripAddress(ByVal value As String) As String
    Dim cleaned As String
    Dim startPos As Long
    Dim endPos As Long

    cleaned = Trim$(value)
    cleaned = Replace(cleaned, Chr$(34), vbNullString)

    startPos = InStr(cleaned, "<")
    endPos = InStr(cleaned, ">")

    If startPos > 0 And endPos > startPos Then
        cleaned = Mid$(cleaned, startPos + 1, endPos - startPos - 1)
    End If

    StripAddress = Trim$(cleaned)
End Function

Private Function JsonEscape(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\"")
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")

    JsonEscape = value
End Function

Public Sub RunAsanaSyncFromRule(ByVal mail As Outlook.MailItem)
    ' Entry point for an "after sending" Outlook rule. Create a rule that
    ' matches your fast-button subjects, choose "run a script", and pick this
    ' procedure if you would rather avoid editing ThisOutlookSession. The rule
    ' is triggered immediately after Outlook hands the message to Exchange, so
    ' your worker receives the email + subject while the details are still
    ' fresh—no need to wait for an Asana webhook to tidy fields later.
    If ShouldSyncMailToAsana(mail) Then
        SyncMailToAsana mail
    End If
End Sub
