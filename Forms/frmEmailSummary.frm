VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmailSummary 
   Caption         =   "Email Summary"
   ClientHeight    =   6700
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7310
   OleObjectBlob   =   "frmEmailSummary.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmailSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    
    InitializeSMTPConfig
    ' Reference the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Problems")

    ' Ensure there is an active cell
    Dim activeRow As Integer
    On Error Resume Next
    activeRow = ActiveCell.row
    If Err.Number <> 0 Or activeRow = 0 Then
        MsgBox "Error: No active cell found. Please select a row before opening the form.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Populate the email fields
    txtFromEmail.Value = SMTPFromEmail ' Default "From" email
    txtToEmail.Value = SMTPToEmail  ' Default "To" email

    ' Populate the subject field
    txtSubject.Value = "[ PRB - " & ws.Cells(activeRow, 6).Value & " ] - [ " & Replace(ws.Cells(activeRow, 18), vbCrLf, " ") & " ] - PROD Code"

    ' Populate the email body
    txtBody.Value = "Description: " & CStr(ws.Cells(activeRow, 8).Value) & vbCrLf & _
                "Issue Date: " & CStr(ws.Cells(activeRow, 9).Value) & vbCrLf & _
                "Frequency: " & CStr(ws.Cells(activeRow, 10).Value) & vbCrLf & _
                "Severity: " & CStr(ws.Cells(activeRow, 11).Value) & vbCrLf & _
                "Components Affected: " & CStr(ws.Cells(activeRow, 12).Value) & vbCrLf & _
                "Environment: " & CStr(ws.Cells(activeRow, 15).Value) & vbCrLf & _
                "Transaction Name: " & CStr(ws.Cells(activeRow, 16).Value) & vbCrLf & _
                "Date Affected: " & CStr(ws.Cells(activeRow, 17).Value) & vbCrLf & _
                "Recent Changes: " & vbCrLf & ws.Cells(activeRow, 18).Value & vbCrLf & _
                "Workaround Available: " & CStr(ws.Cells(activeRow, 19).Value) & vbCrLf & _
                "Able to Reproduce: " & CStr(ws.Cells(activeRow, 20).Value) & vbCrLf & _
                "Steps to Reproduce: " & CStr(ws.Cells(activeRow, 21).Value) & vbCrLf



End Sub



Private Sub cmd_EmailCancel_Click()
    Unload frmEmailSummary
End Sub

Private Sub cmd_EmailSave_Click()
     
    ' Collect values from the form
    Dim fromEmail As String, toEmail As String, subject As String, body As String
    fromEmail = txtFromEmail.Value
    toEmail = txtToEmail.Value
    subject = txtSubject.Value
    body = txtBody.Value
    
    ' Initialize SMTP configuration
    InitializeSMTPConfig

    ' Call the email-sending procedure
    Dim objMessage As Object
    Set objMessage = CreateObject("CDO.Message")
    
    ' SMTP configuration (replace with dynamic settings if needed)
    With objMessage.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPServerPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = SMTPAuthenticate
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPUserName
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPUseSSL
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPConnectionTimeout
        .Update
    End With
    
    ' Email properties
    With objMessage
        .From = fromEmail
        .To = toEmail
        .subject = subject
        .TextBody = body
        .Send
    End With
    
    MsgBox "Email sent successfully!", vbInformation
    
    ' Close the form
    Unload Me


End Sub
