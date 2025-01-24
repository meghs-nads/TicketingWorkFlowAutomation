Attribute VB_Name = "Module9"
Public rowNum As Long

Sub ChecktoSend()

a = 4

Dim ws As Worksheet
Set ws = Worksheets("Problems")

b = Application.WorksheetFunction.CountA(ws.Range("J4:J3000")) + 3

For test = a To b

If Cells(test, 58).Value <> "Email Sent" And Cells(test, 35).Value <> vbNullString And Cells(test, 19).Value = "PROD" And Cells(test, 6).Value <> "1128" Then

Dim emailSubject As String, EmailSignature As String
Dim Email_To As String, Email_CC As String, Email_BCC As String
Dim DisplayEmail As Boolean
Dim OutlookApp As Object, OutlookMail As Object

    emailSubject = "[ PRB - " & Cells(test, 6).Value & " ] - [ " & Cells(test, 18).Value & " ] - PROD Code / Config Change Alert !"
    DisplayEmail = False
    Email_To = "*****"
'    Email_To = "****"
   
    Email_CC = ""
    Email_BCC = ""

    
    'Create an Outlook object and new mail message
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    'Display email and specify To, Subject, etc
    With OutlookMail

        .Display
        .To = Email_To
    '    .CC = Email_CC
    '    .BCC = Email_BCC
        .subject = emailSubject
        .body = "Hi Team," & vbNewLine & vbNewLine _
        & "This PRB has production code / configuration changes as below." & vbNewLine & vbNewLine _
        & "[ PRB ID ] = " & Cells(test, 6).Value & vbNewLine _
        & "[ Prod Go Live (mm/dd/yyyy) ] = " & Cells(test, 31).Value & vbNewLine _
        & "[ HH Owner ] = " & Cells(test, 8).Value & vbNewLine _
        & "[ Primary System ] = " & Cells(test, 18).Value & vbNewLine _
        & "[ Primary Business Area ] = " & Cells(test, 20).Value & vbNewLine _
        & vbNewLine _
        & "[ Code / Config Change Details ] = " & Cells(test, 35).Value & vbNewLine _
        & vbNewLine & vbNewLine _
        & "=======================================" & vbNewLine _
        & "Details of the problem..." & vbNewLine _
        & Cells(test, 10).Value & vbNewLine _
        & vbNewLine & signature


' set to true to view email before sending

        If DisplayEmail = True Then

            .Send

        End If

    End With

'MsgBox "test"
Cells(test, 58).Value = "Email Sent"
End If

Next test


End Sub

