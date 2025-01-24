VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTktInput 
   Caption         =   "Enter Ticket Details"
   ClientHeight    =   9410.001
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   9060.001
   OleObjectBlob   =   "frmTktInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTktInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ValidateBlankControls() As String
    Dim blankControls As String

    If txtIssueDate.Value = "" Then
        blankControls = "Issue date"
    End If

    If txt_IssueDescription.Value = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Description"
    End If

    If cmb_Frequency.Text = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Frequency"
    End If

    If cmb_SeverityPriority.Text = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Severity"
    End If

    If cmb_ComponentsEffected.Text = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Components Affected"
    End If

    If cmb_EnvEffected.Text = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Environment"
    End If

    If cmb_TransactName.Text = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Transaction Name"
    End If

    If txtDateEffected.Value = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Date Affected"
    End If

    If txtRecentChanges.Value = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Recent Changes"
    End If

    If txtWrkArndAvlbl.Visible = True And txtWrkArndAvlbl.Value = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Workaround Available"
    End If

    If txtAblToReproduce.Visible = True And txtAblToReproduce.Value = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Reproducible"
    End If

    If txtStepsToReproduce.Value = "" Then
        If blankControls <> "" Then blankControls = blankControls & ", "
        blankControls = blankControls & "Steps to Reproduce"
    End If

    ValidateBlankControls = blankControls
End Function

Private Sub cmd_Close_Click()
    Unload frmTktInput
End Sub

Private Sub cmd_Save_Click()
    'Call the validation function
    Dim blankColumns As String
    blankColumns = ValidateBlankControls()

    'If any controls are blank, show a message box with the names of blank controls
    If blankColumns <> "" Then
        MsgBox "Please fill the following blank fields: " & vbCrLf & blankColumns
    Else
        'Set the worksheet
        Set ws = ThisWorkbook.Sheets("Problems")
        activeRow = ActiveCell.row
        
        
        
        'Save the data
        ws.Cells(activeRow, 8).Value = txt_IssueDescription.Text
        ws.Cells(activeRow, 9).Value = txtIssueDate.Text
        ws.Cells(activeRow, 10).Value = cmb_Frequency.SelText
        ws.Cells(activeRow, 11).Value = cmb_SeverityPriority.SelText
        ws.Cells(activeRow, 12).Value = cmb_ComponentsEffected.SelText
        ws.Cells(activeRow, 15).Value = cmb_EnvEffected.Text
        ws.Cells(activeRow, 16).Value = cmb_TransactName.Text
        ws.Cells(activeRow, 17).Value = txtDateEffected.Value
        ws.Cells(activeRow, 18).Value = txtRecentChanges.Text

        If txtWrkArndAvlbl.Visible = True Then
            ws.Cells(activeRow, 19).Value = txtWrkArndAvlbl.Value
        Else
            ws.Cells(activeRow, 19).Value = "No"
        End If

        If ob_AbleToReproduce_No.Value = True Then
            ws.Cells(activeRow, 20).Value = txtAblToReproduce.Text
        Else
            ws.Cells(activeRow, 20).Value = "Yes"
        End If

        ws.Cells(activeRow, 21).Value = txtStepsToReproduce.Text
        
        ws.Cells(activeRow, 22).Value = cmb_AsgndTo.Value
        
    
        'Close the Form after saving
        'Unload frmTktInput
    End If
End Sub

Private Sub cmd_SaveAndSendEmail_Click()
     
    Dim blankControls As String
    Dim fromaddrvar As String
    
    ' Validate form controls
    blankControls = ValidateBlankControls()
    
    If blankControls <> "" Then
        MsgBox "Please fill these fields: " & blankControls, vbExclamation, "Missing Fields"
    Else
        ' Save the data into the worksheet
        cmd_Save_Click
        
        emailConfiguration.InitializeSMTPConfig
        
        'fromaddrvar = setFromEmailAddress(cmb_AsgndTo.Value)
        'Load the Email form
        'frmEmailSummary.Show
        
        'my_email
        openEmail
        Unload frmTktInput
        
    End If
End Sub

Function setFromEmailAddress(ByVal sEmailType) As String

    
Dim sFromEmailAddress As String

sFromEmailAddress = "***"


If sEmailType = "CSO" Then
    sFromEmailAddress = emailConfiguration.CSOEmailAddress
    
ElseIf sEmailType = "MSO" Then
    sFromEmailAddress = emailConfiguration.MSOEmailAddress
End If



setFromEmailAddress = sFromEmailAddress
    
End Function

Function setSubjectValue() As String

    
    Dim subjectTxt As String

    If Not IsNull(txt_Subject.Value) And Trim(txt_Subject.Value) <> "" Then
            subjectTxt = " - " & txt_Subject.Value
    Else
            subjectTxt = " "
    End If

setSubjectValue = subjectTxt
    
End Function
Function setToaddressEmail() As String

    Dim MasterDataSheet As Worksheet
    Dim toAddressEmail As String
    
    Set MasterDataSheet = ThisWorkbook.Sheets("MasterData")
    
    If cmb_AsgndTo.Value = "CSO" Then
      toAddressEmail = MasterDataSheet.Range("A2").Value
    'for MSO emails
    Else
      toAddressEmail = MasterDataSheet.Range("D2:D15")
    End If
      
setToaddressEmail = toAddressEmail

End Function

Function setCCaddressEmail() As String
    Dim MasterDataSheet As Worksheet
    Dim CCAddressEmail As String
    Dim cell As Range
    
    Set MasterDataSheet = ThisWorkbook.Sheets("MasterData")
    
    If cmb_AsgndTo.Value = "CSO" Then
        CCAddressEmail = ""
        
       'Loop through B2 to B4
        For Each cell In MasterDataSheet.Range("B2:B15")
            If cell.Value <> "" Then
                If CCAddressEmail = "" Then
                    CCAddressEmail = cell.Value
                Else
                    CCAddressEmail = CCAddressEmail & ";" & cell.Value
                End If
            End If
        Next cell
    'For MSO Emails
    Else
        CCAddressEmail = MasterDataSheet.Range("C2:C15")
    End If
    
    setCCaddressEmail = CCAddressEmail
End Function


Sub my_email()


Set xOutApp = CreateObject("Outlook.Application")
Set xOutMail = xOutApp.CreateItem(0)
    
'    Set xOutApp = CreateObject("Outlook.Application")
'    Set xOutMail = xOutApp.CreateItemFromTemplate("F:\Special\LMS Support\Templates & Shortcuts\LMS - MA CSO Submission - Template.oft")

    
signature = xOutMail.body

With xOutMail
.To = setFromEmailAddress(cmb_AsgndTo.Value)  ' Application.UserName
.CC = ""
'.BCC = ""
 .subject = "Problem Tracker - open in edit mode for now 10 minutes !"
.body = "Hello - Just a friendly reminder that you have had the Problem Tracker open in edit mode for 10 minutes now." & vbNewLine & vbNewLine _
& "Please close it if you are done with it."
.Send

End With

On Error GoTo 0
Set xOutMail = Nothing
Set xOutApp = Nothing


End Sub


Private Sub lbl_Subject_Click()

End Sub

Private Sub ob_AbleToReproduce_No_Click()
    ob_AbleToReproduce_No.Value = True
    txtAblToReproduce.Visible = True
End Sub

Private Sub ob_AbleToReproduce_Yes_Click()
    
    txtAblToReproduce.Visible = False
    
End Sub

Private Sub ob_WrkArndAvlbl_Yes_Click()
'show the Text Control for the WorkAroundText Control

txtWrkArndAvlbl.Visible = True
End Sub

Private Sub op_WrkArndAvlbl_No_Click()
'Hide the Text Control for the WorkAroundText Control
txtWrkArndAvlbl.Visible = False
End Sub
Function ValidateAndFormatDate(ByVal dateInput As String) As String
    Dim month As Integer
    Dim day As Integer
    Dim year As Integer
    Dim FormattedDate As String

    ' Check if the input length is exactly 8 characters
    If Len(dateInput) <> 8 Or dateInput = "*[A-za-z]*" Then
        ValidateAndFormatDate = "Invalid date format"
        Exit Function
    End If

    ' Extract components
    month = CInt(Mid(dateInput, 1, 2))
    day = CInt(Mid(dateInput, 3, 2))
    year = CInt(Mid(dateInput, 5, 4))

    ' Validate month
    If month < 1 Or month > 12 Then
        ValidateAndFormatDate = "Invalid month"
        Exit Function
    End If

    ' Validate day
    If day < 1 Or day > 31 Then
        ValidateAndFormatDate = "Invalid day"
        Exit Function
    End If

    ' Validate year
    If year < 1000 Or year > 9999 Then
        ValidateAndFormatDate = "Invalid year"
        Exit Function
    End If

    ' If all validations pass, format the date
    FormattedDate = Format$(DateSerial(year, month, day), "MMM DD YYYY")

    ValidateAndFormatDate = FormattedDate
End Function

Private Sub txtAblToReproduce_Change()

End Sub

Private Sub txtDateEffected_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
     Dim FormattedDate As String
     
     If KeyCode = vbKeyTab Then
        
        ' Prevent the default Tab key action
        KeyCode = 0
        
        FormattedDate = ValidateAndFormatDate(txtDateEffected.Value)
        If FormattedDate = "Invalid date format" Or FormattedDate = "Invalid month" Or FormattedDate = "Invalid day" Or FormattedDate = "Invalid year" Then
            lbl_InstructionDtEffected.Caption = FormattedDate & " - Enter the date in MMDDYYYY format"
            lbl_InstructionDtEffected.ForeColor = RGB(255, 0, 0)
            txtDateEffected.Value = ""
            txtDateEffected.SetFocus
            Exit Sub
        Else
            txtDateEffected.Value = FormattedDate
            lbl_InstructionDtEffected.Caption = ""
            txtRecentChanges.SetFocus
            Exit Sub
            
55
        End If
     End If
End Sub



Private Sub txtRecentChanges_Change()

End Sub

Private Sub txtStepsToReproduce_Change()

End Sub

Private Sub UserForm_Initialize()

    'Load the Frequency DropDown
    loadItems_Frequency
    'Load the Severity Priotity List Dropdown
    loadItems_SeverityPriority
    'load the components effected
    loadItems_ComponentsEffected
    'load the environnments list
    loadItems_EnvironmentsEffected
    'load the transaction names
    loadItems_TransNames
    loadItems_AsgndTo
    'by default hide the text control
    txtWrkArndAvlbl.Visible = False
    
End Sub


Private Sub loadItems_AsgndTo()
    cmb_AsgndTo.AddItem ("CSO")
    cmb_AsgndTo.AddItem ("MSO")
End Sub

Private Sub loadItems_Frequency()
    cmb_Frequency.AddItem ("Some Time")
    cmb_Frequency.AddItem ("Every Time")
    cmb_Frequency.AddItem ("Op1")
    cmb_Frequency.AddItem ("Op2")
End Sub

Private Sub loadItems_SeverityPriority()
    cmb_SeverityPriority.AddItem ("Low")
    cmb_SeverityPriority.AddItem ("High")
    cmb_SeverityPriority.AddItem ("Medium")
    cmb_SeverityPriority.AddItem ("Critical")
End Sub

Private Sub loadItems_ComponentsEffected()
    cmb_ComponentsEffected.AddItem ("User Interface - UI")
    cmb_ComponentsEffected.AddItem ("RF")
    cmb_ComponentsEffected.AddItem ("Database - DB")
    cmb_ComponentsEffected.AddItem ("Other")
End Sub

Private Sub loadItems_EnvironmentsEffected()
    cmb_EnvEffected.AddItem ("UAT  - User Acceptance")
    cmb_EnvEffected.AddItem ("PROD - Production")
    cmb_EnvEffected.AddItem ("DEV  - Development")
    cmb_EnvEffected.AddItem ("SBX  - SandBox")
    cmb_EnvEffected.AddItem ("OTH  - Other")
End Sub

Private Sub loadItems_TransNames()
    cmb_TransactName.AddItem ("op-1")
    cmb_TransactName.AddItem ("op-2")
    cmb_TransactName.AddItem ("op-3")

End Sub


Private Sub txtIssueDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim FormattedDate As String
     
     If KeyCode = vbKeyTab Then
     
        ' Prevent the default Tab key action
        KeyCode = 0
        
        FormattedDate = ValidateAndFormatDate(txtIssueDate.Value)
        If FormattedDate = "Invalid date format" Or FormattedDate = "Invalid month" Or FormattedDate = "Invalid day" Or FormattedDate = "Invalid year" Then
            lbl_InstructionIssueStDate.Caption = FormattedDate & " - Enter the date in MMDDYYYY format"
            lbl_InstructionIssueStDate.ForeColor = RGB(255, 0, 0)
            txtIssueDate.Value = ""
            txtIssueDate.SetFocus
            Exit Sub
        Else
            txtIssueDate.Value = FormattedDate
            lbl_InstructionIssueStDate.Caption = ""
            cmb_Frequency.SetFocus
           
55
        End If
     End If
End Sub

Sub openEmail()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Problems")

    Dim emailSubject As String, EmailSignature As String
    Dim Email_To As String, Email_CC As String, Email_BCC As String
    Dim emailBody As String
    
    Dim DisplayEmail As Boolean
    Dim OutlookApp As Object, OutlookMail As Object


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
    'txtFromEmail.Value = SMTPFromEmail ' Default "From" email
    'txtToEmail.Value = setFromEmailAddress(cmb_AsgndTo.Value) 'SMTPToEmail  ' Default "To" email

    ' Populate the subject field
    'emailSubject = "[ PRB - " & ws.Cells(activeRow, 6).Value & " ] - [ " & Replace(ws.Cells(activeRow, 18), vbCrLf, " ") & " ] - PROD Code"
    emailSubject = "[ PRB -  " & ws.Cells(activeRow, 6).Value & "  ] - HOME " & setSubjectValue()

    ' Populate the email body
    emailBody = "Description: " & CStr(ws.Cells(activeRow, 8).Value) & vbCrLf & _
                "Issue Date: " & CStr(ws.Cells(activeRow, 9).Value) & vbCrLf & _
                "Frequency: " & CStr(ws.Cells(activeRow, 10).Value) & vbCrLf & _
                "Severity: " & CStr(ws.Cells(activeRow, 11).Value) & vbCrLf & _
                "Components Affected: " & CStr(ws.Cells(activeRow, 12).Value) & vbCrLf & _
                "Environment: " & CStr(ws.Cells(activeRow, 15).Value) & vbCrLf & _
                "Transaction Name: " & CStr(ws.Cells(activeRow, 16).Value) & vbCrLf & _
                "Date Affected: " & CStr(ws.Cells(activeRow, 17).Value) & vbCrLf & _
                "Recent Changes: " & ws.Cells(activeRow, 18).Value & vbCrLf & _
                "Workaround Available: " & CStr(ws.Cells(activeRow, 19).Value) & vbCrLf & vbCrLf & _
                "Able to Reproduce: " & CStr(ws.Cells(activeRow, 20).Value) & vbCrLf & vbCrLf & _
                "Steps to Reproduce: " & vbCrLf & CStr(ws.Cells(activeRow, 21).Value) & vbCrLf


    'Create an Outlook object and new mail message
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    'Display email and specify To, Subject, etc
    With OutlookMail

        .Display
        ' .From = "****"
        ' setFromEmailAddress(cmb_AsgndTo.Value) 'SMTPToEmail  ' Default "To" email
        ' SMTPFromEmail
        ' .To = "****"
        .To = setToaddressEmail()
        .CC = setCCaddressEmail()
    '    .CC = Email_CC
    '    .BCC = Email_BCC
        .subject = emailSubject
        .body = "Hi Team," & vbNewLine & vbNewLine _
        & "Please find below the details of the PRB " & vbNewLine & vbNewLine _
        & emailBody

' set to true to view email before sending

        If DisplayEmail = True Then

            .Send

        End If

    End With



End Sub


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
    Email_To = "****"
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


