Attribute VB_Name = "emailConfiguration"
' Module: SMTPConfig
Option Explicit

' Configuration details
Public SMTPServer As String
Public SMTPFromEmail As String
Public SMTPServerPort As Long
Public SMTPAuthenticate As Long
Public SMTPUserName As String
Public SMTPPassword As String
Public SMTPToEmail As String
Public SMTPUseSSL As Boolean
Public SMTPConnectionTimeout As Long
Public CSOEmailAddress As String
Public MSOEmailAddress As String


' Initialize configuration
Public Sub InitializeSMTPConfig()
    
    '***********************************************************************************
    'SMTPServer , SMTPServerPort, SMTPFromEmail , SMTPUserName , SMTPPassword , SMTPToEmail to be changed.
    '***********************************************************************************
    
    SMTPServer = "smtp.gmail.com" ' To be replaced with the SMTP server address provided by the client. For example: smtp.office365.com, smtp.yahoo.com, etc.
    SMTPServerPort = 465          ' Port used for secure connections with SMTP servers.
    SMTPAuthenticate = 1          ' Set to 1 for using username and password for authentication.
    SMTPUseSSL = True             ' Set to True to enable SSL encryption for SMTP communication.
    SMTPConnectionTimeout = 60

    
    'NOTE : Do not use the credentials below for sending Emails.
    SMTPUserName = "****" 'To be replaced with the client's SMTP Username.
    SMTPFromEmail = "****" 'To be replaced with the client's email address from which emails will be sent.
    SMTPPassword = "****"            'To be replaced with the client's email password.
    
    
    SMTPToEmail = "****" 'To be replaced with the email address to which emails will be sent.

    CSOEmailAddress = "****"
    MSOEmailAddress = "*****"

End Sub

