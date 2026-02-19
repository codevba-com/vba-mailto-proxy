Attribute VB_Name = "MailToProxy"
Option Explicit
' =============================================================================
' Module:        MailToProxy
' Author:        Mark Uildriks, codevba.com
' Description:   Creates email in default mail client in edit mode using MailTo. Provides fallbacks to Gmail/Outlook.com
' Comment:       User needs to press Send. For a more advanced solution see https://www.codevba.com/vba-mailer/
' Dependencies:  None
' Office version 2016 and higher
' License:       MIT License
' Version        1.0
' Repository:    https://github.com/codevba-com/vba-email-composer
' =============================================================================
#If VBA7 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

' Create Email - first try mailto, if that fails online Outlook or Gmail
Public Sub CreateEmail(To_ As String, Subject As String, ByVal Body As String)
    Dim mailtoUri As String
    Dim sep As String
    sep = "?"
    
    ' --- Build mailto URI ---
    mailtoUri = "mailto:" & To_
    
    '    If cc <> "" Then
    '        mailtoUri = mailtoUri & sep & "cc=" & UrlEncodeUtf8(cc)
    '        sep = "&"
    '    End If
    '    If bcc <> "" Then
    '        mailtoUri = mailtoUri & sep & "bcc=" & UrlEncodeUtf8(bcc)
    '        sep = "&"
    '    End If
    
    If Subject <> "" Then
        mailtoUri = mailtoUri & sep & "subject=" & UrlEncodeUtf8(Subject)
        sep = "&"
    End If
    
    If Body <> "" Then
        mailtoUri = mailtoUri & sep & "body=" & UrlEncodeUtf8(Body)
    End If
    
    ' --- 1. Try local mail client via ShellExecute ---
    If IsMailClientConfigured() Then
        Dim result As LongPtr
        result = ShellExecute(0, "open", mailtoUri, vbNullString, vbNullString, 1)
        
        If result > 32 Then Exit Sub
        ' If ShellExecute fails, fall through to webmail
    End If
    
    ' --- 2. Gmail fallback ---
    Dim gmailUrl As String
    gmailUrl = "https://mail.google.com/mail/?view=cm&fs=1" & _
    "&to=" & UrlEncodeUtf8(To_) & _
    "&su=" & UrlEncodeUtf8(Subject) & _
    "&body=" & UrlEncodeUtf8(Body)
    
    '    If cc <> "" Then gmailUrl = gmailUrl & "&cc=" & UrlEncodeUtf8(cc)
    '    If bcc <> "" Then gmailUrl = gmailUrl & "&bcc=" & UrlEncodeUtf8(bcc)
    
    On Error Resume Next
    Shell "cmd /c start """" """ & gmailUrl & """", vbHide
    If Err.Number = 0 Then Exit Sub
    
    ' --- 3. Outlook.com fallback ---
    Dim outlookUrl As String
    outlookUrl = "https://outlook.office.com/mail/deeplink/compose?" & _
    "to=" & UrlEncodeUtf8(To_) & _
    "&subject=" & UrlEncodeUtf8(Subject) & _
    "&body=" & UrlEncodeUtf8(Body)
    
    '    If cc <> "" Then outlookUrl = outlookUrl & "&cc=" & UrlEncodeUtf8(cc)
    '    If bcc <> "" Then outlookUrl = outlookUrl & "&bcc=" & UrlEncodeUtf8(bcc)
    
    Shell "cmd /c start """" """ & outlookUrl & """", vbHide
End Sub

Private Function UrlEncodeUtf8(ByVal s As String) As String
    Dim i As Long, code As Long
    Dim utf8() As Byte
    Dim b As Byte
    
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        
        Select Case code
        Case 0 To &H7F
            ' 1-byte UTF-8
            UrlEncodeUtf8 = UrlEncodeUtf8 & "%" & Right$("0" & Hex(code), 2)
            
        Case &H80 To &H7FF
            ' 2-byte UTF-8
            UrlEncodeUtf8 = UrlEncodeUtf8 & _
            "%" & Right$("0" & Hex(&HC0 Or (code \ 64)), 2) & _
            "%" & Right$("0" & Hex(&H80 Or (code And &H3F)), 2)
            
        Case Else
            ' 3-byte UTF-8
            UrlEncodeUtf8 = UrlEncodeUtf8 & _
            "%" & Right$("0" & Hex(&HE0 Or (code \ 4096)), 2) & _
            "%" & Right$("0" & Hex(&H80 Or ((code \ 64) And &H3F)), 2) & _
            "%" & Right$("0" & Hex(&H80 Or (code And &H3F)), 2)
        End Select
    Next i
End Function

Private Function IsMailClientConfigured() As Boolean
    On Error Resume Next
    Dim wsh As Object, cmd As String
    Set wsh = CreateObject("WScript.Shell")
    cmd = wsh.RegRead("HKEY_CLASSES_ROOT\mailto\shell\open\command\")
    IsMailClientConfigured = (Trim$(cmd) <> "")
End Function

