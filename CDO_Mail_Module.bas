Attribute VB_Name = "Email"
'-------------------------------------------------------------------
'
'   Sends and email to the specified recipient, cc and bcc - with the provided content
'   VIA SMTP DIRECTLY
'
'   @Input  -   SentOnBehalfOfName
'   @Input  -   To
'   @Input  -   CC
'   @Input  -   BCC
'   @Input  -   Subject
'   @Input  -   HTMLBody
'
'-------------------------------------------------------------------
Public Sub CDO_Mail(strSentOnBehalfOfName As String, strTo As String, strCc As String, strBcc As String, strSubject As String, strHtmlBody As String, strAttachment As String)
    
    Dim iConf As Object
    Dim Flds As Variant
    Dim iMsg As Object

    Set iConf = CreateObject("CDO.Configuration")
    Set iMsg = CreateObject("CDO.Message")

    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2   ' cdoSendUsingPort, value 2, for sending the message using the network.
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CStr(Sheets("SMTP").Range("C3"))
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CStr(Sheets("SMTP").Range("C4"))
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = CInt(Sheets("SMTP").Range("C5"))
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = CStr(Sheets("SMTP").Range("C6"))
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = CStr(Sheets("SMTP").Range("C7"))
        .Update
    End With

    With iMsg
        Set .Configuration = iConf
        .To = strTo
        .CC = strCc
        .BCC = strBcc
        .From = strSentOnBehalfOfName
        .Subject = strSubject
        .HTMLBody = strHtmlBody
        If checkFileExists(strAttachment) Then
            .AddAttachment strAttachment
        End If
        .Send
    End With

    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
End Sub


Function checkFileExists(file As String)
    If Dir(file) <> "" Then
        checkFileExists = True
    Else
        checkFileExists = False
    End If
End Function

