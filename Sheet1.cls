VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub btnSend_Click()
    ' Getters
    Dim strSentOnBehalfOfName, strTo, strCc, strBcc, strSubject, strHtmlBody, strAttachment As String
    strSentOnBehalfOfName = Range("C3")
    strTo = Range("C4")
    strCc = Range("C5")
    strBcc = Range("C6")
    strSubject = Range("C7")
    strHtmlBody = Range("C8")
    strAttachment = Range("C9")
    ' Business Logic
    CDO_Mail CStr(strSentOnBehalfOfName), CStr(strTo), CStr(strCc), CStr(strBcc), CStr(strSubject), CStr(strHtmlBody), CStr(strAttachment)
End Sub
