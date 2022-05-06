VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public Sub SaveHuiDan1(item As Outlook.MailItem)
    Call SubClosingPopUp(2, "received email: " + item.Subject, "debug")
    
    Dim object_attachment As Outlook.Attachment
    
    Dim saveFolder  As String
    saveFolder = "C:\Data\Archive"
    
    For Each object_attachment In item.Attachments
        Dim filename As String
        filename = saveFolder & "\" & Format(Now(), "yyyy-MM-dd-hh-mm-ss-ms") & Rnd() & object_attachment.DisplayName
                      
        Call SubClosingPopUp(2, "will save attached file to: " + filename, "debug")
        
        object_attachment.SaveAsFile filename
    Next
    
    Call SubClosingPopUp(2, "finished", "debug")
    
End Sub