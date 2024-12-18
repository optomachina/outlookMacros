Private WithEvents olItems As Outlook.Items
Private WithEvents olInbox As Outlook.folder

Public Sub Initialize()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    
    Set olApp = Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olInbox = olNS.Folders("Technical Support USA").Folders("Inbox")
    Set olItems = olInbox.Items
End Sub

Private Sub olItems_ItemAdd(ByVal Item As Object)
    If TypeOf Item Is Outlook.MailItem Then
        If InStr(Item.Subject, "Edmund Optics Prescription Request") > 0 Then
            ProcessPrescriptionRequests
        End If
    End If
End Sub

Private Sub olItems_ItemChange(ByVal Item As Object)
    On Error Resume Next
    
    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item
        
        If Err.Number = 0 Then  ' Check if we can access the mail item
            If InStr(mail.Subject, "Edmund Optics Prescription Request") > 0 And _
               mail.Categories = "Blaine" Then
                Debug.Print "Processing changed item: " & mail.Subject
                ProcessPrescriptionRequests
            End If
        Else
            Debug.Print "Error accessing mail item: " & Err.Description
            Err.Clear
        End If
    End If
    
    On Error GoTo 0
End Sub 