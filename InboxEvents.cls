Private WithEvents olItems As Outlook.Items

Public Sub Initialize()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.folder
    
    Set olApp = Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.Folders("Technical Support USA").Folders("Inbox")
    Set olItems = olFolder.Items
End Sub

Private Sub olItems_ItemAdd(ByVal Item As Object)
    If TypeOf Item Is Outlook.MailItem Then
        If InStr(Item.Subject, "Edmund Optics Prescription Request") > 0 Then
            ProcessPrescriptionRequests
        End If
    End If
End Sub 