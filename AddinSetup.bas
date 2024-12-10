Option Explicit

Public Sub Auto_Open()
    ' This runs when the add-in is loaded
    CreateRibbonButton
End Sub

Private Sub CreateRibbonButton()
    On Error Resume Next
    
    ' Add a button to the ribbon
    Dim ribbonXml As String
    ribbonXml = "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
                "<ribbon><tabs><tab idMso=""TabMail"">" & _
                "<group id=""PrescriptionGroup"" label=""Prescriptions"">" & _
                "<button id=""ProcessPrescriptions"" label=""Process Requests"" " & _
                "onAction=""ProcessPrescriptionRequests"" " & _
                "imageMso=""EnvelopeNew"" size=""large""/>" & _
                "</group></tab></tabs></ribbon></customUI>"
                
    ' Add the button to Outlook's ribbon
    Application.LoadCustomUI "PrescriptionRequestAddin", ribbonXml
End Sub 