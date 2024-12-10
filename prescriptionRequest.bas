Option Explicit

Sub ProcessPrescriptionRequests()
    ' Add these near the start
    Debug.Print "Starting macro..."
    
    On Error Resume Next  ' Add basic error handling
    Dim errorLog As String
    errorLog = ""
    
    ' Add error checking after critical operations
    If Err.Number <> 0 Then
        errorLog = errorLog & "Error: " & Err.Description & vbCrLf
        Debug.Print "Error occurred: " & Err.Description
        Err.Clear
    End If
    
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.folder
    Dim olItem As Outlook.MailItem
    Dim olMail As Outlook.MailItem
    Dim FilePath As String
    Dim Attachments As String
    Dim UnavailableParts As String
    Dim Parts() As String
    Dim Part As Variant
    Dim Found As Boolean
    Dim RecipientEmail As String
    Dim fso As Object
    Dim prescFolder As Object
    Dim prescFile As Variant
    Dim ResponseMail As Outlook.MailItem
    Dim Attachment As Variant
    
    ' Set paths
    FilePath = "Q:\Released Documents\Optical Prescriptions\Zemax Files\Black box\"
    
    ' Access the shared inbox
    Set olNS = Application.GetNamespace("MAPI")
    Debug.Print "Got MAPI namespace"
    
    Set olFolder = olNS.Folders("Technical Support USA").Folders("Inbox")
    If Err.Number <> 0 Then
        Debug.Print "Error accessing folder: " & Err.Description
        MsgBox "Error accessing folder: " & Err.Description
        Exit Sub
    End If
    Debug.Print "Found folder: " & olFolder.Name
    
    ' Add a check if folder exists
    If olFolder Is Nothing Then
        MsgBox "Could not find specified folder. Please check the folder path.", vbExclamation
        Exit Sub
    End If
    
    ' After folder check
    Debug.Print "Checking folder: " & olFolder.FolderPath
    
    ' Loop through each email
    For Each olItem In olFolder.Items
        If TypeOf olItem Is MailItem Then
            Set olMail = olItem
            Debug.Print "Found email with subject: " & olMail.Subject
            Debug.Print "Categories: " & olMail.Categories
            
            ' Change the criteria to look for prescription requests
            If InStr(olMail.Subject, "Edmund Optics Prescription Request") > 0 And _
               InStr(olMail.Categories, "Blaine") > 0 Then
                Debug.Print "Processing email - matches criteria for Blaine"
                
                ' Extract the recipient email from the body
                RecipientEmail = ExtractRecipientEmail(olMail.Body)
                
                ' Extract product numbers from email body (Question #2)
                Parts = Split(ExtractPartNumbers(olMail.Body), ",")
                Attachments = ""
                UnavailableParts = ""
                
                ' Check each product number
                For Each Part In Parts
                    Part = Trim(Part)
                    Found = False
                    
                    ' Search for the file
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    Set prescFolder = fso.GetFolder(FilePath)
                    
                    ' In the file search loop, add file existence check
                    If Not fso.FolderExists(FilePath) Then
                        MsgBox "Prescription files folder not found: " & FilePath, vbExclamation
                        Exit Sub
                    End If
                    
                    For Each prescFile In prescFolder.Files
                        Debug.Print "Checking file: " & prescFile.Name & " against part: " & Part
                        ' Check if file starts with the part number
                        If Left(prescFile.Name, Len(Part)) = Part Then
                            Debug.Print "Found matching file: " & prescFile.Name & " for part: " & Part
                            Attachments = Attachments & FilePath & prescFile.Name & ";"
                            Found = True
                            Exit For
                        End If
                    Next
                    
                    ' Add to unavailable list if not found
                    If Not Found Then
                        UnavailableParts = UnavailableParts & Part & ", "
                    End If
                Next
                
                ' Create the response email and set properties before displaying
                Set ResponseMail = olMail.Reply
                
                ' Set the sender before displaying
                ResponseMail.SendUsingAccount = Application.Session.Accounts.Item("bwilson@edmundoptics.com")
                ResponseMail.SentOnBehalfOfName = "bwilson@edmundoptics.com"
                
                ' Display to get the signature
                ResponseMail.Display
                
                ' Clear the default recipient and set the correct one
                ResponseMail.To = ""  ' Clear default recipient
                RecipientEmail = ExtractFormField(olMail.Body, "Email Address:")
                ResponseMail.To = RecipientEmail
                Debug.Print "Setting recipient to: " & RecipientEmail
                
                ' Extract the first name from the form submission
                Dim FirstName As String
                FirstName = ExtractFormField(olMail.Body, "First Name:")
                Debug.Print "Extracted First Name: " & FirstName
                
                ' Compose the body
                Dim emailBody As String
                emailBody = "Hi " & FirstName & "," & vbNewLine & vbNewLine & _
                           "Thank you for contacting Edmund Optics." & vbNewLine & vbNewLine & _
                           "Attached is the prescription file you requested. This file has been made with the most up-to-date version of Zemax, " & _
                           "so if you do encounter issues, be sure to check which version you are running. Please note this is a Zemax archive file, " & _
                           "so you will need to open it by clicking on File and then ""Load Archive"". " & _
                           "Please let me know if you have any questions. Have a great day!" & vbNewLine  ' Single line break before signature
                
                ' If there are unavailable parts, add them to the email
                If UnavailableParts <> "" Then
                    UnavailableParts = Left(UnavailableParts, Len(UnavailableParts) - 2)
                    emailBody = emailBody & "Note: The following part numbers were not found: " & UnavailableParts & vbNewLine
                End If
                
                ' Set the body while preserving signature
                ResponseMail.GetInspector.WordEditor.Range(0, 0).InsertBefore emailBody
                
                ' Attach files
                If Len(Attachments) > 0 Then
                    Dim AttachPaths() As String
                    AttachPaths = Split(Attachments, ";")
                    
                    For Each Attachment In AttachPaths
                        If Len(Attachment) > 0 Then
                            Debug.Print "Attempting to attach: " & Attachment
                            On Error Resume Next
                            ResponseMail.Attachments.Add Attachment
                            If Err.Number <> 0 Then
                                Debug.Print "Error attaching file: " & Err.Description
                                errorLog = errorLog & "Failed to attach: " & Attachment & " - " & Err.Description & vbCrLf
                                Err.Clear
                            Else
                                Debug.Print "Successfully attached: " & Attachment
                            End If
                            On Error GoTo 0
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' Cleanup
    Set olNS = Nothing
    Set olFolder = Nothing
    Set olMail = Nothing
    
    On Error GoTo 0  ' Turn off error handling
    
    ' At the end of the macro
    If Len(errorLog) > 0 Then
        Debug.Print "Errors occurred during execution:" & vbCrLf & errorLog
        MsgBox "Some errors occurred. Check the immediate window for details.", vbExclamation
    End If
End Sub

Function ExtractPartNumbers(Body As String) As String
    Dim Lines() As String
    Dim i As Long
    Dim PartNumbers As String
    
    Debug.Print "Searching for part numbers in body:"
    
    ' Split body into lines
    Lines = Split(Body, vbCrLf)
    
    ' Look for question 2
    For i = 0 To UBound(Lines)
        If InStr(Lines(i), "2. What are the stock numbers") > 0 Then
            Debug.Print "Found stock numbers line: " & Lines(i)
            ' Extract everything after the colon and tab
            If InStr(Lines(i), vbTab) > 0 Then
                PartNumbers = Trim(Split(Lines(i), vbTab)(1))
            ElseIf InStr(Lines(i), ":") > 0 Then
                PartNumbers = Trim(Split(Lines(i), ":")(1))
            End If
            Debug.Print "Extracted part numbers: " & PartNumbers
            Exit For
        End If
    Next
    
    ' Clean up the part numbers (remove spaces if any)
    PartNumbers = Replace(PartNumbers, " ", "")
    Debug.Print "Final part numbers string: " & PartNumbers
    ExtractPartNumbers = PartNumbers
End Function

Function ExtractRecipientEmail(Body As String) As String
    Dim RegExp As Object
    Dim Matches As Object
    Dim Match As Object
    
    ' Regular expression to extract email address
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Pattern = "[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}"
    RegExp.Global = False ' We only need the first match
    
    ' Find email in the body
    If RegExp.Test(Body) Then
        Set Matches = RegExp.Execute(Body)
        ExtractRecipientEmail = Matches(0) ' Return the first match
    Else
        ExtractRecipientEmail = "" ' Return empty string if not found
    End If
End Function

Function ExtractFormField(Body As String, FieldName As String) As String
    Dim Lines() As String
    Dim i As Long
    
    ' Split the body into lines
    Lines = Split(Body, vbCrLf)
    Debug.Print "Searching for field: " & FieldName
    
    ' Look for the field
    For i = 0 To UBound(Lines)
        If InStr(Lines(i), FieldName) > 0 Then
            Debug.Print "Found line: " & Lines(i)
            ' Try different delimiters
            If InStr(Lines(i), vbTab) > 0 Then
                ExtractFormField = Trim(Split(Lines(i), vbTab)(1))
            ElseIf InStr(Lines(i), ":") > 0 Then
                ExtractFormField = Trim(Split(Lines(i), ":")(1))
            End If
            Debug.Print "Extracted value: " & ExtractFormField
            Exit Function
        End If
    Next i
    
    ' Return empty string if not found
    ExtractFormField = ""
    Debug.Print "Field not found"
End Function

Function GetDefaultSignature() As String
    Dim appDataPath As String
    Dim signaturePath As String
    Dim fso As Object
    Dim ts As Object
    Dim signatureContent As String
    
    ' Get the path to signature files
    appDataPath = Environ("APPDATA")
    signaturePath = appDataPath & "\Microsoft\Signatures\"
    
    Debug.Print "Looking for signature in: " & signaturePath
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Try different possible filenames
    If fso.FileExists(signaturePath & "Default.htm") Then
        Debug.Print "Found signature file: Default.htm"
        Set ts = fso.OpenTextFile(signaturePath & "Default.htm", 1)
        signatureContent = ts.ReadAll
        ts.Close
        
        ' Fix image paths in the signature
        signatureContent = Replace(signatureContent, "src=""Default_files/", "src=""" & signaturePath & "Default_files/")
        GetDefaultSignature = signatureContent
        
    ElseIf fso.FileExists(signaturePath & "default.htm") Then
        Debug.Print "Found signature file: default.htm"
        Set ts = fso.OpenTextFile(signaturePath & "default.htm", 1)
        signatureContent = ts.ReadAll
        ts.Close
        
        ' Fix image paths
        signatureContent = Replace(signatureContent, "src=""default_files/", "src=""" & signaturePath & "default_files/")
        GetDefaultSignature = signatureContent
        
    Else
        Debug.Print "No signature file found"
        GetDefaultSignature = ""
    End If
    
    Debug.Print "Signature length: " & Len(GetDefaultSignature)
    
    Set fso = Nothing
    Set ts = Nothing
End Function