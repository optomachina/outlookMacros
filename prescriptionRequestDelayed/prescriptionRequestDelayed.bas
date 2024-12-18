Option Explicit

Public olInboxEvents As InboxEvents

Sub InitializeInboxMonitor()
    Set olInboxEvents = New InboxEvents
    olInboxEvents.Initialize
End Sub

Sub ProcessPrescriptionRequestsDelayed()
    ' Add these near the start
    Debug.Print "Starting macro (will process emails older than 10 minutes)..."
    
    On Error Resume Next  ' Add basic error handling
    Dim errorLog As String
    errorLog = ""
    
    ' Add error checking after critical operations
    If Err.Number <> 0 Then
        errorLog = errorLog & "Error: " & Err.Description & vbCrLf
        Debug.Print "Error occurred: " & Err.Description
        Exit Sub
    End If
    
    On Error GoTo 0  ' Resume normal error handling
    
    Dim olNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olItem As Object
    Dim olMail As Outlook.MailItem
    
    ' Get Outlook namespace
    Set olNS = Application.GetNamespace("MAPI")
    
    ' Get inbox folder
    Set olFolder = olNS.GetDefaultFolder(olFolderInbox)
    
    ' Variables for processing
    Dim emailBody As String
    Dim foundCount As Integer
    Dim missingCount As Integer
    Dim partNumbers() As String
    Dim i As Integer
    Dim filePath As String
    Dim fileName As String
    Dim fileNumber As Integer
    Dim replyBody As String
    
    ' Loop through items in inbox
    For Each olItem In olFolder.Items
        If TypeOf olItem Is MailItem Then
            Set olMail = olItem
            Debug.Print "Found email with subject: " & olMail.Subject
            
            ' Check if email is unprocessed, has the correct subject, and is at least 10 minutes old
            If InStr(olMail.Subject, "Edmund Optics Prescription Request") > 0 And _
               olMail.FlagStatus <> olFlagComplete And _
               DateDiff("n", olMail.ReceivedTime, Now) >= 10 Then
                
                Debug.Print "Processing email - matches criteria and is at least 10 minutes old"
                Debug.Print "Email received at: " & olMail.ReceivedTime
                
                ' Mark email as complete and categorize to Blaine
                olMail.Categories = "Blaine"
                olMail.FlagStatus = olFlagComplete
                olMail.Save
                
                ' Extract part numbers from email body
                partNumbers = Split(ExtractPartNumbers(olMail.Body), ",")
                
                ' Reset counters
                foundCount = 0
                missingCount = 0
                emailBody = ""
                
                ' Process each part number
                For i = LBound(partNumbers) To UBound(partNumbers)
                    If Len(Trim(partNumbers(i))) > 0 Then
                        ' Construct file path
                        filePath = "\\us-fs2\Public\Engineering\Zemax Files\Prescriptions\"
                        fileName = Trim(partNumbers(i)) & ".zmx"
                        
                        ' Check if file exists
                        If Dir(filePath & fileName) <> "" Then
                            foundCount = foundCount + 1
                            emailBody = emailBody & partNumbers(i) & vbCrLf
                        Else
                            missingCount = missingCount + 1
                            emailBody = emailBody & partNumbers(i) & " - NOT FOUND" & vbCrLf
                        End If
                    End If
                Next i
                
                ' Create reply based on results
                If foundCount > 0 And missingCount = 0 Then
                    ' All requested files are available
                    If foundCount = 1 Then
                        emailBody = emailBody & _
                            "Attached is the prescription file you requested. This file has been made with the most up-to-date version of Zemax, " & _
                            "and has been checked for accuracy. Please let me know if you have any questions." & vbCrLf & vbCrLf & _
                            "Best regards," & vbCrLf & "Blaine"
                    Else
                        emailBody = emailBody & _
                            "Attached are the prescription files you requested. These files have been made with the most up-to-date version of Zemax, " & _
                            "and have been checked for accuracy. Please let me know if you have any questions." & vbCrLf & vbCrLf & _
                            "Best regards," & vbCrLf & "Blaine"
                    End If
                ElseIf foundCount > 0 And missingCount > 0 Then
                    ' Some files found, some missing
                    emailBody = "I was able to find some of the prescription files you requested:" & vbCrLf & vbCrLf & _
                        emailBody & vbCrLf & _
                        "The files marked as 'NOT FOUND' are not currently available. " & _
                        "I will work on creating these files and send them to you as soon as possible." & vbCrLf & vbCrLf & _
                        "Best regards," & vbCrLf & "Blaine"
                Else
                    ' No files found
                    emailBody = "I was not able to find any of the prescription files you requested:" & vbCrLf & vbCrLf & _
                        emailBody & vbCrLf & _
                        "I will work on creating these files and send them to you as soon as possible." & vbCrLf & vbCrLf & _
                        "Best regards," & vbCrLf & "Blaine"
                End If
                
                ' Create and send reply
                Dim replyMail As Outlook.MailItem
                Set replyMail = olMail.Reply
                
                ' Attach found files
                If foundCount > 0 Then
                    For i = LBound(partNumbers) To UBound(partNumbers)
                        If Len(Trim(partNumbers(i))) > 0 Then
                            fileName = Trim(partNumbers(i)) & ".zmx"
                            If Dir(filePath & fileName) <> "" Then
                                replyMail.Attachments.Add filePath & fileName
                            End If
                        End If
                    Next i
                End If
                
                ' Set reply properties and send
                replyMail.Body = emailBody
                replyMail.Send
                
                Debug.Print "Processed email successfully"
            End If
        End If
    Next
    
    ' Cleanup
    Set olNS = Nothing
    Set olFolder = Nothing
    Set olItem = Nothing
    Set olMail = Nothing
    
    Debug.Print "Macro completed successfully"
    
End Sub

Function ExtractPartNumbers(Body As String) As String
    Dim Lines() As String
    Dim i As Long
    Dim j As Long
    Dim currentLine As String
    Dim partNumber As String
    Dim result As String
    Dim inPartNumberSection As Boolean
    
    ' Split email body into lines
    Lines = Split(Body, vbCrLf)
    inPartNumberSection = False
    result = ""
    
    ' Process each line
    For i = LBound(Lines) To UBound(Lines)
        currentLine = Trim(Lines(i))
        
        ' Check for start of part number section
        If InStr(1, currentLine, "Part Number", vbTextCompare) > 0 Then
            inPartNumberSection = True
            GoTo ContinueLoop
        End If
        
        ' If we're in the part number section
        If inPartNumberSection Then
            ' Skip empty lines
            If Len(currentLine) = 0 Then GoTo ContinueLoop
            
            ' Check if we've reached the end of part numbers
            If InStr(1, currentLine, "Best", vbTextCompare) > 0 Or _
               InStr(1, currentLine, "Thank", vbTextCompare) > 0 Or _
               InStr(1, currentLine, "Regards", vbTextCompare) > 0 Then
                Exit For
            End If
            
            ' Extract part number from line
            partNumber = ""
            For j = 1 To Len(currentLine)
                ' Check if character is a number
                If IsNumeric(Mid(currentLine, j, 1)) Then
                    partNumber = partNumber & Mid(currentLine, j, 1)
                End If
            Next j
            
            ' Add part number to result if valid
            If Len(partNumber) > 0 Then
                If Len(result) > 0 Then result = result & ","
                result = result & partNumber
            End If
        End If
        
ContinueLoop:
    Next i
    
    ExtractPartNumbers = result
End Function
