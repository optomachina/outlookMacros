Option Explicit

Private Const API_KEY As String = "your-api-key-here"  ' Replace with your actual API key
Private Const API_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const TIMEOUT_SEC As Long = 30  ' Timeout in seconds

Function EscapeJsonString(text As String) As String
    ' Similar to how we handle text in ExtractFormField
    Dim result As String
    result = Replace(text, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, "\", "\\")
    EscapeJsonString = result
End Function

Sub DraftReplyToSelectedEmail()
    On Error Resume Next
    
    Debug.Print "Starting macro..."
    
    Dim outlookApp As Outlook.Application
    Dim selectedMail As Outlook.MailItem
    Dim replyMail As Outlook.MailItem
    Dim mailBody As String
    Dim replyBody As String
    Dim senderName As String
    
    ' Initialize Outlook
    Set outlookApp = Application
    Debug.Print "Initialized Outlook"
    
    If outlookApp.ActiveExplorer.Selection.Count > 0 Then
        Set selectedMail = outlookApp.ActiveExplorer.Selection.Item(1)
        Debug.Print "Selected email subject: " & selectedMail.Subject
        
        ' Get sender's name (first name only)
        senderName = GetSenderFirstName(selectedMail)
        Debug.Print "Sender name: " & senderName
        
        ' Clean up the email body
        mailBody = CleanEmailBody(selectedMail.Body)
        Debug.Print "Cleaned email body"
        
        ' Create API request
        Dim http As Object
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
        
        ' Prepare the message, including sender's name
        Dim requestBody As String
        requestBody = "{""model"": ""gpt-4"", ""messages"": [" & _
                     "{""role"": ""system"", ""content"": ""You are a helpful product support engineer at Edmund Optics. " & _
                     "Draft a professional and friendly response. Include relevant documentation links if applicable. " & _
                     "Address the recipient as '" & senderName & "'. " & _
                     "Keep the tone professional but friendly.""}, " & _
                     "{""role"": ""user"", ""content"": """ & EscapeJsonString(mailBody) & """}]}"
        
        Debug.Print "Sending API request..."
        
        Application.StatusBar = "Generating AI response..."
        DoEvents  ' Allow Outlook to process other events
        
        With http
            .setTimeouts 0, 60000, 60000, 60000  ' Resolve, Connect, Send, Receive timeouts in ms
            .Open "POST", API_URL, False
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "Authorization", "Bearer " & API_KEY
            
            On Error Resume Next
            .send requestBody
            
            If Err.Number <> 0 Then
                Debug.Print "Error during API call: " & Err.Description
                MsgBox "Error calling API: " & Err.Description, vbCritical
                Exit Sub
            End If
            
            Debug.Print "Response Status: " & .Status
            Debug.Print "Full Response: " & .responseText
            
            If .Status = 200 Then
                Debug.Print "API request successful"
                
                ' Parse the response more carefully
                Dim jsonResponse As String
                jsonResponse = .responseText
                
                ' Extract content using more robust string handling
                replyBody = ExtractContentFromJSON(jsonResponse)
                
                If replyBody = "" Then
                    MsgBox "Failed to parse AI response.", vbExclamation
                    Exit Sub
                End If
                
                ' Create reply
                Set replyMail = selectedMail.Reply
                
                ' Set the reply body with proper HTML formatting
                replyMail.HTMLBody = "<html><body>" & _
                                   "<p>" & Replace(replyBody, vbNewLine, "</p><p>") & "</p>" & _
                                   "</body></html>"
                
                ' Display the draft
                replyMail.Display
            Else
                MsgBox "Failed to get response from AI. Status: " & .Status, vbExclamation
            End If
        End With
    Else
        MsgBox "Please select an email to reply to.", vbExclamation
    End If

    ' Cleanup
    Set selectedMail = Nothing
    Set replyMail = Nothing
    Set outlookApp = Nothing
    Set http = Nothing
    
    On Error GoTo 0
End Sub

Function ExtractContentFromJSON(jsonText As String) As String
    On Error Resume Next
    
    Debug.Print "Parsing JSON response..."
    
    ' Look for the specific content path in the JSON structure
    Dim contentStart As Long
    Dim contentEnd As Long
    
    ' Look for the content field in the message structure
    contentStart = InStr(1, jsonText, """content"": """)
    If contentStart = 0 Then
        contentStart = InStr(1, jsonText, """content"":""")  ' Try alternative format
    End If
    
    If contentStart > 0 Then
        contentStart = contentStart + Len("""content"": """)  ' Position after the field name
        
        ' Find the closing quote, accounting for escaped quotes
        Dim pos As Long
        pos = contentStart
        Do
            contentEnd = InStr(pos, jsonText, """")
            If contentEnd = 0 Then Exit Do
            
            ' Check if this quote is escaped
            If Mid(jsonText, contentEnd - 1, 1) = "\" Then
                pos = contentEnd + 1
            Else
                Exit Do
            End If
        Loop
        
        If contentEnd > contentStart Then
            ExtractContentFromJSON = Mid(jsonText, contentStart, contentEnd - contentStart)
            
            ' Unescape special characters
            ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\""", """")
            ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\n", vbNewLine)
            ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\r", vbCrLf)
            ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\/", "/")
            ExtractContentFromJSON = Replace(ExtractContentFromJSON, "\t", vbTab)
            
            Debug.Print "Successfully extracted content"
            Debug.Print "Content length: " & Len(ExtractContentFromJSON)
            Debug.Print "First 100 chars: " & Left(ExtractContentFromJSON, 100)
        End If
    End If
    
    If ExtractContentFromJSON = "" Then
        Debug.Print "Failed to extract content from JSON"
        Debug.Print "JSON Text received: " & Left(jsonText, 200)
    End If
End Function

Function CleanEmailBody(bodyText As String) As String
    ' Remove email chain/signature clutter
    Dim cleanBody As String
    Dim lines() As String
    Dim i As Long
    
    lines = Split(bodyText, vbCrLf)
    cleanBody = ""
    
    ' Take only the most recent message
    For i = 0 To UBound(lines)
        If InStr(lines(i), "From:") > 0 Or _
           InStr(lines(i), "Sent:") > 0 Or _
           InStr(lines(i), "To:") > 0 Then
            Exit For
        End If
        cleanBody = cleanBody & lines(i) & vbCrLf
    Next i
    
    CleanEmailBody = Trim(cleanBody)
End Function

Function GetOpenAIResponse(apiKey As String, messages As String) As String
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    
    ' Use MSXML2.XMLHTTP60 for better HTTPS support in Outlook
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    url = "https://api.openai.com/v1/chat/completions"
    
    requestBody = "{""model"": ""gpt-4"", ""messages"": " & messages & ", ""temperature"": 0.7}"
    
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        
        On Error Resume Next
        .send requestBody
        
        If Err.Number <> 0 Then
            GetOpenAIResponse = "Error: " & Err.Description
            Exit Function
        End If
        
        On Error GoTo 0
        
        If .Status = 200 Then
            GetOpenAIResponse = .responseText
        Else
            GetOpenAIResponse = "Error: " & .responseText
        End If
    End With
End Function

Function ParseOpenAIResponse(jsonResponse As String) As String
    On Error GoTo ErrorHandler
    
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(jsonResponse)
    
    ' The correct path to access the message content in OpenAI's response
    ParseOpenAIResponse = jsonObject("choices")(1)("message")("content")
    Exit Function
    
ErrorHandler:
    ' More detailed error information for debugging
    Debug.Print "Error in ParseOpenAIResponse: " & Err.Description
    Debug.Print "JSON Response: " & jsonResponse
    ParseOpenAIResponse = "Error parsing API response: " & Err.Description
End Function

Function GetSenderFirstName(mail As Outlook.MailItem) As String
    Dim senderName As String
    
    ' Try to get the sender's name
    If mail.SenderName <> "" Then
        senderName = mail.SenderName
        
        ' Remove any extra spaces
        senderName = Trim(senderName)
        
        ' Get first name only (split on space and take first part)
        If InStr(senderName, " ") > 0 Then
            senderName = Left(senderName, InStr(senderName, " ") - 1)
        End If
        
        ' Remove any special characters or titles
        senderName = Replace(senderName, """", "")
        senderName = Replace(senderName, "'", "")
        senderName = Replace(senderName, "Dr.", "")
        senderName = Replace(senderName, "Mr.", "")
        senderName = Replace(senderName, "Ms.", "")
        senderName = Replace(senderName, "Mrs.", "")
        
        ' Trim again in case we removed titles from the start
        senderName = Trim(senderName)
    Else
        senderName = "there"  ' Default fallback
    End If
    
    GetSenderFirstName = senderName
End Function
