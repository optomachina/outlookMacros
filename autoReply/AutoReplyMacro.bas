Option Explicit

Private Type WINHTTP_PROXY_INFO
    dwAccessType As Long
    lpszProxy As String
    lpszProxyBypass As String
End Type

#If VBA7 Then
    Private Declare PtrSafe Function WinHttpOpen Lib "winhttp" Alias "WinHttpOpen" ( _
        ByVal lpszAgent As LongPtr, _
        ByVal dwAccessType As Long, _
        ByVal lpszProxy As LongPtr, _
        ByVal lpszProxyBypass As LongPtr, _
        ByVal dwFlags As Long) As LongPtr
        
    Private Declare PtrSafe Function WinHttpConnect Lib "winhttp" Alias "WinHttpConnect" ( _
        ByVal hSession As LongPtr, _
        ByVal lpszServerName As LongPtr, _
        ByVal nServerPort As Long, _
        ByVal dwReserved As Long) As LongPtr
        
    Private Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp" Alias "WinHttpOpenRequest" ( _
        ByVal hConnect As LongPtr, _
        ByVal lpszVerb As LongPtr, _
        ByVal lpszObjectName As LongPtr, _
        ByVal lpszVersion As LongPtr, _
        ByVal lpszReferrer As LongPtr, _
        ByVal lpszAcceptTypes As LongPtr, _
        ByVal dwFlags As Long) As LongPtr
        
    Private Declare PtrSafe Function WinHttpSendRequest Lib "winhttp" Alias "WinHttpSendRequest" ( _
        ByVal hRequest As LongPtr, _
        ByVal lpszHeaders As LongPtr, _
        ByVal dwHeadersLength As Long, _
        ByVal lpOptional As LongPtr, _
        ByVal dwOptionalLength As Long, _
        ByVal dwTotalLength As Long, _
        ByVal dwContext As LongPtr) As Long
        
    Private Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp" Alias "WinHttpReceiveResponse" ( _
        ByVal hRequest As LongPtr, _
        ByVal lpReserved As LongPtr) As Long
        
    Private Declare PtrSafe Function WinHttpQueryDataAvailable Lib "winhttp" Alias "WinHttpQueryDataAvailable" ( _
        ByVal hRequest As LongPtr, _
        ByRef lpdwNumberOfBytesAvailable As Long) As Long
        
    Private Declare PtrSafe Function WinHttpReadData Lib "winhttp" Alias "WinHttpReadData" ( _
        ByVal hRequest As LongPtr, _
        ByVal lpBuffer As LongPtr, _
        ByVal dwNumberOfBytesToRead As Long, _
        ByRef lpdwNumberOfBytesRead As Long) As Long
#End If

Public Sub AutoReplyToEmail()
    Dim objMail As Outlook.MailItem
    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    If objMail Is Nothing Then
        MsgBox "Please select an email first.", vbExclamation
        Exit Sub
    End If
    
    ' Get the body of the selected email
    Dim emailBody As String
    emailBody = objMail.Body
    
    ' Prepare the prompt with the email body
    Dim fullPrompt As String
    fullPrompt = GetPromptTemplate() & vbNewLine & vbNewLine & emailBody
    
    ' Call ChatGPT API
    Dim response As String
    response = CallChatGPTAPI(fullPrompt)
    
    ' Create a reply
    Dim replyMail As Outlook.MailItem
    Set replyMail = objMail.Reply
    
    ' Set the response as the body of the reply
    replyMail.Body = response
    
    ' Display the draft reply (but don't send it automatically)
    replyMail.Display
End Sub

Private Function CallChatGPTAPI(prompt As String) As String
    ' For testing: Return a mock response instead of calling the API
    CallChatGPTAPI = "Dear [Sender]," & vbNewLine & vbNewLine & _
                     "Thank you for your email. This is a test response to demonstrate the auto-reply functionality." & vbNewLine & vbNewLine & _
                     "Here's what the macro did:" & vbNewLine & _
                     "1. Captured your email content" & vbNewLine & _
                     "2. Would normally send it to ChatGPT" & vbNewLine & _
                     "3. Created this draft reply" & vbNewLine & vbNewLine & _
                     "Original email content was:" & vbNewLine & _
                     "-------------------" & vbNewLine & _
                     prompt & vbNewLine & _
                     "-------------------" & vbNewLine & vbNewLine & _
                     "Best regards," & vbNewLine & _
                     "Auto-Reply Test"
End Function

Private Function MakeHttpRequest(url As String, method As String, body As String, apiKey As String) As String
    #If VBA7 Then
        Dim hOpen As LongPtr
        Dim hConnect As LongPtr
        Dim hRequest As LongPtr
    #Else
        Dim hOpen As Long
        Dim hConnect As Long
        Dim hRequest As Long
    #End If
    
    Dim response As String
    Dim responseChunk As String * 1024
    Dim responseLength As Long
    Dim headers As String
    
    ' Parse URL
    Dim urlParts() As String
    urlParts = Split(Replace(url, "https://", ""), "/")
    Dim host As String
    host = urlParts(0)
    Dim path As String
    path = "/" & Join(Application.Index(urlParts, Array(2, UBound(urlParts) + 1)), "/")
    
    ' Set up headers
    headers = "Content-Type: application/json" & vbCrLf & _
             "Authorization: Bearer " & apiKey
    
    ' Initialize WinHttp
    hOpen = WinHttpOpen(StrPtr("VBA/1.0"), 0, 0, 0, 0)
    If hOpen = 0 Then GoTo CleanUp
    
    ' Connect to server
    hConnect = WinHttpConnect(hOpen, StrPtr(host), 443, 0)
    If hConnect = 0 Then GoTo CleanUp
    
    ' Create request
    hRequest = WinHttpOpenRequest(hConnect, StrPtr(method), StrPtr(path), _
                                 StrPtr("HTTP/1.1"), 0, 0, &H800000) ' WINHTTP_FLAG_SECURE
    If hRequest = 0 Then GoTo CleanUp
    
    ' Send request
    If WinHttpSendRequest(hRequest, StrPtr(headers), Len(headers), StrPtr(body), Len(body), Len(body), 0) = 0 Then
        GoTo CleanUp
    End If
    
    ' Wait for response
    If WinHttpReceiveResponse(hRequest, 0) = 0 Then GoTo CleanUp
    
    ' Read response
    Do
        responseLength = 0
        If WinHttpQueryDataAvailable(hRequest, responseLength) = 0 Then Exit Do
        If responseLength = 0 Then Exit Do
        
        If WinHttpReadData(hRequest, StrPtr(responseChunk), Len(responseChunk), responseLength) = 0 Then
            Exit Do
        End If
        
        response = response & Left$(responseChunk, responseLength)
    Loop
    
CleanUp:
    MakeHttpRequest = response
End Function

Private Function JsonEscape(text As String) As String
    Dim result As String
    result = Replace(text, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    JsonEscape = result
End Function

Private Function JsonUnescape(text As String) As String
    Dim result As String
    result = Replace(text, "\n", vbNewLine)
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    JsonUnescape = result
End Function 