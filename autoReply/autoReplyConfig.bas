Option Explicit

' Configuration for ChatGPT API
Public Const OPENAI_API_KEY As String = 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  ' Replace with your actual API key
Public Const OPENAI_API_ENDPOINT As String = "https://api.openai.com/v1/chat/completions"
Public Const MODEL_NAME As String = "gpt-3.5-turbo"  ' You can change this to gpt-4 if needed

' Prompt Configuration
Public Function GetPromptTemplate() As String
    ' This is the system prompt that guides ChatGPT's response style
    GetPromptTemplate = "You are an AI assistant helping to draft email responses. " & _
                       "Your responses should be professional, concise, and maintain a friendly tone. " & _
                       "Format the response appropriately for an email context. " & _
                       "Here is the email to respond to:"
End Function 