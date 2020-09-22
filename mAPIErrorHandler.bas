Attribute VB_Name = "mAPIErrorHandler"
Option Explicit

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000 'Specifies that the function should search the system message-table resource(s) for the requested message
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200 'Specifies that insert sequences in the message definition are to be ignored and passed through to the output buffer unchanged. This flag is useful for fetching a message for later formatting. If this flag is set, the Arguments parameter is ignored.

Private Declare Function APIFormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Function LastDLLErrorDescription(LastDLLError As Long) As String
  
  Dim sMessage   As String
  Dim lBuffer     As Long
  Dim lRet        As Long
  
  lBuffer = 256
  sMessage = Space$(lBuffer)
  
  lRet = APIFormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, LastDLLError, 0&, sMessage, lBuffer, 0&)

  If lRet = 0 Then
    sMessage = "FormatMessage API execution Error. Couldn't fetch error description."
  Else
    sMessage = Left$(sMessage, lRet - 2)
  End If
  
  LastDLLErrorDescription = sMessage

End Function

Public Sub CheckForError()
  
  Select Case Err.LastDLLError
    Case 0 ' No error
    Case 2, 3  'The system cannot find the path specified.
      Err.Raise 53, , "File not found"
    Case 5
      'Err.Raise 54, , "Bad file mode"
      Err.Raise 70, , "Permission denied"
    Case 18 'There are no more files. (Do nothing)
    Case 21
      Err.Raise 71, , "Disk not ready"
    Case 32 'The process cannot access the file because it is being used by another process.
      Err.Raise 70, , "Permission denied"
    Case 80, 183
      Err.Raise 58, , "File already exists"
    Case 234 'More data is available. (Do nothing)
    Case Else
      'Debug.Print Err.LastDLLError
      'Debug.Print LastDLLErrorDescription(Err.LastDLLError)
      'Stop
      Err.Raise Err.LastDLLError, , LastDLLErrorDescription(Err.LastDLLError)
  End Select
  
End Sub

