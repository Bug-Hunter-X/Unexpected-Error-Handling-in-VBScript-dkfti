Function MyFunction(param1, param2)
  On Error GoTo ErrorHandler

  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 13, , "Both parameters must be provided"
  End If

  ' ... rest of your function code ...

  Exit Function

ErrorHandler:
  MsgBox "An error occurred: " & Err.Description, vbCritical
  Err.Clear
End Function