Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 13, , "Both parameters must be provided"
  End If
  ' ... rest of your function code ...
End Function