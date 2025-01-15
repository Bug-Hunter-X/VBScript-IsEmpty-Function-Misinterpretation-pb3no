Function MyFunction(param1)
  If param1 = "" Then
    ' Handle empty string parameter
    param1 = "Default Value"
  ElseIf IsArray(param1) And UBound(param1) < 0 Then
    ' Handle empty array parameter
    param1 = Array("Default Value")
  End If
  ' ... rest of the function
End Function