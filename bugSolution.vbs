Function f(x)
  If IsNumeric(x) Then
    x = CDbl(x) 'Ensure x is a double
    If x > 10 Then
      f = x + 10
    Else
      f = x - 10
    End If
  Else
    f = "Invalid Input: Not a Number"
  End If
end Function
MsgBox f(20)
MsgBox f("abc")