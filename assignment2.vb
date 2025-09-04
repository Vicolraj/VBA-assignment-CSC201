' Name: Raji Victor Oluwapelumi
' Department: PROJECT MANAGEMENT TECHNOLOGY
' Matric Number: PMT/22/1543
' QUESTION 2
Private Sub UserForm_Activate()
    Dim x As Double, y As Double
    Dim val1 As Double, val2 As Double, val3 As Double
    
    x = 9
    y = 13
    
    val1 = 2 * x + y
    val2 = (x + y) - 2
    val3 = x * y
    
    MsgBox "For x = " & x & " and y = " & y & vbCrLf & _
           "2x + y = " & val1 & vbCrLf & _
           "(x + y) - 2 = " & val2 & vbCrLf & _
           "x * y = " & val3, vbInformation, "Results"
End Sub
