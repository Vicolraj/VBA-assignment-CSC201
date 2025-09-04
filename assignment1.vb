' Name: Raji Victor Oluwapelumi
' Department: PROJECT MANAGEMENT TECHNOLOGY
' Matric Number: PMT/22/1543
' QUESTION 1
Private Sub CommandButton1_Click()
    Dim n1 As Double, n2 As Double
    
    If IsNumeric(TextBox1.Text) = False Or IsNumeric(TextBox1.Text) = False Then
        result.Caption = "Please enter valid numbers!"
        Exit Sub
    End If
    
    n1 = CDbl(TextBox1.Text)
    n2 = CDbl(TextBox1.Text)
    
    If n1 < 100 Then
        result.Caption = "First number (" & n1 & ") is small"
    Else
        result.Caption = "First number (" & n1 & ") is large"
    End If
    
    If n2 < 100 Then
        result.Caption = result.Caption & vbCrLf & "Second number (" & n2 & ") is small"
    Else
        result.Caption = result.Caption & vbCrLf & "Second number (" & n2 & ") is large"
    End If
End Sub