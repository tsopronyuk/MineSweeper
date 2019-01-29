'matrix within buttons
Dim ButtonCellArray(1 To 5, 1 To 5) As New ClassForEventCmd

'string with 10 random numbers, separeted by comma
Public mines10 As String



Private Sub UserForm_Initialize()
Dim i, j, n As Integer

Dim elem As MSForms.CommandButton 'MyCell

'create random ships
Call shuffle

'set blue buttons 1,2,...,25 to the form
n = 1
For i = 1 To 5
    For j = 1 To 5
    
    Set elem = Controls.Add("Forms.CommandButton.1", "btnCell" & n)
    
    With elem
        .Height = 40
        .Width = 40
        .Left = (j - 1) * .Width
        .Top = (i - 1) * .Height
        .Caption = n
        .BackColor = vbBlue
        .ForeColor = vbBlue
    
    End With
    Set ButtonCellArray(i, j).cmdCell = elem
    
    n = n + 1
    Next
Next

End Sub

Sub shuffle()
    'Create array of unique random number between 1 and 25
    Dim mines(1 To 25) As String
    mines10 = ""
    For i = 1 To 25
        mines(i) = Str(i)
    Next i

    'shuffle numbers
    For i = 1 To 25
         r = Int(25 * Rnd + 1)
         temp = mines(r)
         mines(r) = mines(i)
         mines(i) = temp
    Next i
    
    'create string with 10 random numbers, separeted by comma
    For i = 1 To 10
    mines10 = mines10 + "," + mines(i)
    'mines10 = mines10 & "," & i ' test
    Next
    
'mines10 = " ,2,8,4,11,12,17,16,21,22,6"  ' test
'MsgBox mines10
End Sub



Private Sub txtCheck_Change()

'if there are no mines surrrounding the selected cell
    If Trim(txtCheck.Text) = Trim("0") Then
    
        'find indexes for the cell
        For i = 1 To 5
            For j = 1 To 5
                If Trim(ButtonCellArray(i, j).cmdCell.Caption) = Trim("0") Then
                indI = i
                indJ = j
                GoTo m
                End If
            Next j
        Next i
        
        'open all surrrounding cells. Set Green color for these cells
m:
        ButtonCellArray(indI, indJ).cmdCell.Caption = "-1"
        ButtonCellArray(indI, indJ).cmdCell.BackColor = vbGreen
        ButtonCellArray(indI, indJ).cmdCell.ForeColor = vbGreen
        
        If indI + 1 < 6 And indJ + 1 < 6 Then
            ButtonCellArray(indI + 1, indJ + 1).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI + 1, indJ + 1).cmdCell.ForeColor = vbGreen
        End If
        
        If indJ + 1 < 6 Then
            ButtonCellArray(indI, indJ + 1).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI, indJ + 1).cmdCell.ForeColor = vbGreen
        End If
        
        If indI - 1 > 0 And indJ + 1 < 6 Then
            ButtonCellArray(indI - 1, indJ + 1).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI - 1, indJ + 1).cmdCell.ForeColor = vbGreen
        End If
        
        If indI + 1 < 6 Then
            ButtonCellArray(indI + 1, indJ).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI + 1, indJ).cmdCell.ForeColor = vbGreen
        End If
        
        If indI + 1 < 6 And indJ - 1 > 0 Then
            ButtonCellArray(indI + 1, indJ - 1).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI + 1, indJ - 1).cmdCell.ForeColor = vbGreen
        End If
        
        If indI - 1 > 0 And indJ - 1 > 0 Then
            ButtonCellArray(indI - 1, indJ - 1).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI - 1, indJ - 1).cmdCell.ForeColor = vbGreen
        End If
        
        If indI - 1 > 0 Then
            ButtonCellArray(indI - 1, indJ).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI - 1, indJ).cmdCell.ForeColor = vbGreen
        End If
        
        If indJ - 1 > 0 Then
            ButtonCellArray(indI, indJ - 1).cmdCell.BackColor = vbGreen
            ButtonCellArray(indI, indJ - 1).cmdCell.ForeColor = vbGreen
        End If
        
            
        
        txtCheck.Text = "-1"
    End If
    
    'Calculate green cells
    CountGreen = 0
    For i = 1 To 5
        For j = 1 To 5
            If ButtonCellArray(i, j).cmdCell.BackColor = vbGreen Then
                CountGreen = CountGreen + 1
            End If
        Next j
    Next i
    
    'Check winner
    If CountGreen = 15 Then
        MsgBox "You win!"
        End
    End If
    
End Sub

