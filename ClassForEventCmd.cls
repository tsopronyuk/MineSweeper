VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassForEventCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents cmdCell As MSForms.CommandButton
Attribute cmdCell.VB_VarHelpID = -1

'board within mines
Dim board(0 To 6, 0 To 6) As String

'matrix  with positions corresponding the matrix ButtonCellArray
Dim cells(0 To 6, 0 To 6) As String

Private Sub cmdCell_click()
    Dim mines() As String
    
    'create array of mines positions
    mines = Split(Trim(UserForm1.mines10), ",")
    
    'create board within mines
    For i = 0 To 6
        For j = 0 To 6
            board(i, j) = "0"
        Next
    Next
    
    n = 1
    For i = 1 To 5
        For j = 1 To 5
            For k = 1 To 10
                If Trim(Str(n)) = Trim(mines(k)) Then
                    board(i, j) = n
                End If
            Next k
                
            n = n + 1
        Next j
    Next i
    
    'fill matrix positions
    n = 1
    For i = 1 To 5
        For j = 1 To 5
            cells(i, j) = n
            n = n + 1
        Next j
    Next i
    
    'find mines in the current button
    For i = 1 To 10
    If cmdCell.BackColor = vbBlue And Trim(cmdCell.Caption) = Trim(mines(i)) Then
        cmdCell.BackColor = vbRed
        cmdCell.ForeColor = vbRed
        MsgBox "You lose!"
        
        End
        End If
    Next
    
    'find indexes of current button
    For i = 1 To 5
        For j = 1 To 5
            If Trim(cells(i, j)) = Trim(cmdCell.Caption) Then
            indI = i
            indJ = j
            GoTo m
            End If
        Next j
    Next i
    
'calculate the neighboring mines
m: Count = 0


If Val(board(indI + 1, indJ + 1)) > 0 Then Count = Count + 1
If Val(board(indI, indJ + 1)) > 0 Then Count = Count + 1
If Val(board(indI - 1, indJ + 1)) > 0 Then Count = Count + 1
If Val(board(indI + 1, indJ)) > 0 Then Count = Count + 1
If Val(board(indI + 1, indJ - 1)) > 0 Then Count = Count + 1
If Val(board(indI - 1, indJ - 1)) > 0 Then Count = Count + 1
If Val(board(indI - 1, indJ)) > 0 Then Count = Count + 1
If Val(board(indI, indJ - 1)) > 0 Then Count = Count + 1

'show numbers of the neighboring mines
cmdCell.BackColor = vbGreen
cmdCell.Caption = Str(Count)
cmdCell.ForeColor = vbBlack

'fill txtCheck for generating the event onChange
UserForm1.txtCheck.Text = Str(Count)
    
End Sub
