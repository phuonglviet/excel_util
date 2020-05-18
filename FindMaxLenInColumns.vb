
' Find max len in column
Sub FindMaxLenInColumns()

    Dim rActiveSheet As Worksheet
    'Dim c As range
    Dim Msg As String
    Dim sRange As range
    Dim wRange As range
    Dim maxLength As Long
    Dim sAddress As String
    Dim lenOfCell As Long
    Dim resltRow As Integer
    Dim resltCol As Integer
    Dim countCol As Integer
    Dim countRow As Integer
    
    ' Setting
    resltRow = 1 'write result at row number
    Set rActiveSheet = ActiveSheet
    Set sRange = rActiveSheet.range("B4:AA24") 'Range to find
    
    'Find ma len each column in range
    For Each rng In sRange.Columns
        maxLength = 0
        For Each ce In rng.Cells
        
            lenOfCell = Len(ce.Value)
            
            If (lenOfCell > maxLength) Then
                maxLength = lenOfCell
                resltCol = ce.Column
                'sAddress = Cell.Address
            End If
            'Set lenOfCell = Len(Cell.Value)
        
        Next ce
        
        ' write result
        ' Cells(resltRow, ce.Column).Value = maxLength
        Cells(resltRow, resltCol).Value = maxLength
        countCol = countCol + 1
        countRow = countRow + 1
    Next rng
    
    MsgBox "Done! Columns:" & countCol & " , Rows:" & countRow & " has done."
   
End Sub

