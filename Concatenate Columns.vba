Sub Concatenate_Column()    
    Const FirstRow = 1 ' First row with data    
    Dim LastRow As Long    
    Dim CurRow As Long    
    Dim Parts As Variant    
    Dim i As Long        
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row    
    For CurRow = LastRow To FirstRow Step -1        
        For i = CurRow - 1 To FirstRow Step -1            
            If Cells(CurRow, 2) = Cells(i, 2) Then                
            Cells(i, 1) = Cells(i, 1) & vbCrLf & Cells(CurRow, 1)                
            Cells(CurRow, 1) = ""                
            Cells(CurRow, 2) = ""            
            End If        
        Next    
    Next
End Sub
