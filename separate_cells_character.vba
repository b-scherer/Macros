Sub separate_with_character()    
	Const FirstRow = 2 ' First row with data    
	Dim LastRow As Long    
	Dim CurRow As Long    
	Dim Parts As Variant    
	Dim i As Long
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row    
	For CurRow = LastRow To FirstRow Step -1        
		Parts = VBA.split(Cells(CurRow, 2).Value, ";")        
		If UBound(Parts) > 0 Then            
		For i = UBound(Parts) To 1 Step -1                
			Cells(CurRow + 1, 1).EntireRow.Insert                
			Cells(CurRow + 1, 1).Value = Cells(CurRow, 1).Value                
			Cells(CurRow + 1, 2).Value = Parts(i)                
			Cells(CurRow + 1, 3).Value = Cells(CurRow, 3).Value            
		Next i            
		Cells(CurRow, 2).Value = Parts(0)        
		End If    
	Next CurRow
End Sub
