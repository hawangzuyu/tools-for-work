Sub 比较文件()

    Application.ScreenUpdating = False
    
    ThisWorkbook.Worksheets(1).UsedRange.Clear
    
    MyPath = ThisWorkbook.Path
    File_INITEM = MyPath & "\M_INITEM.xlsx"
    File_OTHER_INITEM = MyPath & "\M_OTHER_INITEM.xlsx"
    Set wb1 = Workbooks.Open(File_INITEM)
    wb1.Worksheets(1).Range("a1").EntireColumn.Copy ThisWorkbook.Worksheets(1).Range("a1")
    wb1.Worksheets(1).Range("l1").EntireColumn.Copy ThisWorkbook.Worksheets(1).Range("b1")
    wb1.Close
    Set wb2 = Workbooks.Open(File_OTHER_INITEM)
    wb2.Worksheets(1).Range("a1").EntireColumn.Copy ThisWorkbook.Worksheets(1).Range("d1")
    wb2.Worksheets(1).Range("i1").EntireColumn.Copy ThisWorkbook.Worksheets(1).Range("e1")
    wb2.Close
    
    row1 = ThisWorkbook.Worksheets(1).Range("a1").End(xlDown).Row
    row2 = ThisWorkbook.Worksheets(1).Range("d1").End(xlDown).Row
    row0 = Application.WorksheetFunction.Max(row1, row2)
    
    With ThisWorkbook.Worksheets(1)
    For i = 2 To row0

        If .Cells(i, 1) <> .Cells(i, 4) Or .Cells(i, 2) <> .Cells(i, 5) Then
            .Cells(i, 7) = "不一致"
        End If
    Next i
    
    For j = 2 To row0
    
        If .Cells(j, 7) = "不一致" Then
            MsgBox "不一致"
            Exit Sub
        End If
    Next j
    
    End With
    
    MsgBox "一致"
    
    Application.ScreenUpdating = True
    
End Sub
