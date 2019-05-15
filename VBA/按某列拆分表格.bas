Sub 拆分表格_备份()

    c = 3 '拆分列号
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    arr = [a1].CurrentRegion
    lc = UBound(arr, 2)
    Set Rng = [a1].Resize(, lc)
    Set d = CreateObject("scripting.dictionary")
    For i = 2 To UBound(arr)
        If Not d.exists(arr(i, c)) Then
            Set d(arr(i, c)) = Cells(i, 1).Resize(1, lc)
        Else
            Set d(arr(i, c)) = Union(d(arr(i, c)), Cells(i, 1).Resize(1, lc))
        End If
    Next
    k = d.Keys
    t = d.Items
    For i = 0 To d.Count - 1
        With Workbooks.Add(xlWBATWorksheet)
            Rng.Copy .Sheets(1).[a1]
            t(i).Copy .Sheets(1).[a1]
            .SaveAs Filename:=ThisWorkbook.Path & "\" & k(i) & ".csv", FileFormat:=xlCSV, CreateBackup:=False
            .Close
        End With
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "完毕"

End Sub
