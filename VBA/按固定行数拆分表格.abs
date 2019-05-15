Sub 拆分()

    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Application.SheetsInNewWorkbook = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(ThisWorkbook.Path & "\拆分表") Then
        Set f = fso.GetFolder(ThisWorkbook.Path & "\拆分表")
        f.Delete
    End If
    On Error Resume Next
    MkDir ThisWorkbook.Path & "\拆分表"
    
    j = 10000 '拆分行数
    file_name = 1
    
    For i = 1 To Worksheets("其他出库").UsedRange.Rows.Count Step j
        Set wb = Workbooks.Add
        ThisWorkbook.Worksheets("其他出库").Rows(i).Resize(j).Copy wb.Worksheets(1).[a1]
        wb.SaveAs Filename:=ThisWorkbook.Path & "\拆分表\" & file_name & ".csv", FileFormat:=xlCSV, CreateBackup:=False
        wb.Close False
        file_name = file_name + 1
    Next i
    MsgBox "所需数据已经导出到" & ThisWorkbook.Path & "\【拆分表】文件夹内"
    
    Application.ScreenUpdating = True
    
End Sub
