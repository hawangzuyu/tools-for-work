Sub 合并目录所有工作簿全部工作表()

Dim MP, MN, AW, Wbn, wn

Dim Wb As Workbook

Dim i, a, b, d, c, e

Application.ScreenUpdating = False

MP = ActiveWorkbook.Path

MN = Dir(MP & "\" & "*.xls")

AW = ActiveWorkbook.Name

Num = 0

e = 1

Do While MN <> ""

If MN <> AW Then

Set Wb = Workbooks.Open(MP & "\" & MN)

a = a + 1

With Workbooks(1).ActiveSheet

For i = 1 To Sheets.Count

If Sheets(i).Range("a1") <> "" Then

Wb.Sheets(i).Range("a1").Resize(1, Sheets(i).UsedRange.Columns.Count).Copy .Cells(1, 1)

d = Wb.Sheets(i).UsedRange.Columns.Count

c = Wb.Sheets(i).UsedRange.Rows.Count - 1

wn = Wb.Sheets(i).Name

.Cells(1, d + 1) = "表名"

.Cells(e + 1, d + 1).Resize(c, 1) = MN & wn

e = e + c

Wb.Sheets(i).Range("a2").Resize(c,d).Copy .Cells(.Range("a1048576").End(xlUp).Row + 1, 1)

End If

Next

Wbn = Wbn & Chr(13) & Wb.Name

Wb.Close False

End With

End If

MN = Dir

Loop

Range("a1").Select

Application.ScreenUpdating = True

MsgBox "共合并了" & a & "个工作薄下全部工作表。如下：" & Chr(13) & Wbn, vbInformation, "提示"

End Sub
