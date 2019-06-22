Sub SQL()

    Dim cnn As Object, SQL$
    Set cnn = CreateObject("ADODB.Connection")
    cnn.Open "provider=Microsoft.ACE.OLEDB.12.0;extended properties=excel 12.0;data source=" & ThisWorkbook.Path & "\sql.xlsm"
    SQL = "SELECT a.学号, b.语文 from [学号$]a left join [成绩$]b on a.学号=b.学号"
    Sheets("VBA").Select
    Range("a1") = "学号"
    Range("b1") = "语文"
    Range("a2").CopyFromRecordset cnn.Execute(SQL)
    cnn.Close
    Set cnn = Nothing

End Sub
