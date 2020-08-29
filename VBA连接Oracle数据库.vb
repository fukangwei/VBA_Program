Sub linkOracle()
    Dim strConn As String
    Dim dbConn As Object
    Dim resSet As Object
    Dim db_sid, db_user, db_pass As String

    ' 在这里设置连接参数
    db_sid = "orcl"
    db_user = "SYS"
    db_pass = "Fuxinzi171720"
    '---------------------------
    Set dbConn = CreateObject("ADODB.Connection")
    Set resSet = CreateObject("ADODB.Recordset")
    strConn = "Provider=OraOLEDB.Oracle.1; user id=" & db_user & "; password=" & db_pass & "; data source = " & db_sid & "; Persist Security Info=True"
    '-------------------------------------------
    dbConn.Open strConn ' 打开数据库
    '-------------------------------------------
    Dim InsertCommand As String

    For i = 2 To 100
        CLASSID = Cells(i, "A")
        CLASSNAME = Cells(i, "B")

        ' 合成SQL语句
        InsertCommand = "insert into CLASSINFO (CLASSID, CLASSNAME) values (" & _
                        CLASSID & _
                        "," & _
                        "'" & CLASSNAME & "'" & _
                        ")"
        Set resSet = dbConn.Execute(InsertCommand) ' 使用SQL语句插入到数据库中
    Next
    '-------------------------------------------
    dbConn.Close ' 关闭数据库的连接
    '-------------------------------------------
    
    MsgBox "数据插入完成"
End Sub