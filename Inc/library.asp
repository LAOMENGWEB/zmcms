<%
Const OBJ_RST = "ADODB.Recordset"
Const OBJ_CONN = "ADODB.Connection"
Const OBJ_STRM = "ADODB.Stream"
Const OBJ_FSO = "Scripting.FilesyStemObject"
Const OBJ_XHTP = "MSXML2.XMLHTTP"
Const OBJ_DOM = "MSXML2.DOMDocument"
'/////////基础操作函数部分

'过程：输出字符串[代替Response.Write]

Sub echo(Str)
    response.Write(Str)
End Sub

'过程：结束页面并输出字符串

Sub die(Str)
    response.Write(Str)
    response.End()
End Sub


'函数：将ASP文件运行结果返回为字串

Function ob_get_contents(Path)
    Dim tmp, a, b, t, matches, m
    Dim Str
    Str = file_iread(Path)
    tmp = "dim htm : htm = """""&vbCrLf
    a = 1
    b = InStr(a, Str, "<%") + 2
    While b > a + 1
        t = Mid(Str, a, b - a -2)
        t = Replace(t, vbCrLf, "{::vbcrlf}")
        t = Replace(t, vbCr, "{::vbcr}")
        t = Replace(t, """", """""")
        tmp = tmp & "htm = htm & """ & t & """" & vbCrLf
        a = InStr(b, Str, "%\>") + 2
        tmp = tmp & str_replace("^\s*=", Mid(Str, b, a - b -2), "htm = htm & ") & vbCrLf
        b = InStr(a, Str, "<%") + 2
    Wend
    t = Mid(Str, a)
    t = Replace(t, vbCrLf, "{::vbcrlf}")
    t = Replace(t, vbCr, "{::vbcr}")
    t = Replace(t, """", """""")
    tmp = tmp & "htm = htm & """ & t & """" & vbCrLf
    tmp = Replace(tmp, "response.write", "htm = htm & ", 1, -1, 1)
    tmp = Replace(tmp, "echo", "htm = htm & ", 1, -1, 1)
    'execute(tmp)
    executeglobal(tmp)
    htm = Replace(htm, "{::vbcrlf}", vbCrLf)
    htm = Replace(htm, "{::vbcr}", vbCr)
    ob_get_contents = htm
End Function








'/////////字符串操作函数部分

'函数：正则验证

Function str_test(Pattern, Str)
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = True
        .Global = True
        .Pattern = Pattern
        tmp = .Test(Str)
    End With
    Set reg = Nothing
    str_test = tmp
End Function

'函数：正则替换[不区分大小写]

Function str_replace(Pattern, byval Str, s)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = True
        .Global = True
        .Pattern = Pattern
        tmp = .Replace(Str, s)
    End With
    Set reg = Nothing
    str_replace = tmp
End Function

'函数：正则替换[区分大小写]

Function str_ireplace(Pattern, byval Str, s)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = False
        .Global = True
        .Pattern = Pattern
        tmp = .Replace(Str, s)
    End With
    Set reg = Nothing
    str_ireplace = tmp
End Function



'函数：执行正则搜索并返回结果集[不区分大小写]

Function str_execute(Pattern, byval Str)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = True
        .Global = True
        .Pattern = Pattern
        Set tmp = .Execute(Str)
    End With
    Set reg = Nothing
    Set str_execute = tmp
End Function












'函数：检测文件/文件夹是否存在

Function file_exists(Path)
    Dim tmp
    tmp = False
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    If fso.FileExists(server.MapPath(Path)) Then tmp = True
    If fso.FolderExists(server.MapPath(Path)) Then tmp = True
    Set fso = Nothing
    file_exists = tmp
End Function







Function file_read(Path)
    Dim tmp
    tmp = ""
    If Left(Path, 7) = "http://" Then '读取远程文件
        Dim xmlhttp
        Set xmlhttp = server.CreateObject(OBJ_XHTP)
        With xmlhttp
            .Open "get", Path, False
            .send()
            tmp = .responsetext
        End With
        Set xmlhttp = Nothing
    Else '读取本地文件
        If Not file_exists(Path) Then
            file_read = tmp
            Exit Function
        End If
        Dim stream
        Set stream = server.CreateObject(OBJ_STRM)
        With stream
            .Type = 2 '文本类型
            .mode = 3 '读写模式
            .charset = "utf-8"
            .Open
            .loadfromfile(server.MapPath(Path))
            tmp = .readtext()
        End With
        stream.Close
        Set stream = Nothing
    End If
    file_read = tmp
End Function



'函数:读取ASP类型文件的全部内容

Function file_iread(Path)
    Dim Str
    Str = file_read(Path)
    Dim Pattern
    Pattern = "<\!--#include[ ]+?file[ ]*?=[ ]*?""(\S+?)""--\>"
    Dim matches
    Set matches = str_execute(Pattern, Str)
    Dim m, f, tmp
    For Each m in matches
        f = Mid(Path, 1, instrrev(Path, "/"))&m.submatches(0)
        tmp = file_read(f)
        If str_test(Pattern, tmp) Then tmp = file_iread(f) '处理子包含
        Str = Replace(Str, m.Value, tmp)
    Next
    Pattern = "<%@[ ]*?LANGUAGE[ ]*?=[ ]*?""[a-zA-Z]+?""[ ]+?CODEPAGE[ ]*?=[ ]*?""[0-9]+?""[ ]*?%\>"
    Str = str_replace(Pattern, Str, "")
    file_iread = Str
End Function



'函数：指定字串数组的元素是否含有指定字串





'函数：执行SQL语句-只读方式

Function ado_query(byval sql)
    Set ado_query = ado_iquery(sql, conn, 3, 1)
End Function

'函数：执行SQL语句-可写方式

Function ado_mquery(byval sql)
    Set ado_mquery = ado_iquery(sql, conn, 3, 3)
End Function

'函数：执行SQL语句

Function ado_iquery(byval sql, conn, cursortype, locktype)
    If Trim(sql) = "" Then Exit Function
    If Trim(n) = "" Or Not IsNumeric(n) Then n = 1
    Dim rs
    If LCase(Left(LTrim(sql), 6)) = "select" Then
        Set rs = server.CreateObject(OBJ_RST)
        rs.CursorLocation = 3
        rs.Open sql, conn, cursortype, locktype
    Else
        Set rs = conn.Execute(sql)
    End If
    Set ado_iquery = rs
End Function



%>
