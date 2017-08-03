
<!--#include file="../inc/sub.asp" -->

<%
dim rs_Hits,m_SQL
dim m_ID
m_ID = ReplaceBadChar(Request.QueryString("id"))
m_LX = ReplaceBadChar(Request.QueryString("LX"))
tj = ReplaceBadChar(Request.QueryString("tj"))
If tj = "yes" Then
conn.Execute("update "&m_LX&" set hits = hits + 1 where ID=" & m_ID & "")
Else
m_SQL = "select hits from "&m_LX&" where ID=" & m_ID
Set rs_Hits = conn.Execute(m_SQL)
response.write "document.write("&rs_Hits(0)&");"	
rs_Hits.Close
set rs_Hits=nothing
End If

Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function
%>
