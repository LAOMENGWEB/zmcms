<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" --> 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>获取静态页下面的规格价格</title>
</head>

<body>
<%
proid=request.QueryString("proid")

Set rspro=Server.CreateObject("ADODB.RecordSet") 
sql="select * from product where id="&proid
rspro.Open sql,conn,1,1


If rspro("P_SpecificationName") = "" Or IsNull(rspro("P_SpecificationName")) Then 
End If 
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed" then
If InStr(rspro("P_SpecificationName"),",") = 0 Then 
q = 0
P_SpecificationValue = Replace(rspro("P_SpecificationValue"),"|",",")
If Right(P_SpecificationValue,1) = "|" Then P_SpecificationValue = cut_Right(P_SpecificationValue)
If Right(P_SpecificationValue,1) = "," Then P_SpecificationValue = cut_Right(P_SpecificationValue)
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &"：</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and id in("& P_SpecificationValue &") and Lang='"&Language&"'")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 
P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);""><img src="""&sitepath&"/"&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);"">"& rs(0) &"</a></div>"
End If 
rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& id &""" id=""hidv_0001_"& id &"_0""></dl>"
Else
For q = 0 To UBound(Split(rspro("P_SpecificationName"),","))
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &"：</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"' and id in("& Split(rspro("P_SpecificationValue"),"|")(q) &")")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 

P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);""><img src="""&sitepath&""&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);"">"& rs(0) &"</a></div>"
End If 

rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& proid &""" id=""hidv_0001_"& proid &"_"& q &"""></dl>"
Next 
End If 



else

set rsjg2=server.CreateObject("adodb.recordset")
sqljg2="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg2.open sqljg2,conn,1,1

set rsjg12=server.CreateObject("adodb.recordset")
sqljg12="select * from product_price where grade_id="&rsjg2("GroupLevel")&" and product_id="&proid&" and Lang='"&Language&"'"
rsjg12.open sqljg12,conn,1,1

If InStr(rspro("P_SpecificationName"),",") = 0 Then 
q = 0
P_SpecificationValue = Replace(rspro("P_SpecificationValue"),"|",",")
If Right(P_SpecificationValue,1) = "|" Then P_SpecificationValue = cut_Right(P_SpecificationValue)
If Right(P_SpecificationValue,1) = "," Then P_SpecificationValue = cut_Right(P_SpecificationValue)
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &"：</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and Lang='"&Language&"' and id in("& P_SpecificationValue &")")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 
P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price") &"','0001',0);""><img src="""&sitepath&"/"&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price")  &"','0001',0);"">"& rs(0) &"</a></div>"
End If 
rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& id &""" id=""hidv_0001_"& id &"_0""></dl>"
Else
For q = 0 To UBound(Split(rspro("P_SpecificationName"),","))
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &"：</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"' and id in("& Split(rspro("P_SpecificationValue"),"|")(q) &")")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 

P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price")  &"','0001',0);""><img src="""&sitepath&""&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price")  &"','0001',0);"">"& rs(0) &"</a></div>"
End If 

rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& proid &""" id=""hidv_0001_"& proid &"_"& q &"""></dl>"
Next 
End If 
end if
response.write(rsprotemp)



















set rspro=nothing
rspro.close







%>

</body>
</html>
