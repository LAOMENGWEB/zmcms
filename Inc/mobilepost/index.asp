<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../sub.asp"-->
<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title></title>
<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" name="viewport" />
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="black" name="apple-mobile-web-app-status-bar-style" />
<meta content="telephone=no" name="format-detection" />
<script src="<%=sitepath%>wap/js/layer.m.js"></script>
</head>
<body  STYLE=" background:#dddddd">

<%
if DateDiff("s",session("time"),now())<5 then  
%>
<script language=javascript>
layer.open({
btn: ['<%=word(624)%>'],
content:'<%=word(370)%>',
shade:false,
yes: function(index){
location.href="<%=sitepath%>wap/guest/?m=<%=request.QueryString("m")%>&language=<%=language%>";
layer.close(index)
}
})
</script>
<%
response.end
else  
session("time")=now()
end if
if chkpost=0 then
%>
<script language=javascript>
layer.open({
btn: ['<%=word(624)%>'],
content:'<%=word(491)%>',
shade:false,
yes: function(index){
location.href="<%=sitepath%>wap/guest/?m=<%=request.QueryString("m")%>&language=<%=language%>";
layer.close(index)
}
})
</script>
<%
response.end
end if
mMemID=request.QueryString("MemberID")
set rs=server.createobject("adodb.recordset")
sql="select * from ly "
if trim(request.form("qqh"))="1" then
qqh=1
else
qqh=0
end if
rs.open sql,conn,1,3
title=Trim(request.form("title")) 
phone=Trim(request.form("phone"))
emaill=Trim(request.form("emaill"))
qq=Trim(request.form("qq"))
content=Trim(request.form("content"))

if not ChkSB(title) then
%>

<script language=javascript>

layer.open({
	
    btn: ['<%=word(624)%>'],
    content:'<%=word(400)%>',
	shade:false,
	yes: function(index){
       window.history.back(-1);
	   layer.close(index)
    }
})


</script>

<%
response.End()
End If
if not ChkSB(content) then
%>

<script language=javascript>

layer.open({
	
    btn: ['<%=word(624)%>'],
    content:'<%=word(401)%>',
	shade:false,
	yes: function(index){
       window.history.back(-1);
	   layer.close(index)
    }
})


</script>

<%
response.End()
End If


if maillock=1 and newguest=1 then 
  SendStat   =   Jmail(""&PostMail&"",""&word(406)&""&title&""&word(407)&"",""&word(408)&""&title&""&word(409)&"<br>"&word(410)&""&phone&"<br>"&word(411)&""&emaill&"<br>"&word(412)&""&qq&"<br>"&word(413)&""&content&"","GB2312","text/html")   
end if
rs.addnew
rs("title")=title
rs("emaill")=emaill
rs("phone")=phone
rs("sj")=now()
rs("content")=content
rs("qq")=qq
rs("lyip")=getip()
rs("lang")=language
rs("qqh")=qqh
rs("MemID")=mMemID
if ly=1 then
rs("sh")=0
end if
if ly=2 then
rs("sh")=1
end if
rs.update
rs.Close
Set rs = Nothing
if ly=1 then
%>

<script language=javascript>

layer.open({
	
    btn: ['<%=word(624)%>'],
    content:'<%=word(414)%>',
	shade:false,
	yes: function(index){
     location.href="<%=sitepath%>wap/guest/?m=<%=request.QueryString("m")%>&language=<%=language%>";
	   layer.close(index)
    }
})


</script>

<%
else
if maillock=1 and newguest=1 then
%>

<script language=javascript>

layer.open({
	
    btn: ['<%=word(624)%>'],
    content:'<%=word(415)%>',
	shade:false,
	yes: function(index){
     location.href="<%=sitepath%>wap/guest/?m=<%=request.QueryString("m")%>&language=<%=language%>";
	   layer.close(index)
    }
})


</script>

<%
else
%>

<script language=javascript>

layer.open({
	
    btn: ['<%=word(624)%>'],
    content:'<%=word(416)%>',
	shade:false,
	yes: function(index){
     location.href="<%=sitepath%>wap/guest/?m=<%=request.QueryString("m")%>&language=<%=language%>";
	   layer.close(index)
    }
})


</script>

<%
end if 
end if
%>


</body>
</html>