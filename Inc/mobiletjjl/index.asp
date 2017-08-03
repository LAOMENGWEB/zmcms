<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../sub.asp"-->
<!DOCTYPE HTML>
<html>
<head>
<meta content='text/html; charset=utf-8' http-equiv='Content-Type' />
<meta charset='utf-8' />
<title>移动端简历提交</title>
<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" name="viewport" />
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="black" name="apple-mobile-web-app-status-bar-style" />
<meta content="telephone=no" name="format-detection" />
<script src="<%=sitepath%>wap/js/layer.m.js"></script>
</head>

<body STYLE=" background:#dddddd">
<%
if DateDiff("s",session("time"),now())<5 then  
%>

<script language=javascript>

layer.open({
	
    btn: ['<%=word(624)%>'],
    content:'<%=word(370)%>',
	shade:false,
	yes: function(index){
       window.history.back(-1);
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
       window.history.back(-1);
	   layer.close(index)
    }
})


</script>



<%
response.end
end if
action=trim(request.QueryString("action"))	
if action="add" then 
Quarters=trim(request.Form("Quarters"))   
mMemID=request.QueryString("MemberID") 
Uname=trim(request.Form("Name"))
Sex=trim(request.Form("Sex"))
Marry=trim(request.Form("Marry"))
Birthday=trim(request.Form("Birthday"))
Stature=trim(request.Form("Stature"))
School=trim(request.Form("School"))
Studydegree=trim(request.Form("Studydegree"))
Specialty=trim(request.Form("Specialty"))
Gradyear=trim(request.Form("Gradyear"))	
Residence=trim(request.Form("Residence"))
Edulevel=trim(request.Form("Edulevel"))
Edulevel=trim(ChangeChr(Edulevel))
Experience=trim(request.Form("Experience"))
Experience=trim(ChangeChr(Experience))
Phone=trim(request.Form("Phone"))
Email=trim(request.Form("Email"))
Add=trim(request.Form("Add"))
Postcode=trim(request.Form("Postcode"))
if School="" then
School=null
end if
if Studydegree="" then
Studydegree=null
end if
if Specialty="" then
Specialty=null
end if
if Gradyear="" then
Gradyear=null
end if
if Add="" then
Add=null
end if
if Postcode="" then
Postcode=null
end if
if maillock=1 and newyp=1 then 

  SendStat   =   Jmail(""&PostMail&"",""&word(417)&""&Uname&""&word(418)&""&Quarters&""&word(419)&"",""&word(420)&""&Uname&""&word(421)&"<br>"&word(422)&""&Quarters&"<br>"&word(423)&""&Uname&"<br>"&word(424)&""&Sex&"<br>"&word(425)&""&Marry&"<br>"&word(426)&""&Birthday&"<br>"&word(427)&""&Stature&"<br>"&word(428)&""&School&"<br>"&word(429)&""&Studydegree&"<br>"&word(430)&""&Specialty&"<br>"&word(431)&""&Gradyear&"<br>"&word(432)&""&Residence&"<br>"&word(433)&""&Edulevel&"<br>"&word(434)&""&Experience&"<br>"&word(435)&""&Phone&"<br>"&word(436)&""&Email&"<br>"&word(437)&""&Add&"<br>"&word(438)&""&Postcode&"<br>"&word(439)&""&mc&""&word(440)&"","GB2312","text/html")  
end if



set rs=server.createobject("adodb.recordset")
sql="select * from HrDemandAccept where (id is null)" 
rs.open sql,conn,1,3
rs.addnew
rs("Quarters")=Quarters
rs("name")=Uname	 
rs("Sex")=Sex	 
rs("Marry")=Marry
rs("Birthday")=Birthday
rs("Stature")=Stature	
rs("School")=School
rs("Studydegree")=Studydegree
rs("Specialty")=Specialty
rs("Gradyear")=Gradyear   
rs("Residence")=Residence	  
rs("Edulevel")=Edulevel
rs("Experience")=Experience
rs("Phone")=Phone
rs("Email")=Email
rs("Add")=Add
rs("Postcode")=Postcode
rs("Adddate")=now()
rs("MemID")=mMemID
rs("lang")=language
rs.update
rs.close
set rs=nothing



if maillock=1 and newyp=1 then
%>

<script language=javascript>

layer.open({
	
    btn: ['<%=word(624)%>'],
    content:'<%=word(459)%>',
	shade:false,
	yes: function(index){
     location.href="<%=sitepath%>wap/job/?m=<%=request.QueryString("m")%>&language=<%=language%>";
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
    content:'<%=word(460)%>',
	shade:false,
	yes: function(index){
     location.href="<%=sitepath%>wap/job/?m=<%=request.QueryString("m")%>&language=<%=language%>";
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