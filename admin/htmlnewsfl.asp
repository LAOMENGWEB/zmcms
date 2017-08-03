<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="htmlconfig.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
if Instr(session("AdminPurview"),"|114,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
if yy = 1 then
  response.Write "<script language='javascript'>alert('请先在【系统参数配置】中将静态HTML设置为开启！');history.go(-1);</script>"
  response.End
end If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
</head>

<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="300"><table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td style="border-bottom: #ccc 1px solid; border-top: #ccc 1px solid; border-left: #ccc 1px solid; border-right: #ccc 1px solid"><img src="Images/Survey_1.gif" width="0" height="16" id="bar_img" name="bar_img" align="absmiddle"></td>
        </tr>
      </table></td>
    <td><span id="bar_txt2" name="bar_txt2" style="font-size:12px; color:red;"></span><span id="bar_txt1" name="bar_txt1" style="font-size:12px">0</span><span style="font-size:12px">%</span></td>
  </tr>
</table>
<%
call htmll("/"&iinfo&"/","",""&Newclass&".html",""&iinfo&"/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
totalrec=Conn.Execute("Select count(*) from news where lang='"&lang&"'") (0) '获取新闻的总数   
totalpage=int(totalrec/NewInfo)  '计算分页数
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&iinfo&"/","",""&NewClass&""&Separated&"1.html",""&iinfo&"/index.asp","m=",minfo,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&iinfo&"/","",""&NewClass&""&Separated&""&i&".html",""&iinfo&"/index.asp","m=",minfo,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",i)
next
end If


'新闻一级分类生成开始
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from newsclass where lang='"&lang&"' order by px desc"
rs.open sql,conn,1,1
If rs.eof Then 
Class_Num=0
Else 
Class_Num=1
do while not rs.eof
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where ParentClassID = 0 and lang='"&lang&"' "
rs1.open sql1,conn,1,1
do while not rs1.eof
pone=rs1("ID")
m=rs1("m") ' 获取一级分类的高亮值
totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If

creatwjj("../"&rs1("mulu")&"/") '建立新闻分类存放的文件夹

if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&NewClass&""&Separated&""&pone&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&NewClass&""&Separated&""&pone&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",i)
next
End If

Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from newsclass where ParentClassID ="&pone&" and lang='"&lang&"'"
rs2.open sql2,conn,1,1
do while not rs2.eof
ptwo=rs2("ID")
m1=rs2("ID")  ' 获取二级分类的高亮值
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and zmtwo="&ptwo&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&NewClass&""&Separated&""&ptwo&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&NewClass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",i,"","")
next
End If

Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from newsclass where ParentClassID ="&ptwo&" and lang='"&lang&"'"
rs3.open sql3,conn,1,1
do while not rs3.eof
pthree=rs3("ID")
m2=rs3("ID") '获取三级分类的高亮值
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&NewClass&""&Separated&""&pthree&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&NewClass&""&Separated&""&pthree&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",i)
next
End If
rs3.movenext
Loop
rs3.close
set rs3=nothing
rs2.movenext
Loop
rs2.close
set rs2=nothing
rs1.movenext
Loop
rs1.close
set rs1=nothing
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
call addlog("生成"&xinwen&"分类")
conn.close
set conn=nothing
%>
</body>
</html>
