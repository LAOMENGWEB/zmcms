<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="htmlconfig.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
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

totalrec=Conn.Execute("select count(*) from job where lang='"&lang&"'")(0)
sql="Select * from job where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
content=rs("HrDetail")
zdfy=rs("zdfy")
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&zpname&""&Separated&""&ID&""&Separated&"1.html",""&showjob&"/index.asp","jobid=",ID,"m=",mzp,"language=",lang,"page=",1,"","","","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&zpname&""&Separated&""&ID&""&Separated&""&i&".html",""&showjob&"/index.asp","jobid=",ID,"m=",mzp,"language=",lang,"page=",i,"","","","")
next
end if
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
call addlog("生成所有"&zhaopin&"")
conn.close
set conn=Nothing
%>
</body>
</html>
