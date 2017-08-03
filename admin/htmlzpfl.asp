<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="htmlconfig.asp"-->
<!--#include file="conn.asp"-->
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
call htmll("/"&job&"/","",""&zpclass&".html",""&job&"/index.asp","m=",mzp,"language=",lang,"","","","","","","","")
totalrec=Conn.Execute("Select count(*) from job where lang='"&lang&"'")(0)
totalpage=int(totalrec/zpInfo)
If (totalpage * zpInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&job&"/","",""&zpclass&""&Separated&"1.html",""&job&"/index.asp","m=",mzp,"language=",lang,"Page=",1,"","","","","","")
else
for i=1 to totalpage
call htmll("/"&job&"/","",""&zpclass&""&Separated&""&i&".html",""&job&"/index.asp","m=",mzp,"language=",lang,"Page=",i,"","","","","","")
Response.Write "<script>bar_img.width="&Fix((i/totalpage)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&i&"个分类的HTML静态页面。完成比例：" & formatnumber(i/totalpage*100) & """;</script>"
Response.Flush
next
end if
rs.close
set rs=Nothing
call addlog("生成所有应聘页面")
Response.Write "<script>bar_img.width=300;bar_txt2.innerHTML='';bar_txt1.innerHTML=""成功生成相关静态文件。完成比例：100"";</script>"
Response.end
conn.close
set conn=Nothing
%>
</body>
</html>
