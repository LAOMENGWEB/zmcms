<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Server.ScriptTimeOut=1000000%>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<!--#include file="htmlconfig.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="admin_html_function.asp"-->


<%
if Instr(session("AdminPurview"),"|114,")=0 then
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
if yy = 1 then
  response.Write "<script language='javascript'>alert('请先在【系统参数配置】中将静态HTML设置为开启！');history.go(-1);</script>"
  response.End
end If
%>
<br />
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="300"><table width="90%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td style="border-bottom: #ccc 1px solid; border-top: #ccc 1px solid; border-left: #ccc 1px solid; border-right: #ccc 1px solid"><img src="Images/Survey_1.gif" width="0" height="16" id="bar_img" name="bar_img" align="absmiddle"></td>
        </tr>
      </table></td>
    <td><span id="bar_txt2" name="bar_txt2" style="font-size:12px; color:red;"></span><span id="bar_txt1" name="bar_txt1" style="font-size:12px">0</span><span style="font-size:12px">%</span></td>
  </tr>
</table>
<%
Call Htmlabout
call htmlnewsfl
call htmlprofl
call htmlalfl
call htmltvfl
call htmlxzfl
call htmlzpfl
call htmlnews
call htmlpro
call htmlal
call htmltv
call htmlxz
call htmlzp

call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
Response.Write "<script>bar_img.width="&Fix((1/7)*300)&";bar_txt1.innerHTML=""成功生成静态首页。完成比例" & formatnumber(1/7*10) & """;</script>"
Response.Flush
call htmll("/"&iinfo&"/","",""&Newclass&".html",""&iinfo&"/index.asp","m=",minfo,"language=",lang,"","","","","","","","")

Response.Write "<script>bar_img.width="&Fix((2/7)*300)&";bar_txt1.innerHTML=""成功生成“新闻列表”静态页面。完成比例：" & formatnumber(2/7*20) & """;</script>"
Response.Flush
call htmll("/"&product&"/","",""&Productclass&".html",""&product&"/index.asp","m=",mpro,"language=",lang,"","","","","","","","")

Response.Write "<script>bar_img.width="&Fix((3/7)*300)&";bar_txt1.innerHTML=""成功生成“产品列表”静态页面。完成比例：" & formatnumber(3/7*30) & """;</script>"
Response.Flush
call htmll("/"&photo&"/","",""&alclass&".html",""&photo&"/index.asp","m=",mph,"language=",lang,"","","","","","","","")

Response.Write "<script>bar_img.width="&Fix((4/7)*300)&";bar_txt1.innerHTML=""成功生成“案例列表”静态页面。完成比例：" & formatnumber(4/7*40) & """;</script>"
Response.Flush

call htmll("/"&job&"/","",""&zpclass&".html",""&job&"/index.asp","m=",mzp,"language=",lang,"","","","","","","","")
Response.Write "<script>bar_img.width="&Fix((5/7)*300)&";bar_txt1.innerHTML=""成功生成“人才列表”静态页面。完成比例：" & formatnumber(5/7*50) & """;</script>"
Response.Flush
call htmll("/"&down&"/","",""&Downclass&".html",""&down&"/index.asp","m=",mdo,"language=",lang,"","","","","","","","")

Response.Write "<script>bar_img.width="&Fix((6/7)*300)&";bar_txt1.innerHTML=""成功生成“下载列表”静态页面。完成比例：" & formatnumber(6/7*60) & """;</script>"
Response.Flush
call htmll("/"&video&"/","",""&tvclass&".html",""&video&"/index.asp","m=",mtv,"language=",lang,"","","","","","","","")

Response.Write "<script>bar_img.width="&Fix((7/7)*300)&";bar_txt1.innerHTML=""成功生成“视频列表”静态页面。完成比例：" & formatnumber(7/7*70) & """;</script>"
Response.Flush

Response.Write "<script>bar_img.width=300;bar_txt2.innerHTML='';bar_txt1.innerHTML=""成功生成相关静态文件。完成比例：100"";</script>"
Response.End

%>