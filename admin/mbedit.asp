<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="admin.asp"-->
<!--#include file="conn.asp"-->

<%
Dim FileHTML, FileS , FileText
FileS = request("file")

If FileS = "" then 
Response.write "对不起，文件名未输入"
Response.end
End if 

Dim strText 
strText =  ReadFromTextFile(""&files&"","utf-8")
%>

<%
html=request("html")
 if Request("submit")="保存" then
   Call WriteToTextFile(""&files&"",""&html&"","utf-8")
   call addlog("修改模板文件")
	Response.Write("<script>alert('更新成功!');;window.history.back();</script>")
	
	
end if %>
<script language="JavaScript" src="../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
	 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>

<form name="form1" method="post" action="">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">模板修改</div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="template.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
  <tr>
    <td>您正在编辑<%=request("file")%> 文件，请您做好备份！</td>
  </tr>
  <tr>
    <td><textarea name="html" cols="100" rows="40" id="html"><%=strText%></textarea>
	<input type="hidden" name="file" value="<%=fileS%>">
	</td>
  </tr>
  <tr>
    <td> <input type="submit" name="Submit" class="bnt" value="保存"></td>
  </tr>
</table>



</form>
</body>
</html>
<%
'------------------------------------------------- 
'函数名称:ReadTextFile 
'作用:利用AdoDb.Stream对象来读取UTF-8格式的文本文件 
'---------------------------------------------------- 
public function ReadFromTextFile (FileUrl,CharSet) 
dim str 
set stm=server.CreateObject("adodb.stream") 
stm.Type=2 '以本模式读取 
stm.mode=3 
stm.charset=CharSet 
stm.open 
stm.loadfromfile server.MapPath(FileUrl) 
str=stm.readtext 
stm.Close 
set stm=nothing 
ReadFromTextFile=str 
end function 
'------------------------------------------------- 
'函数名称:WriteToTextFile 
'作用:利用AdoDb.Stream对象来写入UTF-8格式的文本文件 
'---------------------------------------------------- 
public Sub WriteToTextFile (FileUrl,Str,CharSet) 
set stm=server.CreateObject("adodb.stream") 
stm.Type=2 '以本模式读取 
stm.mode=3 
stm.charset=CharSet 
stm.open 
stm.WriteText str 
stm.SaveToFile server.MapPath(FileUrl),2 
stm.flush 
stm.Close 
set stm=nothing 
end Sub 


%>