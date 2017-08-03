<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<%

	''通过xmlhtml方式获取远程文件内容
	public function gethttp(byval t0,byval t1)
		select case t1
			case "get","post"
			case else:t1="get"
		end select
		on error resume next
		dim http
		set http=server.createobject("microsoft.xmlhttp")
			http.open t1,t0,false
			http.setRequestHeader "If-Modified-Since","0"
			http.send()
			gethttp=http.responsetext
		set http=nothing
	end function

%>
<%
if Instr(session("AdminPurview"),"|03,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='adminlog.asp';</script>" 
response.end()
else
conn.Execute ( "delete from adminlog where id in ("&Request("id")&")" )
call addlog("删除管理员日志")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href='adminlog.asp'</script>"
end if
end if
%>
<%
If request.Form("button") = "删除3天" Then
if DateDiff("d",""&rs("logintime")&"",Date())<=0 then

conn.Execute ( "delete from adminlog " )
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""adminlog.asp""</script>"
end if
end if
%>
<%
If request.Form("button") = "删除所有" Then

conn.Execute ( "delete from adminlog " )
call addlog("删除管理员日志")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""adminlog.asp""</script>"

end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
 <link href="css/style.css" rel="stylesheet" type="text/css">
	 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
		<SCRIPT language=javascript>
function CheckAll(){ 
 for (var i=0;i<eval(zmqy.elements.length);i++){ 
  var e=zmqy.elements[i]; 
  if (e.name!="allbox") e.checked=zmqy.allbox.checked; 
 } 
} 
function ConfirmDel()
{
   if(confirm("确定要删除选中的登录日志吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
</head>

<body>
<form name="zmqy" method="Post">
<%
DB_Sql = "select * from adminlog "
DB_Sql = DB_Sql & "order by ID desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn,1,1
If Not(rs.Bof and rs.Eof) Then


'================================================
'调用分页类开始
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 20 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum

Obj_Page.Setpage_Pathurl = "adminlog.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">管理员操作日志</div></td>
</tr>
</table>



              <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2" >
                
                <tr  class="wz">
				<td width="4%" > 
                  <div align="center">全选</div></td>
                  <td width="4%" ><div align="center">编号</div></td> 
                  <td width="11%" > 
                  <div align="center">登录名</div></td>
                  <td width="11%" ><div align="center">用户名</div></td>
                  <td width="12%" ><div align="center">操作内容</div></td>
                  <td width="13%" ><div align="center">登录IP</div></td>
                  <td width="35%" ><div align="center">登录时浏览器</div></td>
				  <td width="10%" style="border-right:none;"><div align="center">操作时间</div></td>
                </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
               
                <tr  >
				 <td><div align="center">
                   <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/>
                 </div></td>
                  <td  ><div align="center"><%=rs("id")%></div></td> 
                  <td   > 
                    <div align="center"><%=rs("UserName")%></div></td>
                  <td ><div align="center"><%=rs("glyName")%></div></td>
                  <td ><div align="center"><%=rs("logcontent")%></div></td>
                  
                  <td  > 
				    <div align="center">
					<%=gethttp("http://whois.pconline.com.cn/ip.jsp?ip="&rs("loginip")&"","")%>
			    	          </div></td>
                  <td ><div align="center"><%=rs("loginsoft")%></div></td>
				  <td style="border-right:none;"><div align="center"><%=rs("logintime")%></div></td>
				  
                </tr>
 <%
rs.MoveNext
Next	
%>    
 <tr  >
                  <td > 
                    
                      
                      <div align="center">
                        <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
      </div></td>
					  <td  colspan="7" style="padding:5px; text-align:left" > 
                      <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的登录日志吗，将不可恢复?');"/> 
					  <input class="bnt" type="submit" name="button" id="button" value="删除所有"   onClick="return confirm('确定要删除全部的登录日志吗，将不可恢复?');"/>
					  <!--<input class="bnt" type="submit" name="button" id="button" value="删除3天"   onClick="return confirm('确定要删除3天以以前的登录日志吗，将不可恢复?');"/>-->
					  </td>
                </tr> 
  </table>
		<%If Obj_Page.OutErrInfo <> "" Then
		Response.Write Obj_Page.OutErrInfo '报错信息显示
		Response.End
		else
	Call Obj_Page.Page_Navigation() '分页导航
	end if
    else
Response.Write "目前暂无任何信息"
Response.end
	%>
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</form>
</body>
</html>
