<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="zmwbk.asp"-->
<%
If request.Form("submit") = "修改" Then
wwsm=Trim(request.form("wwsm"))
ww=Trim(request.form("ww"))
px=Trim(request.form("px"))
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ww where id = "&request.QueryString("id")&"" 
rs.Open sql, conn, 1, 3
rs("wwsm") =wwsm
rs("ww") = ww
rs("px") = px


rs.update
rs.close
set rs=nothing
call addlog("修改旺旺客服")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""ww.asp""</script>"
end if
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From ww where id="&id&" and lang='"&lang&"'", conn,1,3
	If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""ww.asp""</script>"
    response.End
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
	<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>
 <form   method="post" name="myform" >


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">旺旺客服修改 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0"  cellpadding="0" cellspacing="0" width="99%" class="dtable3" align="center">
 <tr>
 <td ><div class="dmeun"><a href="kefu.asp" target="main">返回菜单</a></div>

<div class="dmeun"><a href="ww.asp" target="main">旺旺客服管理</a></div>
 </td>
 

 </tr>



 
</table> 
   <table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
                  
<tr  > 
                    <td width="11%"  > 旺旺客服说明: </td>
                    <td width="89%">
                      <input name="wwsm" type="text" class="inputtext" id="wwsm" value="<%=rs("wwsm")%>" size="25" maxlength="30">
      </td>
     </tr>
				  <tr  > <td>旺旺号码:</td>
                    <td  >  <input name="ww"type="text" class="inputtext" id="ww" value="<%=rs("ww")%>"   size="20" maxlength="30"></td>
                  </tr>
				  
                  <tr  >
		     <td >排序：</td><td>
		      
		       <input name="px" type="text" id="px" value="<%=rs("px")%>" style="width:30px"/>
		      </td>
	      </tr>
                  
                  
                  
                  
                  <tr  > 
                    <td  colspan="2"> 
                        <input class="bnt"name="submit" type="submit" value="修改">
                        &nbsp;</td>
                  </tr>
     </table>
            
          </form>
       
</body>
</html>
