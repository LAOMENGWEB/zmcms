<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/function.asp" -->

<%
if Instr(session("AdminPurview"),"|513,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>





<%If request.Form("submit") = "修改" Then%>
<%
id=request("id")
shipmentName=Trim(request.form("shipmentName")) 
Content=Trim(request.form("Content"))
shipmentMoney=Trim(request.form("shipmentMoney"))

Set rs = server.CreateObject("adodb.recordset")
sql="select * from shoppingshipment where id="&id
rs.open sql,conn,1,3

rs("shipmentMoney")=shipmentMoney
rs("Content")=Content
rs("shipmentName")=shipmentName
rs("adddate")=now()
rs("lang")=lang
rs.update
rs.close
call addlog("修改配送方式")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""psfs.asp""</script>"
end if
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From shoppingshipment where id="&id&" and lang='"&lang&"'", conn,1,1
	If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""psfs.asp""</script>"
    response.End

end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>


<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->

</head>

<body>
 <form   method="post" name="myform" class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 配送方式修改  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="psfs.asp" target="main">配送方式管理</a></div></td>
 

 </tr>
</table>

              <table width="99%"  border="0" cellpadding="6" cellspacing="0"  align="center" class="dtable4">
                  
				  <tr  > 
                    <td  > 配送方式名称:</td>
					<td>
                      <input name="shipmentName"type="text" class="inputtext" id="shipmentMoney" value="<%=rs("shipmentName")%>"   size="30" maxlength="30" datatype="*" nullmsg="名称不能为空"/></td>
                  </tr><tr  > 
                    <td  > 配送金额: </td>
					<td>
                      <input name="shipmentMoney" type="text" class="inputtext" id="shipmentName" value="<%=rs("shipmentMoney")%>" size="8" maxlength="30" datatype="*" nullmsg="金额不能为空"/>
                                      </td>
                    </tr>
				  

				    <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="Content"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
   afterBlur: function(){this.sync();}
 });
 	K('input[name=appendHtml]').click(function(e) {
					editor.insertHtml('[zmcmsfy]');
				});
			
			prettyPrint();
 });
 </script>
				  <tr  > 
                    <td  > 详细说明:</td>
					<td>  <textarea name="Content" id="Content" style="width:700px;height:300px;visibility:hidden;" datatype="*" nullmsg="说明不能为空"><%=rs("content")%></textarea></td>
                  </tr>
                    

                  
                  <tr  > 
                    <td  colspan="2"> 
                        <input class="bnt" name="submit" type="submit" value="修改">
&nbsp;                                      </td>
                  </tr>
            </table>
            
</form>
       

</body>
</html>
<%
rs.close
set rs=nothing
%>