<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<% 
id=int(request("id"))
pid=request.QueryString("pid")
set rs=server.createobject("adodb.recordset")   
sql="select * from Definitons where id="&id   
rs.open sql,conn,1,3  
if request("action")<>"" then
id=int(request("id"))
Label=trim(request("Label"))
pid=trim(request("pid"))
px=trim(request("px"))
Labe2=trim(request("Labe2"))
type1=trim(request("Type"))
FieldName=trim(request("FieldName"))
xxz=request("xxz")
gaoji=request("gaoji")


rs("Type")=Type1
rs("tmaxlength")=tmaxlength
rs("mmaxlength")=mmaxlength
rs("Label")=Label
rs("Labe2")=Labe2
rs("gaoji")=gaoji
rs("px")=px
rs("FieldName")=FieldName
rs("Required")=request("Required")
rs("xxz")=xxz
rs.update

call addlog("修改表单参数")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""form.asp?xx="&pid&"""</script>"

end if
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from Definitons where id="&request.QueryString("id")&""
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <script language="JavaScript" src="js/ChinesePY.js"></script>
<script language="JavaScript" src="js/getpy.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
</head>

<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >修改表单参数</td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
 <div class="dmeun"><a href="diyformmanage.asp" target="main">自定义表单管理</a></div>
 
 </td>
 

 </tr>
</table>
<form method="POST" action="?id=<%=id%>&action=save&pid=<%=pid%>" class="checkform">
<table width=99% border="0" cellpadding="0" cellspacing="0"   align="center" class="dtable4" >

  <tr>
  <td>文本： </td>
  <td width="85%">
  <input type="text" name="Label" id="Label" size="30" value="<%=rs("Label")%>" datatype="*" nullmsg="文本不能为空"> 
  </td>
  </tr>
      <tr  > <td > 提示说明：</td> <td><input type="text" name="Labe2" size="23" value="<%=rs("Labe2")%>"></td></tr>
       <tr  > <td> 变量名：</td>
       <td>
        
		<input type="text" name="FieldName" id="FieldName" size="7" value="<%=rs("FieldName")%>" datatype="*" nullmsg="变量名不能为空">
		
		 <input type="button" value="转换名称首字母" class="btnSearch" id="FieldName_6" onClick="GetPy('', 'Label', 'FieldName', 6, '', '');" />
<input type="button" value="转换名称全拼音" class="btnSearch" id="FieldName_3" onClick="GetPy('', 'Label', 'FieldName', 3, '', '');" /> 
		
		</td></tr>
		<tr  > 
		<td>类型：</td>
		
		<td>
        <select size="1" name="Type">
          <option value="t" <%if rs("type")="t" then Response.Write("selected")%>>输入框</option>
            <option value="r" <%if rs("type")="r" then Response.Write("selected")%>>单选</option>
          <option value="c"<%if rs("type")="c" then Response.Write("selected")%>>复选</option>
          <option value="l" <%if rs("type")="l" then Response.Write("selected")%>>列表框</option>
          <option value="m" <%if rs("type")="m" then Response.Write("selected")%>>备注型</option>
		  <option value="s" <%if rs("type")="s" then Response.Write("selected")%>>时间</option>
        </select>
        <select size="1" name="Required">
          <option value="0" <%if rs("Required")=0 then Response.Write("selected")%>>不必选</option>
          <option value="1" <%if rs("Required")=1 then Response.Write("selected")%>>必选</option>
        </select>
		</td>
		</tr>
		<tr>
		<td>高级验证：</td>
		<td>
		<select size="1" name="gaoji">     
              <option value="0" <%if rs("gaoji")=0 then Response.Write("selected")%>>无限制</option>
            
              <option value="1" <%if rs("gaoji")=1 then Response.Write("selected")%>>email</option>
              <option value="2" <%if rs("gaoji")=2 then Response.Write("selected")%>>身份证</option>
			  <option value="3" <%if rs("gaoji")=3 then Response.Write("selected")%>>必须为数字</option>
            
              <option value="4" <%if rs("gaoji")=4 then Response.Write("selected")%>>手机号码</option>
              <option value="5" <%if rs("gaoji")=5 then Response.Write("selected")%>>邮政编码</option>
			
            </select>  高级验证 必须是 必选条件下有
       </td></tr> 
	   <tr  > <td >单选和多选和列表框的值：</td>
	   <td>
        <input type="text" name="xxz" size="80" value="<%=rs("xxz")%>"> 
		<br /><br />
		当类型选择为单选或者是多选时请填写，其他请留空,多个值之间请用|隔开。
       当为列表框时：请选择性别|,男|男,女|女</td></tr>
        	
 
	  

      <tr  > <td>
      排序：</td>   <td><input name="px" type="text" id="px" value="<%=rs("px")%>" size="5" />
    </td>   
</tr>
   <tr  > <td colspan="2">  
      <input name="pid" type="hidden" id="pid" value="<%=rs("pid")%>" /><input type="submit" value="提交" name="B1" class="bnt"> 
    </td>   
</tr>  
</table>   
</form> 

</body>
</html>
