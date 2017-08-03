<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|10415,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
pd=request.QueryString("pd")
dim action : action = request.QueryString("action")
dim id : id = request.QueryString("id")

Select case action	
	case "delall" : delAll
	case "del" : delInfo id
End Select





%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script> 
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->

</head>

<body>







<form method="post" name="form1" id="form1"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">自定义字段管理  <!--<span id="yuyan" style="margin-left:20px"></span>--></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="coulmnadd.asp" target="main">添加自定义字段</a></div></td>
 

 </tr>
</table>
    <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable1" >
  		
          <tr   >
            <td width="12%" >按照频道分类查找：</td>
			<td width="88%">
<select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
<option value="">请选择分类</option>
<option value="coulmn.asp">所有字段</option>
<option value="?pd=1" <%If int(request.QueryString("pd"))=1 then Response.Write("selected ") end if%> ><%=danye%> </option>
<option value="?pd=2" <%If int(request.QueryString("pd"))=2 then Response.Write("selected ") end if%> ><%=xinwen%> </option>
<option value="?pd=3" <%If int(request.QueryString("pd"))=3 then Response.Write("selected ") end if%> ><%=chanpin%> </option>
<option value="?pd=4" <%If int(request.QueryString("pd"))=4 then Response.Write("selected ") end if%> ><%=tupian%></option>
<option value="?pd=5" <%If int(request.QueryString("pd"))=5 then Response.Write("selected") end if%> ><%=shipin%></option>
<option value="?pd=6" <%If int(request.QueryString("pd"))=6 then Response.Write("selected ") end if%> ><%=xiazai%></option>
<option value="?pd=7" <%If int(request.QueryString("pd"))=7 then Response.Write("selected ") end if%> ><%=zhaopin%></option>
<option value="?pd=8" <%If int(request.QueryString("pd"))=8 then Response.Write("selected ") end if%> ><%=yingxiao%></option>
</select></td>

  </tr>
</table>   
<%
if pd<>"" then
if pd=1 then
'DB_Sql = "select * from Mycolumn where typeid=1 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=1 "
end if
if pd=2 then
'DB_Sql = "select * from Mycolumn where typeid=2 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=2 "
end if
if pd=3 then
'DB_Sql = "select * from Mycolumn where typeid=3 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=3 "
end if
if pd=4 then
'DB_Sql = "select * from Mycolumn where typeid=4 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=4 "
end if
if pd=5 then
'DB_Sql = "select * from Mycolumn where typeid=5 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=5 "
end if
if pd=6 then
'DB_Sql = "select * from Mycolumn where typeid=6 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=6 "
end if
if pd=7 then
'DB_Sql = "select * from Mycolumn where typeid=7 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=7 "
end if
if pd=8 then
'DB_Sql = "select * from Mycolumn where typeid=8 and lang='"&lang&"' "
DB_Sql = "select * from Mycolumn where typeid=8 "
end if
else
'DB_Sql = "select * from Mycolumn where lang='"&lang&"' "
DB_Sql = "select * from Mycolumn  "
end if

DB_Sql = DB_Sql & "order by Sequence"
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
if pd<>"" then
if pd=1 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=1" '执行分页的页面
end if
if pd=2 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=2" '执行分页的页面
end if
if pd=3 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=3" '执行分页的页面
end if
if pd=4 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=4" '执行分页的页面
end if
if pd=5 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=5" '执行分页的页面
end if
if pd=6 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=6" '执行分页的页面
end if
if pd=7 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=7" '执行分页的页面
end if
if pd=8 then
Obj_Page.Setpage_Pathurl = "coulmn.asp?pd=8" '执行分页的页面
end if
else
Obj_Page.Setpage_Pathurl = "coulmn.asp" '执行分页的页面
end if

Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>     
              <table width="99%"  align="center" cellpadding="0" cellspacing="0" class="dtable2">
               
                <tr class="wz">
                  <td ><div align="center">全选</div></td>
                  <td  ><div align="center">编号</div></td>
                  <td ><div align="center">排序</div></td>
                  <td ><div align="center"> 字段名称</div></td>
                  <td ><div align="center">字段说明</div></td>
                  <td >字段类型</td>
                  <td >所属频道</td>
                  <td ><div align="center"> 操作</div></td>
                </tr>
              <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr  >
                  <td align="center" ><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/></td>
                  <td align="center" ><%=rs("id")%></td>
                  <td align="center" ><iframe name="zmcms" width="150" height="50" frameborder="0" scrolling="no" 
				src="xgzdpx.asp?id=<%=rs("id")%>" ></iframe></td>
                  <td height="28" align="center" ><%=rs("ColumnName")%></td>
                  <td align="center" ><%=rs("Columninfo")%></td>
                  <td align="center" ><%if rs("ColumnType")=0 then response.Write("文本") end if 
				  if rs("ColumnType")=1 then response.Write("数字") end if
				  if rs("ColumnType")=2 Or rs("ColumnType")=3 then response.Write("备注") end if
				  %></td>
				  <td align="center" >
				  <%
				  if rs("TypeID")=1 then response.Write(""&danye&"") end if
				  if rs("TypeID")=2 then response.Write(""&xinwen&"") end if
				  if rs("TypeID")=3 then response.Write(""&chanpin&"") end if
				  if rs("TypeID")=4 then response.Write(""&tupian&"") end if
				  if rs("TypeID")=5 then response.Write(""&shipin&"") end if
				  if rs("TypeID")=6 then response.Write(""&xiazai&"") end if
				  if rs("TypeID")=7 then response.Write(""&zhaopin&"") end if
				  if rs("TypeID")=8 then response.Write(""&yingxiao&"") end if
				  
				  
				  %>
				  </td>
				  <td align="center" style="border-right:none"><a href="xgcoulmn.asp?id=<%=rs("id")%>">修改</a>  <a href="?action=del&id=<%=rs("id")%>" onClick="return confirm('确定要删除该字段信息吗？');">删除</a></td>
                </tr>
			
 <%
rs.MoveNext
Next
	
%>      
<tr  >
                  <td height="28" align="center" > 
                    <div align="center">
                      <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
      </div></td>
      <td height="28" colspan="7"  style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="批量删除"  onclick="if(confirm('确定要删除吗')){form1.action='?action=delall';}else{return false}"/>	  </td>
</tr>
  </table>

  <%If Obj_Page.OutErrInfo <> "" Then
		Response.Write Obj_Page.OutErrInfo '报错信息显示
		Response.End
		else
	Call Obj_Page.Page_Navigation() '分页导航
	end if
    else

	%>
	       
  	  <table width="99%" border="0" cellpadding="0" cellspacing="0" align="center"   class="table1" style="margin-top:10px">
<tr>
 <td >目前暂无任何信息</td>
</tr>
</table>

</form><%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
Sub delInfo(id)
    dim rsObj
	set rsObj=server.CreateObject("adodb.recordset")
	Sqlp ="select TypeID,ColumnName from Mycolumn where id="&id&""
	rsObj.open sqlp,conn,1,1 
	if not rsObj.eof then
	if sqlno=1 then
    insertSql1 = "alter table "&idToName(rsObj(0))&" drop "&rsObj(1)&""
	end if
	if sqlno=2 then
    insertSql1 = "alter table "&idToName(rsObj(0))&" drop COLUMN "&rsObj(1)&""
	end if
    conn.execute insertSql1
	end if
	rsObj.close
	set rsObj=nothing
	conn.Execute ( "delete from Mycolumn where id="&id&"" )	
	call addlog("删除字段信息")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""coulmn.asp""</script>"
End Sub

Sub delAll
	dim ids,idArray,idsLen,i
	ids =request.form("id")
	idArray=split(ids,",") : idsLen=ubound(idArray)
	for i=0 to idsLen
	delInfo idArray(i)
	next
call addlog("删除字段信息")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""coulmn.asp""</script>"
End Sub
Function idToName(byval WebType)
    select case cint(WebType)
		case 1 : idToName="dy"
		case 2 : idToName="news"
		case 3 : idToName="product"
		case 4 : idToName="al"
		case 5 : idToName="zxsp"
		case 6 : idToName="download"
		case 7 : idToName="job"
		case 8 : idToName="NetWork"
	end select
End Function
%>
</body>
</html>

