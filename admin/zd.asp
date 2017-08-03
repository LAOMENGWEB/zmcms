<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="class_page.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">

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
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>


<%
id=Request.QueryString("id")
diyform=Request.QueryString("form")



If request.Form("button") = "设置后台显示" Then
if request.Form("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='zd.asp?form="&diyform&"&id="&id&"';</script>" 
response.end()
end if
dim sql3 
sql3="update FreeField set xs=1 where id in ("&request.Form("id")&")"
conn.Execute ( sql3 )
call addlog("审核信息")
end if

If request.Form("button") = "取消后台显示" Then
if request.Form("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='zd.asp?form="&diyform&"&id="&id&"';</script>" 
response.end()
end if
dim sql4 
sql4="update FreeField set xs=0 where id in ("&request.Form("id")&")"
conn.Execute (sql4)
call addlog("取消审核信息")
end if
%>

	



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">字段管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="bd.asp" target="main">自定义表单管理</a></div>

<div class="dmeun"><a href="zdadd.asp?form=<%=diyform%>&id=<%=id%>" target="main">添加字段</a></div>

 </td>
 

 </tr>
</table>




<%

 DB_Sql = "select * from FreeField where ChannelId="&id&" and lang='"&lang&"' "

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

Obj_Page.Setpage_Pathurl = "zd.asp" '执行分页的页面

Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<form  method="post" name="form1" id="form1"> 

            
              <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
               
                <tr class="wz">
	                <td width="66" >全选</td>
                  <td width="47"    >编号</td>
                  <td width="183"    >字段别名</td>
                  <td width="144"     >字段名称</td>
                <td width="144"     >字段类型</td>
                <td width="144"     >后台是否显示</td>
                <td width="144"     >添加时间</td>
                  <td width="402"> 操作</td>
                </tr>
              <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr  >
                 <td  ><input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" class="none"/></td>
                  <td  ><%=rs("id")%></td>
                  <td  ><%=rs("Alias")%></td>
              
                  <td   ><%=rs("FieldName")%></td>
                  <td   >
	                  <%if rs("FieldType")=1 then Response.write("文本框")%>
                      <%if rs("FieldType")=2 then Response.write("单选框")%>
	                  <%if rs("FieldType")=3 then Response.write("多选框")%>
		              <%if rs("FieldType")=4 then Response.write("列表框")%>
			          <%if rs("FieldType")=5 then Response.write("备注型")%>
				      <%if rs("FieldType")=6 then Response.write("时间")%>


	                  </td>
	                    <td   ><%if rs("xs")=1 then Response.write("显示") else Response.write("不显示")   end if%></td>
                    <td   ><%=rs("addtime")%></td>
               
                  <td >				  
                    <a href="zdxg.asp?form=<%=diyform%>&id=<%=rs("id")%>&id1=<%=id%>">修改</a>
                    &nbsp;
                    
                    <a href="zdDel.asp?form=<%=diyform%>&id=<%=rs("id")%>&id1=<%=id%>"onClick="return confirm('确定要删除该字段吗？');" >删除</a>
				  </td>
                </tr>
			
 <%
rs.MoveNext
Next
	
%>      
<tr  >
                  <td   > 
                   
                      <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
     </td>
      <td  colspan="7"  style="padding:5px; text-align:left">


<input class="bnt" type="submit" name="button" id="button" value="设置后台显示" />
					<input class="bnt" type="submit" name="button" id="button" value="取消后台显示" />


	      
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

	%>
	<table width="99%"  cellpadding="0" cellspacing="0" class="table1" style="margin-top:10px">
  		
          <tr   >
            <td  height="30"  style="padding-left:10px;border-bottom:none; text-align: center;">  <%Response.Write "目前暂无任何信息"
Response.end%> </td>       
  </table>

</form><%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</body>
</html>

