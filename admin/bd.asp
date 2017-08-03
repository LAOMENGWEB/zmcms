<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|107,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
%>
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
If request.Form("button") = "开启表单" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='bd.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update freeform set IsStatus=1 where id in ("&Request("id")&")"
conn.Execute ( sql4 )

call addlog("开启表单")
end if
If request.Form("button") = "关闭表单" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='bd.asp';</script>" 
response.end()
end if
dim sql5
sql5="update freeform set IsStatus=0 where id in ("&Request("id")&")"
conn.Execute ( sql5 )
call addlog("关闭表单")
end if

%>



	


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 自定义表单管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="bdadd.asp" target="main">添加自定义表单</a></div></td>
 

 </tr>
</table>




<%

 DB_Sql = "select * from FreeForm where lang='"&lang&"' "

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

Obj_Page.Setpage_Pathurl = "bd.asp" '执行分页的页面

Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<form action="bd.asp" method="post" name="form1" id="form1"> 

            
              <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
               
                <tr class="wz">
                  <td width="66" >全选</td>
                  <td width="47"    >编号</td>
                  <td width="183"    >表单名称</td>
                  <td width="144"     >数据表名称</td>
                <td width="144"     >是否开启</td>
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
                  <td  ><%=rs("FormName")%></td>
              
                  <td   ><%=rs("FormTableName")%></td>
                    <td   ><%if rs("IsStatus")=0 then Response.write("关闭") else Response.write("开启")  end if%></td>
               
                  <td >				  
                    <a href="bdxg.asp?id=<%=rs("id")%>">修改</a>
                    &nbsp;
                     <a href="zd.asp?form=<%=rs("FormTableName")%>&id=<%=rs("id")%>">字段管理</a>
                    &nbsp;
                    <a href="zdadd.asp?form=<%=rs("FormTableName")%>&id=<%=rs("id")%>">添加字段</a>
                    &nbsp;
                    <a href="fankui.asp?form=<%=rs("FormTableName")%>&id=<%=rs("id")%>">查看反馈</a>
                    &nbsp;
                    <a href="bdDel.asp?ID=<%=rs("ID")%>" onClick="return confirm('确定要删除该表单信息吗？将要一起删除 数据表？');" >删除</a>
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

<input class="bnt" type="submit" name="button" id="button" value="开启表单"  onClick="return confirm('确定要开启表单吗？');"/>

<input class="bnt" type="submit" name="button" id="button" value="关闭表单"  onClick="return confirm('确定要关闭表单吗？');"/>


	      
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

