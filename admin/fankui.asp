<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="admin.asp"-->

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
<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(myform.elements.length);i++){ 
  var e=myform.elements[i]; 
  if (e.name!="allbox") e.checked=myform.allbox.checked; 
 } 
} 
--> 
</script> 

</head>

<body>



<%
id=Request.QueryString("id")
form=Request.QueryString("form")


If request.Form("button") = "批量删除" Then
if request.Form("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='fankui.asp?form="&form&"&id="&id&"';</script>" 
response.end()
end if

sql5="delete from "&form&" where id in ("&Request.form("id")&")"
conn.Execute ( sql5 )
end if

If request.Form("button") = "批量审核" Then
if request.Form("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='fankui.asp?form="&form&"&id="&id&"';</script>" 
response.end()
end if
dim sql3 
sql3="update "&form&" set sh=1 where id in ("&request.Form("id")&")"
conn.Execute ( sql3 )
call addlog("审核信息")
end if

If request.Form("button") = "取消审核" Then
if request.Form("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='fankui.asp?form="&form&"&id="&id&"';</script>" 
response.end()
end if
dim sql4 
sql4="update "&form&" set sh=0 where id in ("&request.Form("id")&")"
conn.Execute (sql4)
call addlog("取消审核信息")
end if

%>
	



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">表单反馈信息  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="bd.asp" target="main">自定义表单管理</a></div></td>
 

 </tr>
</table>


<form  method="POST" name="myform" action="fankui.asp?form=<%=form%>&id=<%=id%>"> 

<%
Set rs = Server.CreateObject("ADODB.Recordset")
DB_Sql = "select * from "&form&" where lang='"&lang&"' "
DB_Sql = DB_Sql & "order by ID desc"
rs.Open DB_Sql,Conn,1,1
If Not(rs.Bof and rs.Eof) Then


'================================================
'调用分页类开始
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 20 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum

Obj_Page.Setpage_Pathurl = "fankui.asp?form="&form&"&id="&id&" " '执行分页的页面

Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>



            
              <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
               
                <tr class="wz">
                  <td width="66" >全选</td>
                  <td width="47"    >编号</td>
  <%
  set rs2=server.createobject("adodb.recordset")
sql2="select * from FreeField where ChannelId="&id&" and xs=1 and lang='"&lang&"' order by px asc"
rs2.open sql2,conn,1,3
do while not rs2.eof

%>
  
  <TD    ><div align="center"><%=rs2("Alias")%></div></TD>
  <%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>




<td >是否回复</td>
<td >反馈时间</td>
<td >提交IP</td>
<td >审核状态</td>

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
                 
                <%
  set rs2=server.createobject("adodb.recordset")
sql2="select * from FreeField where ChannelId="&id&" and xs=1 and lang='"&lang&"' order by px asc"
rs2.open sql2,conn,1,3
do while not rs2.eof

%>
  
  <TD    ><%=rs(""&rs2("FieldName")&"")%></TD>
  <%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>
                 
                  <TD    >
	                  <%if rs("replay")<>"" then Response.write("已回复于 "&rs("replaytime")&"") else Response.write("未回复")  end if%>
                    
	                  </TD>
                  <TD    ><%=rs("adddate")%></TD>
                   <TD    ><%=rs("ip")%></TD>
               <TD    ><%if rs("sh")=1 then
			  response.Write("[已审核]")
			  else
			  response.Write("<font color=#FF0000>[未审核]</font>")
			  end if%></TD>
                  <td >				  
                    <a href="fankuixx.asp?form=<%=form%>&id1=<%=id%>&id=<%=rs("id")%>">回复详情</a>
                    &nbsp;
                   
                    <a href="fankuisc.asp?form=<%=form%>&id1=<%=id%>&id=<%=rs("id")%>" onClick="return confirm('确定要删除该反馈信息吗？');" >删除</a>
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
      <td  colspan="100"  style="padding:5px; text-align:left">

<input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的信息吗，将不可恢复?');"/>
<input class="bnt" type="submit" name="button" id="button" value="批量审核" />
					<input class="bnt" type="submit" name="button" id="button" value="取消审核" />


	      
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
	<table width="99%"  cellpadding="0" cellspacing="0"  class="table1" style="margin-top:10px">
  		
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

