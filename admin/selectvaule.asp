<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|36,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")


If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='selectvaule.asp';</script>" 
response.end()
end if

conn.Execute ( "delete from SelectType where s_id in ("&Request("id")&")" )
call addlog("批量删除属性类别")
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>属性类别管理</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
        
        <SCRIPT language=javascript>
function CheckAll(){ 
 for (var i=0;i<eval(zmqy.elements.length);i++){ 
  var e=zmqy.elements[i]; 
  if (e.name!="allbox") e.checked=zmqy.allbox.checked; 
 } 
} 
function ConfirmDel()
{
   if(confirm("确定要删除选中的附属类别吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}

</SCRIPT>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->

</head>

<body>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">产品附属属性类别管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shuxing.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="selectvauleadd.asp" target="main">添加属性类别</a></div></td>
 

 </tr>
</table>

<table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable1">
  		
          <tr   >
            <td width="10%"  >按照分类查找：</td>
			
			<td width="90%">
			<select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from cpclass where ParentClassID=0 and lang='"&lang&"' order by px asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
	   Response.Write("<option value=""selectvaule.asp"">所有信息</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("ID") & """" & "")
		 If int(request("pone"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("ClassName") & " </option>")  & VbCrLf
		 
		    Sqlpp ="select * from cpclass Where ParentClassID="&Rsp("ID")&" and lang='"&lang&"'  order by px asc"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """?ptwo=" & Rspp("ID") & " """ & "")
				If int(request("ptwo"))=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">　|-" & Rspp("ClassName") & "")
				
				
				
				
				Response.Write("</option>" )   & VbCrLf
				
				
				
						    Sqlppp ="select * from cpclass Where ParentClassID="&Rspp("ID")&" and lang='"&lang&"'  order by px asc"   
   			Set Rsppp=server.CreateObject("adodb.recordset")   
   			rsppp.open sqlppp,conn,1,1
			Do while not Rsppp.Eof 
				Response.Write("<option value=" & """?pthree=" & Rsppp("ID") & " """ & "")
				If int(request("pthree"))=Rsppp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">　　　|-" & Rsppp("ClassName") & "")
				Response.Write("</option>" )   & VbCrLf
				
				Rsppp.Movenext   
      		Loop
			rsppp.close
			set rsppp=nothing
				
				
			Rspp.Movenext   
      		Loop
			rspp.close
			set rspp=nothing
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select>           </td>
          </tr>
          
        
  </table>

<form name="zmqy" method="Post">
<%
if pone="" then
DB_Sql = "select * from SelectType where lang='"&lang&"' "
end if
if pone<>""  then
DB_Sql = "select * from SelectType where zmone="&Pone&" and lang='"&lang&"' "
end if
if ptwo<>""   then
DB_Sql = "select * from SelectType where  zmtwo="&Ptwo&" and lang='"&lang&"' "
end if
if pthree<>""   then
DB_Sql = "select * from SelectType where  zmthree="&Pthree&" and lang='"&lang&"' "
end if

DB_Sql = DB_Sql & "order by s_ID desc"
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


if pone="" then
Obj_Page.Setpage_Pathurl = "selectvaule.asp" '执行分页的页面
end if
if pone<>"" then
Obj_Page.Setpage_Pathurl = "selectvaule.asp?pone="&pone&"" '执行分页的页面
end if
if ptwo<>"" then
Obj_Page.Setpage_Pathurl = "selectvaule.asp?ptwo="&ptwo&"" '执行分页的页面
end if
if pthree<>"" then
Obj_Page.Setpage_Pathurl = "selectvaule.asp?pthree="&pthree&"" '执行分页的页面
end if

Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
  <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable2">
  		
          <tr class="wz">
            <td ><div align="center">全选</div></td>
            <td   ><div align="center">编号</div></td>
            <td   ><div align="center">所属类别</div></td>
            <td   ><div align="center">属性类别</div></td>
            <td   style="border-right:none"><div align="center">管理操作</div></td>
          </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
          <tr     style="line-height:3;"> 
            <td width="10%" > 
              <div align="center">
  <input name="ID" type="checkbox" id="ID" value="<%=rs("s_id")%>" style="border:none"/>
 </div></td>
		    <td width="20%"><div align="center"><%=rs("s_id")%></div></td>
		    <td width="20%"><div align="center">
		      <%dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from cpclass",conn,1,1
		  while not rsc.eof
		    if rs("zmtwo")=0 and rs("zmthree")=0 and rs("zmone")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			if rs("zmone")<>0  and rs("zmthree")=0 and  rs("zmtwo")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			if rs("zmone")<>0 and rs("zmtwo")<>0 and  rs("zmthree")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %>
	        </div></td>
		    <td width="20%"><div align="center"><%= (rs("s_name")) %></div></td>
		    <td width="30%" style="border-right:none"><div align="center"><a href="selectvaulemodiy.asp?ID=<%=rs("s_id")%>">修改</a> 
		          <a href="selectvauleDel.asp?ID=<%=rs("s_ID")%>" onClick="return ConfirmDel();">删除  </a>
            </div></td>
          </tr>
		      <%
rs.MoveNext
Next	
%>          <tr    >
           <td   ><div align="center">
             <input type="checkbox" name="allbox" onClick="CheckAll()" style="border:none"/>
           </div></td>
           <td colspan="5" style="padding:5px; text-align:left">
             <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的附属类别吗，将不可恢复?');"/>
&nbsp;</td>
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
	<table width="99%" align="center" cellpadding="0" cellspacing="0" class="table1">
  		
          <tr   >
            <td  height="30"  style="padding-left:10px;border-bottom:none">  <%Response.Write "目前暂无任何信息"
Response.end%> </td>       
  </table>

	
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If

%>
 
</form>

</body>
</html>
