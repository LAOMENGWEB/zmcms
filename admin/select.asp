<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%

if Instr(session("AdminPurview"),"|36,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if


mode=request.QueryString("mode")
pone=request.QueryString("pone")


If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='select.asp?mode="&mode&"';</script>" 
response.end()
end if

conn.Execute ( "delete from SelectName where s_id in ("&Request("id")&")" )
conn.Execute ( "delete from SelectValue where s_id in ("&Request("id")&")" )
if mode=0 then
call addlog("批量删除筛选属性")
end if
if mode=1 then
call addlog("批量删除规格属性")
end if
if mode=2 then
call addlog("批量删除商品属性")
end if
end if








%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>
<%
if mode=0 then response.Write("筛选属性管理") end if
if mode=1 then response.write("规格属性管理") end if
if mode=2 then response.write("商品属性管理") end if
%>
</title>

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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%
if mode=0 then response.Write("筛选属性管理") end if
if mode=1 then response.write("规格属性管理") end if
if mode=2 then response.write("商品属性管理") end if
%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="shuxing.asp" target="main">返回菜单</a></div>
 
<%if mode=0 then%> <div class="dmeun"><a href="selectadd.asp?mode=0" target="main">添加筛选属性</a></div><%end if%>
 <%if mode=1 then%> <div class="dmeun"><a href="selectadd.asp?mode=1" target="main">添加商品规格</a></div><%end if%>
 <%if mode=2 then%> <div class="dmeun"><a href="selectadd.asp?mode=2" target="main">添加商品属性</a></div><%end if%>
 
 </td>
 

 </tr>
</table>



<table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable1">
  		
          <tr   >
            <td width="11%" >按照分类查找：</td>
			
			<td width="89%">
			<select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from SelectType where lang='"&lang&"' order by s_order asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
	   Response.Write("<option value=""select.asp?mode="&mode&""">所有信息</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("s_ID") & "&mode="&mode&"""" & "")
		 If int(request("pone"))=Rsp("s_ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("s_name") & " </option>")  & VbCrLf
		 
		    
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


DB_Sql = "select * from SelectName where s_mode="&mode&" and lang='"&lang&"' "

end if


if pone<>""  then
DB_Sql = "select * from SelectName where typeid="&Pone&" and s_mode="&mode&" and lang='"&lang&"' "
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
Obj_Page.Setpage_Pathurl = "select.asp?mode="&mode&"" '执行分页的页面

end if

if pone<>"" then
Obj_Page.Setpage_Pathurl = "select.asp?mode="&mode&"&typeid="&pone&"" '执行分页的页面

end if


Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
  <table width="99%" align="center" cellpadding="0" cellspacing="0" class="dtable2">
  		
          <tr  class="wz">
            <td ><div align="center">全选</div></td>
            <td  ><div align="center">编号</div></td>
            <td  ><div align="center">所属类别</div></td>
            <td ><div align="center">属性名称</div></td>
            <td ><div align="center">属性值</div></td>
            <td  ><div align="center">管理操作</div></td>
          </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
          <tr     style="line-height:3;"> 
            <td width="3%" > 
              <div align="center">
  <input name="ID" type="checkbox" id="ID" value="<%=rs("s_id")%>" style="border:none"/>
 </div></td>
		    <td ><div align="center"><%=rs("s_id")%></div></td>
		    <td ><div align="center">
		      <%dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from SelectType where lang='"&lang&"'",conn,1,1
		  while not rsc.eof
		    if rs("typeid")=rsc("s_id")  then
			response.Write("" & rsc("s_name") & "")
			end if
			
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %>
	        </div></td>
		    <td ><div align="center"><%= (rs("s_name")) %></div></td>
		    <td >
            
            <%if rs("s_type")=1 then%>
            属于文本框，无属性值
   
            <%
			else
Set rsc = conn.execute("select s_value from [SelectValue] where lang='"&lang&"' and s_id = "& rs("s_id") &"")
If Not rsc.eof Then 
		datavalue = rsc.getrows()
		datacount = 1
Else
		datacount = 0
End If 
rsc.close : Set rsc = Nothing 
If datacount = 1 Then 
		For ssss = 0 To UBound(datavalue,2)
				response.write(datavalue(0,ssss))
				If ssss < UBound(datavalue,2) Then response.write("、")
		Next 
End If 
end if
%>
            
            
            
            
            
            
            </td>
		    <td ><div align="center"><a href="selectxgname.asp?ID=<%=rs("s_id")%>&mode=<%=mode%>">修改属性名称</a> 
            <%if rs("s_type")<>1 then%>
             <a href="selectglvaule.asp?ID=<%=rs("s_ID")%>&mode=<%=mode%>&typeid=<%=rs("typeid")%>" >属性值管理  </a>
             <%end if%>
            
		          <a href="selectDel.asp?ID=<%=rs("s_ID")%>&mode=<%=mode%>" onClick="return ConfirmDel();">删除  </a>
	            </div></td>
          </tr>
		      <%
rs.MoveNext
Next	
%>          <tr    >
           <td   ><div align="center">
             <input type="checkbox" name="allbox" onClick="CheckAll()" style="border:none"/>
           </div></td>
           <td colspan="6" style="padding:5px; text-align:left">
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
	<table width="99%" align="center" cellpadding="0" cellspacing="0" >
  		
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
