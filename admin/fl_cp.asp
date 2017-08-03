<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
if Instr(session("AdminPurview"),"|21,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
response.end
end if
'========判断是否具有管理权限
If request.Form("button") = "设为显示" Then
if Request("chk")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='fl_cp.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update cpclass set xs=1 where id in ("&Request("chk")&")"
conn.Execute ( sql4 )
'判断是否有产品分类
set rs_num=server.CreateObject("adodb.recordset")
rs_num.open "select * from cpclass where lang='"&lang&"' ",conn,1,1 

total7=rs_num.recordcount
rs_num.close
set rs_num=nothing 

if yy=2 then
if total7<>0 then
	call htmlprofl
	call htmlpro
    call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
    call htmll("/"&product&"/","",""&Productclass&".html",""&product&"/index.asp","m=",mpro,"language=",lang,"","","","","","","","")
	end if
	end if
call addlog(""&chanpin&"栏目——设为显示")
end if
If request.Form("button") = "取消显示" Then
if Request("chk")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='fl_cp.asp';</script>" 
response.end()
end if
dim sql2 
sql2="update cpclass set xs=0 where id in ("&Request("chk")&")"
conn.Execute ( sql2 )
'判断是否有产品分类
set rs_num=server.CreateObject("adodb.recordset")
rs_num.open "select * from cpclass where lang='"&lang&"' ",conn,1,1 

total7=rs_num.recordcount
rs_num.close
set rs_num=nothing 

if yy=2 then
if total7<>0 then
	call htmlprofl
	call htmlpro
      call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
    call htmll("/"&product&"/","",""&Productclass&".html",""&product&"/index.asp","m=",mpro,"language=",lang,"","","","","","","","")
	end if
	end if
call addlog(""&chanpin&"栏目——取消显示")
end if
ID=checknum(request.QueryString("ID"),0)
ID1=checknum(request.QueryString("ID1"),0)
ID2=checknum(request.QueryString("ID2"),0)
if request("action")="Del" then	
set rschk=server.CreateObject("ADODB.RECORDSET")
if id<>""then
sqlchk="select * from cpclass where ParentClassID="+ID
end if
if id1<>"" then
sqlchk="select * from cpclass where ParentClassID="+ID1
end if
if id2<>"" then
sqlchk="select * from cpclass where ParentClassID="+ID2
end if
rschk.open sqlchk,conn,1,1
if not rschk.eof then
response.Write "<script language=javascript>alert(""请先删除该类别的子类"");window.history.back(-1);</script>"
response.End
end if
rschk.close
set rschk=nothing

set rschk1=server.CreateObject("ADODB.RECORDSET")
if id<>""  then
sqlchk1="select * from cpclass where ID="+ID
end if
if id1<>"" then
sqlchk1="select * from cpclass where ID="+ID1
end if
if id2<>""  then
sqlchk1="select * from cpclass where ID="+ID2
end if
	rschk1.open sqlchk1,conn,1,1
	ParentClassID1=checknum(rschk1("ParentClassID"),0)
	SmallImage=rschk1("SmallImage")
	rschk1.close
	set rschk1=nothing





'删除分类语句
if id<>"" then
sql="Delete from cpclass where id="+ID
end if
if id1<>"" then
sql="Delete from cpclass where id="+ID1
end if
if id<>"" then
sql="Delete from cpclass where id="+ID2
end if


'删除子分类的同时产品自动归属上一级栏目
if id1<>0 and id<>0 and id2=0 then
sql2="update product set zmtwo=0 where zmtwo="&ID1&""
end if
if id1<>0 and id<>0 and id2<>0 then
sql2="update product set zmthree=0 where zmthree="&ID2&""
end if







'删除分类的同时删除html文件条件


'一级分类
if id<>"" and id1=0 and id2=0 then
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from cpclass where id="&id&"",conn,1,1

delwjj("../"&rsc1("mulu")&"/")


totalrec1=Conn.Execute("Select count(*) from product where zmone="&id&" and lang='"&lang&"' ")(0)
totalpage1=int(totalrec1/productInfo)       
If (totalpage1 * productInfo)<totalrec1 Then
totalpage1=totalpage1+1
End If
if totalpage1<=1 then
call DeleteFile("../"&product&"/"&productclass&""&Separated&""&id&""&Separated&"1.html")
else
for f=1 to totalpage1
call DeleteFile("../"&product&"/"&productclass&""&Separated&""&id&""&Separated&""&f&".html")
next
End If
end if




'有二级分类
if id<>"" and id1<>"" and id2=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&id&"",conn,1,1
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from cpclass where id="&id1&"",conn,1,1
delwjj("../"&rsc2("mulu")&"/"&rsc3("mulu")&"/")
totalrec1=Conn.Execute("Select count(*) from product where zmone="&id&" and zmtwo="&id1&" and lang='"&lang&"' ")(0)
totalpage1=int(totalrec1/productInfo)       
If (totalpage1 * productInfo)<totalrec1 Then
totalpage1=totalpage1+1
End If
if totalpage1<=1 then
call DeleteFile("../"&product&"/"&productclass&""&Separated&""&id1&""&Separated&"1.html")
else
for f=1 to totalpage1
call DeleteFile("../"&product&"/"&productclass&""&Separated&""&id1&""&Separated&""&f&".html")
next
End If
end if






'删除全部html文件
totalrec1=Conn.Execute("Select count(*) from product where lang='"&lang&"'")(0)
totalpage1=int(totalrec1/productInfo)       
If (totalpage1 * productInfo)<totalrec1 Then
totalpage1=totalpage1+1
End If
if totalpage1<=1 then
call DeleteFile("../"&product&"/"&productclass&""&Separated&"1.html")
else
for i=1 to totalpage1
call DeleteFile("../"&product&"/"&productclass&""&Separated&""&i&".html")
next
End If
call DeleteFile("../"&product&"/"&productclass&".html")




'三级分类

if id<>"" and id1<>"" and id2<>"" then

set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from cpclass where id="&id&"",conn,1,1

set rsc5=server.CreateObject("adodb.recordset")
rsc5.open "select * from cpclass where id="&id1&"",conn,1,1


set rsc6=server.CreateObject("adodb.recordset")
rsc6.open "select * from cpclass where id="&id2&"",conn,1,1


delwjj("../"&rsc4("mulu")&"/"&rsc5("mulu")&"/"&rsc6("mulu")&"/")

totalrec1=Conn.Execute("Select count(*) from product where zmone="&id&" and zmtwo="&id1&" and zmthree="&id2&" and lang='"&lang&"' ")(0)
totalpage1=int(totalrec1/productInfo)       
If (totalpage1 * productInfo)<totalrec1 Then
totalpage1=totalpage1+1
End If
if totalpage1<=1 then
call DeleteFile("../"&product&"/"&productclass&""&Separated&""&id2&""&Separated&"1.html")
else
for f=1 to totalpage1
call DeleteFile("../"&product&"/"&productclass&""&Separated&""&id2&""&Separated&""&f&".html")
next
End If
end if


conn.execute(sql2)
call deletefile1(SmallImage)
conn.execute(sql)


ChildCount=conn.execute("select count(*) from [cpclass] where ParentClassID="&ParentClassID1&" and lang='"&lang&"'")(0)
conn.execute("update [cpclass] set c_child = "& checknum(ChildCount,0) &" where id="& ParentClassID1 &"")


'判断是否有产品分类
set rs_num=server.CreateObject("adodb.recordset")
rs_num.open "select * from cpclass where lang='"&lang&"' ",conn,1,1 

total7=rs_num.recordcount
rs_num.close
set rs_num=nothing 

if yy=2 then
if total7<>0 then
	call htmlprofl
	call htmlpro
    call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
    call htmll("/"&product&"/","",""&Productclass&".html",""&product&"/index.asp","m=",mpro,"language=",lang,"","","","","","","","")
	end if
	end if
call addlog("删除"&chanpin&"类别")
response.Write "<script language=javascript>alert(""类别删除成功"");window.location.href=""fl_cp.asp""</script>"
response.End
end if

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

<SCRIPT language=javascript>
function CheckAll(){ 
 for (var i=0;i<eval(zmqy.elements.length);i++){ 
  var e=zmqy.elements[i]; 
  if (e.name!="allbox") e.checked=zmqy.allbox.checked; 
 } 
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

 <form name="zmqy" method="Post">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%=chanpin%>分类管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chanpin.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="fladd_cp.asp" target="main">添加分类</a></div></td>
 

 </tr>
</table>





   <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
    <tr class="wz">
      <td   ><div align="center" >全选</div></td>
      <td ><div align="center" >ID</div></td>
      <td ><div align="center">分类图片</div></td>
      <td ><div align="center" >栏目管理</div></td>
      <td ><div align="center" >html存放路径</div></td>
      <td ><div align="center" >是否显示</div></td>
      <td ><div align="center" >排序</div></td>
      <td ><div align="center" >基本操作</div></td>
    </tr>
    <%
			   Set rs = Server.CreateObject("ADODB.Recordset")
			   sql="select * from cpclass where ParentClassID = 0 and lang='"&lang&"'  order by px asc"
			   rs.open sql,conn,1,1
				 
			   for i=1 to rs.recordcount
			   if rs.eof then
			   exit for
			   end if
			   str=rs("ClassName")

set rs_num=server.CreateObject("adodb.recordset")
rs_num.open "select * from cpclass where ParentClassID="&rs("id")&" and lang='"&lang&"'",conn,1,1 

pdnx=rs_num.recordcount
rs_num.close
set rs_num=nothing 
%>
    <tr >
      <td width="7%" ><div align="center">
          <input name="chk" type="checkbox" id="chk" value="<%=rs("id")%>" style="border:none"/>
      </div></td>
      <td width="4%" ><div align="center"><%=rs("ID")%></div></td>
      <td width="12%" ><div align="center">
        <%if rs("SmallImage")<>"" then%>
        <img src="../<%=rs("smallimage")%>" width="80" height="50">
        <%else%>
        <img src="../images/demo.jpg" width="80" height="50">
        <%end if%>
      </div></td>
      <td width="23%" ><div align="center">|-&nbsp;&nbsp;<%=str%></div></td>
      <td width="11%" style="padding-left:3px"><%=rs("mulu")%></td>
      <td width="8%" ><div align="center">
          <% if rs("xs")=1 then  Response.Write "<font color=red>显示</font>"  else Response.Write "<font color=red>不显示</font>"  end if%>
      </div></td>
      <td width="9%" ><div align="center">
          <iframe name="zmcms" width="100" height="50" frameborder="0" scrolling="no" 
				src="xgcpflpx.asp?id=<%=rs("id")%>" ></iframe>
      </div></td>
      <td width="26%" style="border-right:none"><div align="center"><a href="fladd_cp.asp?ID=<%=rs("ID")%>">添加子栏目</a>/<a href="flxg_cp.asp?ID=<%=rs("ID")%>">编辑</a>/<a href="?ID=<%=rs("ID")%>&action=Del" onClick="return confirm('确定删除？')">删除</a> 
          <%if pdnx=0 then%>
          <a href="productmanage.asp?pone=<%=rs("ID")%>"><font color="red">&nbsp;&nbsp;&nbsp;查看所属产品信息</font></a>
          <%end if%>
      </div></td>
    </tr>
    <%
			   Set rs1 = Server.CreateObject("ADODB.Recordset")
			   sql1="select * from cpclass where ParentClassID = "& rs("ID") &" and lang='"&lang&"'  order by px asc"
			   rs1.open sql1,conn,1,1
			   
			   for j=1 to rs1.recordcount
			   if rs1.eof then
			   exit for
			   end if
			   str1=rs1("ClassName")

set rs_num1=server.CreateObject("adodb.recordset")
rs_num1.open "select * from cpclass where ParentClassID="&rs1("id")&" and lang='"&lang&"'",conn,1,1 

pdnx1=rs_num1.recordcount
rs_num1.close
set rs_num1=nothing 
%>
    <tr >
      <td  ><div align="center">
          <input name="chk" type="checkbox" id="chk" value="<%=rs1("id")%>" style="border:none"/>
      </div></td>
      <td  ><div align="center"><%=rs1("ID")%></div></td>
      <td  ><div align="center">
        <%if rs1("SmallImage")<>"" then%>
        <img src="../<%=rs1("smallimage")%>" width="80" height="50">
        <%else%>
        <img src="../images/demo.jpg" width="80" height="50">
        <%end if%>
      </div></td>
      <td  ><div align="center" style="padding-left:50px">&nbsp;&nbsp;&nbsp;&nbsp;|-&nbsp;&nbsp;<%=str1%></div></td>
      <td  ><div  style="padding-left:3px"><%=rs("mulu")%>/<%=rs1("mulu")%></div></td>
      <td  ><div align="center">
          <% if rs1("xs")=1 then  Response.Write "<font color=red>显示</font>"  else Response.Write "<font color=red>不显示</font>"  end if%>
      </div></td>
      <td  ><div align="center">
        <div align="center">
            <iframe name="zmcms" width="100" height="50" frameborder="0" scrolling="no" 
				src="xgcpflpx.asp?id=<%=rs1("id")%>" ></iframe>
        </div>
      </div></td>
      <td  style="border-RIGHT:none"><div align="center"><a href="fladd_cp.asp?ID=<%=rs1("ID")%>">添加子栏目</a>/<a href="flxg_cp.asp?ID1=<%=rs1("ID")%>">编辑</a>/<a href="?ID=<%=rs("ID")%>&ID1=<%=rs1("ID")%>&action=Del" onClick="return confirm('确定删除？')">删除</a>
	  <%if pdnx1=0 then%>
	  <a href="productmanage.asp?ptwo=<%=rs1("ID")%>"><font color="red">&nbsp;&nbsp;&nbsp;查看所属产品信息</font></a></div></td>
	  <%END IF%>
    </tr>
	
	 <%
			   Set rs2 = Server.CreateObject("ADODB.Recordset")
			   sql2="select * from cpclass where ParentClassID = "& rs1("ID") &" and lang='"&lang&"'  order by px asc"
			   rs2.open sql2,conn,1,1
			   
			   for h=1 to rs2.recordcount
			   if rs2.eof then
			   exit for
			   end if
			   str2=rs2("ClassName")
		 %>
    <tr >
      <td  ><div align="center">
          <input name="chk" type="checkbox" id="chk" value="<%=rs2("id")%>" style="border:none"/>
      </div></td>
      <td  ><div align="center"><%=rs2("ID")%></div></td>
      <td  ><div align="center">
        <%if rs2("SmallImage")<>"" then%>
        <img src="../<%=rs2("smallimage")%>" width="80" height="50">
        <%else%>
        <img src="../images/demo.jpg" width="80" height="50">
        <%end if%>
      </div></td>
      <td  ><div align="center" style="padding-left:150px">&nbsp;&nbsp;&nbsp;&nbsp;|-&nbsp;&nbsp;<%=str2%></div></td>
           <td  ><div  style="padding-left:3px"> <%=rs("mulu")%>/<%=rs1("mulu")%>/<%=rs2("mulu")%></div></td>
      <td  ><div align="center">
          <% if rs2("xs")=1 then  Response.Write "<font color=red>显示"  else Response.Write "<font color=red>不显示</font>"  end if%>
      </div></td>
      <td  ><div align="center">
        <div align="center">
            <iframe name="zmcms" width="100" height="50" frameborder="0" scrolling="no" 
				src="xgcpflpx.asp?id=<%=rs2("id")%>" ></iframe>
        </div>
      </div></td>
      <td  style="border-right:none"><div align="center"><a href="flxg_cp.asp?ID2=<%=rs2("ID")%>">编辑</a>/<a href="?ID=<%=rs("ID")%>&ID1=<%=rs1("ID")%>&ID2=<%=rs2("ID")%>&action=Del" onClick="return confirm('确定删除？')">删除</a><a href="productmanage.asp?pthree=<%=rs2("ID")%>"><font color="red">&nbsp;&nbsp;&nbsp;查看所属产品信息</font></a></div></td>
    </tr>
	
	
	
	
	
    <%      rs2.movenext
			next
			rs2.close
			set rs2=nothing
			rs1.movenext
			next
			rs1.close
			set rs1=nothing
			rs.movenext
			next
			rs.close
			Set rs=nothing
    %>
    <tr >
      <td ><div align="center">
          <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
</div></td>
      <td colspan="8"  style="padding:5px; text-align:left" >
        <input class="bnt" type="submit" name="button" id="button" value="设为显示" onClick="return confirm('确定要显示选中的栏目吗？');"/>
        <input class="bnt" type="submit" name="button" id="button" value="取消显示" onClick="return confirm('确定要取消显示选中的栏目吗？');"/></td>
    </tr>
  </table>

</form>
</body>
</html>
