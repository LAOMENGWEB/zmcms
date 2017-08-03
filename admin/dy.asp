<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
if Instr(session("AdminPurview"),"|13,")=0 then 
  response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<% 
Result=request.QueryString("Result")
keytag=request.Form("k1")
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='dy.asp';</script>" 
response.end()
end if
for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	


set rs=server.CreateObject("adodb.recordset")
sql = "select * from dy where id ="&ArticleID&""
rs.open sql,conn,1,3
pone=rs("zmone")
zdfy=rs("zdfy")	
content=rs("content")	

if Instr(rs("content"),"[zmcmsfy]")=0 then
call DeleteFile("../html/"&aboutname&""&Separated&""&pone&""&Separated&""&ArticleID&""&Separated&"1.html")
else
arrContent=split(rs("content"),"[zmcmsfy]")
pages=Ubound(arrContent)+1
for j=1 to pages
call DeleteFile("../html/"&aboutname&""&Separated&""&pone&""&Separated&""&ArticleID&""&Separated&""&j&".html") 
next
end if

rs.close:set rs=Nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from dy where ID in ("&ArticleID&")",conn,0,1

IcoImage=rs("SmallImage")
SmallImage = rs("SmallImage")
dim picUrl
dim picUrlArray
dim x,y
picUrl = getPicUrl(rs("content"))
if picUrl <> "" then
picUrl = left(picUrl,len(picUrl)-1)
picUrlArray = split(picUrl,",")
for x = 0 to ubound(picUrlArray)
if instr(picUrlArray(x),"") > 0 then
call DeleteFile(picUrlArray(x))
end if
next
end if

dim picUrl1
dim picUrlArray1
dim x1,y1
picUrl1 = rs("ImagePath")
if picUrl1 <> "" then
picUrlArray1 = split(picUrl1,"|")
for x1 = 0 to ubound(picUrlArray1)
call DeleteFile(picUrlArray1(x1))
next
end if
call deletefile1(SmallImage)
call deletefile(IcoImage)

next
rs.close:set rs=Nothing
conn.Execute ( "delete from dy where id in ("&Request("id")&")" )
if yy=2 then
call htmlabout
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
end if
call addlog("删除"&danye&"信息")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""dy.asp""</script>"
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



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%=danye%>列表管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="danye.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="dyadd.asp" target="main">添加<%=danye%></a></div></td>
 

 </tr>
</table>




<table width="99%"  cellpadding="0" cellspacing="0" class="dtable1" align="center">
  		
          <tr   >
            <td width="11%">按照分类查找：</td>
            <td width="54%"><select class="wbk" name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from mnclass where ParentClassID=0 and lang='"&lang&"' order by px asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
	   Response.Write("<option value=""dy.asp"">所有分类信息</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("ID") & """" & "")
		 If int(request("pone"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">" & Rsp("ClassName") & " </option>")  & VbCrLf
		 
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select></td>
            <form id="form2" name="form2" method="post" action="dy.asp?Result=ok">
            <td width="8%"  >
            快速搜索：                      </td>
            <td width="27%"  > <input name="k1" type="text" id="k1" />
             
              <input class="bnt" type="submit" name="button" id="button" value="搜索" /> </td>
        </form>
  </tr>
</table>


<%
if Result="ok"   then
 DB_Sql = "select * from dy where title like '%"&keytag&"%' and lang='"&lang&"' "
else
if pone="" then
DB_Sql = "select * from dy where lang='"&lang&"'  "
end if
if pone<>""  then
DB_Sql = "select * from dy where zmone="&Pone&" and lang='"&lang&"' "
end if
end if
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
if Result="ok" then
Obj_Page.Setpage_Pathurl = "dy.asp?Result=ok" '执行分页的页面
else
if pone="" then
Obj_Page.Setpage_Pathurl = "dy.asp" '执行分页的页面
else
Obj_Page.Setpage_Pathurl = "dy.asp?pone="&pone&"" '执行分页的页面
end if
end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<form action="dy.asp" method="post" name="form1" id="form1"> 

            
              <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
               
                <tr class="wz">
                  <td width="66" >全选</td>
                  <td width="47"    >编号</td>
                  <td width="183"    >排序号</td>
                  <td width="144"     >栏目名称</td>
                  <td width="150"    >所属分类</td>
                  <td width="196"    >查看组别</td>
                  <td width="196"    >是否外链</td>
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
                  <td  ><iframe name="zmcms" width="100" height="50" frameborder="0" scrolling="no" 
				src="xgdypx.asp?id=<%=rs("id")%>" ></iframe> </td>
              
                  <td   ><%=rs("Title")%></td>
                  <td  ><%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from mnclass where lang='"&lang&"'",conn,1,1
		  while not rsc.eof
		    if rs("zmone")=rsc("id") then
			response.Write("" & rsc("classname") & "")
			end if
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %></td>
                  <td  ><%if rs("Exclusive")=">=" then
        Response.Write ""&ViewGroupName(rs("GroupID"))&" <font color='green'>隶</font>" & vbCrLf
      else
        Response.Write ""&ViewGroupName(rs("GroupID"))&" <font color='red'>专</font>" & vbCrLf
	  end if%></td>
                  <td  ><%
				  if rs("wlurl")<>"" then 
				  Response.Write "是"
				  else
				  Response.Write "不是"
				  end if
				 %></td>
                  <td >				  
                    
                    &nbsp;<a href="dyModify.asp?ID=<%=rs("ID")%>">修改</a>&nbsp; 
                    &nbsp;<a href="dyDel.asp?ID=<%=rs("ID")%>" onClick="return confirm('确定要删除该<%=danye%>信息吗？');" >删除</a>
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
      <td  colspan="7"  style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="批量删除"  onClick="return confirm('确定要批量删除选中的<%=danye%>吗？');"/>
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
	<table width="99%"  cellpadding="0" cellspacing="0" class="table1">
  		
          <tr   >
            <td  height="30"  style="padding-left:10px;border-bottom:none">  <%Response.Write "目前暂无任何信息"
Response.end%> </td>       
  </table>

</form><%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</body>
</html>
<%
Function ViewGroupName(GruopID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from hyGroup where GroupID='"&GruopID&"' and lang='"&lang&"'"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    ViewGroupNameSi="未设组别"
  else
    ViewGroupName=rs("GroupName")
  end if
  rs.close
  set rs=nothing
end Function
%>
