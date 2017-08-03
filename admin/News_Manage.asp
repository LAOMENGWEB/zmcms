<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%

if Instr(session("AdminPurview"),"|23,")=0 then 
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
If request.Form("button") = "设为推荐" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
dim sql4 
sql4="update news set sytj=1 where id in ("&Request("id")&")"
conn.Execute ( sql4 )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
call addlog("设置推荐"&xinwen&"")
end if


If request.Form("button") = "设为头条" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
dim sql20 
sql20="update news set gg=1 where id in ("&Request("id")&")"
conn.Execute ( sql20 )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
call addlog("设置头条")
end if

If request.Form("button") = "取消头条" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
dim sql21 
sql21="update news set gg=0 where id in ("&Request("id")&")"
conn.Execute ( sql21 )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
call addlog("取消头条")
end if






If request.Form("button") = "更新时间" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()

end if
dim sqlg 
if sqlno=1 then
sqlg="update news set adddate=now() where id in ("&Request("id")&")"
end if
if sqlno=2 then
sqlg="update news set adddate=getdate() where id in ("&Request("id")&")"
end if
conn.Execute ( sqlg )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
call addlog("更新"&xinwen&"的发表时间为最新时间")
end if

If request.Form("button") = "取消推荐" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
dim sql2 
sql2="update news set sytj=0 where id in ("&Request("id")&")"
call addlog("取消推荐"&xinwen&"")
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
conn.Execute ( sql2 )
end if
If request.Form("button") = "设为置顶" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()

end if
dim sql5 
sql5="update news set tt=1 where id in ("&Request("id")&")"
conn.Execute ( sql5 )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
call addlog("设为置顶"&xinwen&"")
end if
If request.Form("button") = "取消置顶" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
dim sql6 
sql6="update news set tt=0 where id in ("&Request("id")&")"
conn.Execute ( sql6 )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
call addlog("取消置顶"&xinwen&"")
end if
If request.Form("button") = "设为幻灯" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
dim sql7
sql7="update news set tj=1 where id in ("&Request("id")&")"
call addlog("设置幻灯"&xinwen&"")
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
conn.Execute ( sql7 )
end if
If request.Form("button") = "取消幻灯" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
dim sql8
sql8="update news set tj=0 where id in ("&Request("id")&")"
conn.Execute ( sql8 )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/info/","",""&Newclass&".html","info/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
Response.Write "<script>window.location.href='news_manage.asp';</script>" 
end if
call addlog("取消幻灯"&xinwen&"")
end if
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='news_manage.asp';</script>" 
response.end()
end if
for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	
set rs=server.CreateObject("adodb.recordset")
sql = "select * from news where id ="&ArticleID&""
rs.open sql,conn,1,3
zdfy=rs("zdfy")	
content=rs("content")
pone=rs("zmone") '一级分类
ptwo=rs("zmtwo") '二级分类
pthree=rs("zmthree") '三级分类


if Instr(rs("content"),"[zmcmsfy]")=0 then
call DeleteFile("../html/"&NewName&""&Separated&""&pone&""&Separated&""&ArticleID&""&Separated&"1.html")
else
arrContent=split(rs("content"),"[zmcmsfy]")
pages=Ubound(arrContent)+1
for j=1 to pages
call DeleteFile("../html/"&NewName&""&Separated&""&pone&""&Separated&""&ArticleID&""&Separated&""&j&".html") 
next
end if







rs.close:set rs=Nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from news where ID in ("&ArticleID&")",conn,0,1
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
conn.Execute ( "delete from news where id in ("&Request("id")&")" )
if yy=2 then
call htmlnewsfl
call htmlnews
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&iinfo&"/","",""&Newclass&".html",""&iinfo&"/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
end if
call addlog("删除"&xinwen&"")
response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""news_manage.asp""</script>"
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
function ConfirmDel()
{
   if(confirm("确定要删除选中的<%=xinwen%>吗？一旦删除将不能恢复！"))
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%=xinwen%>列表管理  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="xinwen.asp" target="main">返回菜单</a></div>
 <div class="dmeun"><a href="news_add.asp" target="main">添加<%=xinwen%></a></div>
 
 </td>
 

 </tr>
</table>


<table width="99%" align="center" cellpadding="0" cellspacing="0"  class=" dtable1">
  		
          <tr   >
            <td width="10%">按照分类查找：</td>
			<td width="55%"><select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from newsclass where ParentClassID=0 and lang='"&lang&"' order by px asc"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
	   Response.Write("<option value=""news_manage.asp"">所有分类信息</option>") 
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("ID") & """" & "")
		 If int(request("pone"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("ClassName") & " </option>")  & VbCrLf
		 
		    Sqlpp ="select * from newsclass Where ParentClassID="&Rsp("ID")&" and lang='"&lang&"'  order by px asc"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """?ptwo=" & Rspp("ID") & " """ & "")
				If int(request("ptwo"))=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">　|-" & Rspp("ClassName") & "")
				
				
				
				
				Response.Write("</option>" )   & VbCrLf
				
				
				
						    Sqlppp ="select * from newsclass Where ParentClassID="&Rspp("ID")&" and lang='"&lang&"'  order by px asc"   
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
        </select></td><form id="form1" name="form1" method="post" action="news_manage.asp?Result=ok">
            <td width="11%">
             快速搜索：</td>
			 <td width="24%">
                <input name="k1" type="text" id="k1" />
             
                            <label><input class="bnt" type="submit" name="button" id="button" value="搜索" /></label>
          </td> </form>
          </tr>
          
        
  </table>
<form name="zmqy" method="Post">
 <%
if Result="ok"   then
 DB_Sql = "select * from news where title like '%"&keytag&"%' and lang='"&lang&"' "
else
if pone="" then
DB_Sql = "select * from news where lang='"&lang&"' "
end if
if pone<>""  then
DB_Sql = "select * from news where zmone="&Pone&"  and lang='"&lang&"' "
end if
if ptwo<>""   then
DB_Sql = "select * from news where  zmtwo="&Ptwo&" and lang='"&lang&"' "
end if
if pthree<>""   then
DB_Sql = "select * from news where  zmthree="&Pthree&" and lang='"&lang&"' "
end if
end if
DB_Sql = DB_Sql & "order by tt desc, ID desc"
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
if Result="ok"   then
Obj_Page.Setpage_Pathurl = "news_manage.asp?Result=ok" '执行分页的页面
else

if pone="" then
Obj_Page.Setpage_Pathurl = "news_manage.asp" '执行分页的页面
end if
if pone<>"" then
Obj_Page.Setpage_Pathurl = "news_manage.asp?pone="&pone&"" '执行分页的页面
end if
if ptwo<>"" then
Obj_Page.Setpage_Pathurl = "news_manage.asp?ptwo="&ptwo&"" '执行分页的页面
end if
if pthree<>"" then
Obj_Page.Setpage_Pathurl = "news_manage.asp?pthree="&pthree&"" '执行分页的页面
end if
end if
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
  <table width="99%" align="center" cellpadding="0" cellspacing="0"  class="dtable2">
  		
          <tr class="wz">
            <td    ><div align="center">全选</div></td>
            <td   ><div align="center">编号</div></td>
            <td   ><div align="center">所属分类</div></td>
            <td   ><div align="center"><%=xinwen%>属性</div></td>
            <td   ><div align="center"><%=xinwen%>标题</div></td>
            <td   ><div align="center">查看组别</div></td>
            <td   ><div align="center">发布日期</div></td>
            <td   ><div align="center">发布人</div></td>
            <td   ><div align="center">浏览次数</div></td>
            <td  ><div align="center">管理操作</div></td>
          </tr>
<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
          <tr   > 
            <td width="5%" > 
              <div align="center">
  <input name="ID" type="checkbox" id="ID" value="<%=rs("id")%>" style="border:none"/>
 </div></td>
		    <td width="5%"><div align="center"><%=rs("id")%></div></td>
		    <td width="8%"><div align="center">
		      <%dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from newsclass",conn,1,1
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
		    <td width="9%"><div align="center">
			
			<%if rs("tt")=0 and rs("gg")=false and rs("sytj")=false and rs("tj")=false and  rs("ft")=false then
			
			response.Write("<font color=#FF0000>无</font>")
			end if
			%>
			
			
<%
if rs("tt")=1 then
response.Write("<font color=#FF0000>{置顶}</font>")
else
end if
	         
if rs("sytj")=1 then
response.Write("<font color=#FF0000>{推荐}</font>")
else
end if

if rs("gg")=1 then
response.Write("<font color=#FF0000>{头条}</font>")
else
end if

if rs("tj")=1 then
response.Write("<font color=#FF0000>{幻灯}</font>")
else
end if
%>
         
		    </div></td>
		    <td width="21%" style="padding-left:10px">
	




        <font color="<%=rs("btcolor")%>"><%=rs("title")%></font></td>
		    <td width="9%">  <div align="center">
		      <%if rs("Exclusive")=">=" then
        Response.Write ""&ViewGroupName(rs("GroupID"))&" <font color='green'>隶</font>" & vbCrLf
      else
        Response.Write ""&ViewGroupName(rs("GroupID"))&" <font color='red'>专</font>" & vbCrLf
	  end if%>
	        </div></td>
		    <td width="13%"><div align="center"><%= rs("adddate") %></div></td>
		    <td width="7%"><div align="center"><%= rs("user") %></div></td>
		    <td width="7%"><div align="center"><%= (rs("hits")) %></div></td>
		    <td width="16%" style="border-right:none"><div align="center"><a href="News_modi.asp?ID=<%=rs("id")%>">修改</a> 
		          <a href="News_Del.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel();">删除  </a>
            </div></td>
          </tr>
		      <%
rs.MoveNext
Next	
%>          <tr    >
<td><input type="checkbox" name="allbox" onClick="CheckAll()"  class="none"/></td>
           <td  colspan="10" style="padding:5px; text-align:left">
		   <input class="bnt" type="submit" name="button" id="button" value="更新时间" onClick="return confirm('确定要更新选中的<%=xinwen%>为最新时间吗？');" />
             <input class="bnt" type="submit" name="button" id="button" value="设为推荐"  onClick="return confirm('确定要推荐吗？');"/>
			  <input class="bnt" type="submit" name="button" id="button" value="取消推荐" onClick="return confirm('确定要取消推荐吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="设为置顶" onClick="return confirm('确定要设为置顶吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="取消置顶" onClick="return confirm('确定要取消置顶吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="设为幻灯" onClick="return confirm('确定要设为幻灯<%=xinwen%>吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="取消幻灯" onClick="return confirm('确定要取消幻灯<%=xinwen%>吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="设为头条" onClick="return confirm('确定要设为头条吗？');" />
			  <input class="bnt" type="submit" name="button" id="button" value="取消头条" onClick="return confirm('确定要取消头条吗？');" />
              <input class="bnt" type="submit" name="button" id="button" value="批量删除"   onClick="return confirm('确定要批量删除选中的<%=xinwen%>吗，将不可恢复?');"/>
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
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td style="padding:10px 0 ;border-right:none;border-bottom:none">	<div align="center">
      <%
Response.Write "目前暂无任何信息"
Response.end
	%>
    </div></td>
  </tr>
</table>

	
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If

%>
 
</form>

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
