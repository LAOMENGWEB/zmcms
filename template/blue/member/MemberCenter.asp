<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%>-<%=tw(70)%></title>
<meta name="keywords" content="<%=gjc%>" />
<meta name="description" content="<%=ms%>" />
<link href="<%=template%>css/style.css" rel="stylesheet" type="text/css">
<!--#include file="../include/js.asp"-->
</head>
<body >

<!--#include file="../top.asp"-->
<div class="wrap">
<div class="main">
 <!--左侧-->
<div class="fyLeft">
<dl class="l_pro">

<dt><%=tw(71)%></dt>

<dd><a href="<%=sitepath%>member/memberinfo/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(72)%></a></dd>
<dd><a href="<%=sitepath%>member/memberguest/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(73)%></a></dd>
<dd><a href="<%=sitepath%>member/memberorder/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(74)%></a></dd>
<dd><a href="<%=sitepath%>member/memberjob/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(75)%></a></dd>
<dd><a href="<%=sitepath%>inc/MemberLogin/?Action=Out&m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(76)%></a></dd>
</dl>
<dl class="l_content">
<dt><%=tw(4)%>：<%=lxdz%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(5)%>：<%=sj%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(6)%>：<%=cz%></dt>
</dl>
<dl class="l_content">
<dt><%=tw(7)%>：<%=lxr%></dt>
</dl>

</div>

<!--右侧-->
<div class="fyRight">
<div class="title"> <span class="fl"><%=tw(70)%></span>
<div class="fr"><%=tw(10)%>：<%=tw(70)%></div>
</div>
<% if session("MemName")<>"" or session("MemLogin")="Succeed" then %>
<font style="font-size:16px;font-weight:bold;"><%if qqopenapi<>"" then%> <%=tw(77)%> 丨 <a href="<%=sitepath%>plug/qqapi/user.asp?act=jb"><%=tw(78)%></a> <%else%><%=tw(79)%> 丨 <a href="<%=sitepath%>member/membercenter/?action=bangding"><%=tw(80)%></a><%end if%></font>
<%end if%>
<br />
<br />
<div class="mainRightMain">

<% if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed" then %>
				<table width="100%">
 					<tr>
					
 						<td width="80%" valign="top">
						  <%if session("GroupID")=999177 then%>
						   <span style="line-height:200%;">
                        	<font style="font-size:14px;color:#FF0000;font-weight:bold;"><%=tw(81)%></font><br />
                        <%=tw(82)%><br />
						
						
                          </span>
						   <%else%>
                        	<span style="line-height:200%;">
                        	<font style="font-size:14px;color:#FF0000;font-weight:bold;"><%=tw(83)%><a href="../memberuserlogin/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(84)%></a>/<a href="../MemberRegister/?m=<%=request.QueryString("m")%>&language=<%=language%>"><%=tw(85)%></a>？</font><br />
                        
							<%=tw(86)%><font color="#1874CD"><%=tw(87)%></font>，<%=tw(88)%>：<br />
 							<%=tw(89)%><br />
 							<%=tw(90)%><br />
 							<%=tw(91)%><br />
 							<%=tw(92)%><br />
 							……
                          </span>
						
						 
						  <%end if%>
                      </td>
 					</tr>
				</table>
				<% else %>
				<table width="100%">
					<tr>
 					
						<td width="80%" valign="top">
                        	<span style="line-height:200%;">
                        	<font style="font-size:14px;color:#FF0000;font-weight:bold;"><%=tw(93)%><%=session("MemName")%>，<%=tw(94)%></font><br />
                            <%=tw(86)%><font color="#1874CD"><%=MemGroup(session("GroupID"))%></font>，<%=tw(95)%>：<br />
						
 						<%=tw(89)%><br />
 							<%=tw(90)%><br />
 							<%=tw(91)%><br />
 							<%=tw(92)%><br />
 							……
                            </span>
                        </td>
 					</tr>
 				</table>
 				<% end if %>
</div>
</div>

</div>
</div>
<%
if request.QueryString("action")="bangding" then
call qqbangding
end if
%>
<%
sub qqbangding()
session("bangding")=1
Response.Redirect ""&sitepath&"plug/qqapi/"
End sub


%>

<!--#include file="../bottom.asp"-->

</body>
</html>
