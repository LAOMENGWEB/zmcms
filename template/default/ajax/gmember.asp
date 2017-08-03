<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" --> 

<%  if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed" then %>请<a href="<%=sitepath%>member/memberuserlogin/" class="a-login">登陆</a>或<a href="<%=sitepath%>member/MemberRegister/" class="a-reg">免费注册</a><%else%><a href="<%=sitepath%>member/membercenter/" class="a-reg"><%=session("MemName")%>(<%=MemGroup(session("GroupID"))%>)</a> <a href="<%=sitepath%>inc/MemberLogin/?Action=Out" class="a-reg">退出</a><%end if%>

