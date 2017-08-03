 <%
Response.Charset="utf-8"
if session("username")="" or session("AdminPurview")="" or session("LoginSystem")<>"Succeed" then
     response.redirect "login.asp"
  end if
%>

<frameset rows="50,*"  frameborder="NO" border="0" framespacing="0">
<frame src="admin_top.asp" noresize="noresize" frameborder="NO" name="topFrame" scrolling="auto" marginwidth="0" marginheight="0" target="main" />
<frameset cols="120,*"  rows="2000,*" id="frame">
<frame src="left.asp" name="left" noresize="noresize" marginwidth="0" marginheight="0" frameborder="0" scrolling="auto" target="main" />
<frame src="right.asp" name="main" marginwidth="0" marginheight="0" frameborder="0" scrolling="auto" target="_self" />
<noframes>
  <body></body>
</noframes>
</html>