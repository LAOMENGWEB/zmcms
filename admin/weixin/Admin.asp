
<%
Response.Charset="utf-8"
if session("username")="" or session("AdminPurview")="" or session("LoginSystem")<>"Succeed" then
response.write "您还没有登录或登录已超时，请<a href='../Login.asp' target='_parent'><font color='red'>返回登录</font></a>!"
response.end
end if
%>
