<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Inc/sub.asp" -->  

<%
 dim rs,sql,GroupLevel
 GroupID= request.QueryString("GroupID")
 Exclusive=request.QueryString("Exclusive")
 newsid=request.QueryString("newsid")
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from hyGroup where GroupID='"&GroupID&"' and Lang='"&Language&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing

  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then   ' 当用户权限数字小于模块权限数字, 则无权访问
	    response.write(""&word(531)&"")
		else
		response.write(""&N("新闻内容")&"")
	  end if
    case "="
      if not session("GroupLevel") = GroupLevel then
	response.write(""&word(531)&"")
		else
		response.write(""&N("新闻内容")&"")
      end if
  end select

%>
