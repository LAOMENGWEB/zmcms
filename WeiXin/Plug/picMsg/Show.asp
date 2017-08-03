<!--#include file="../../lib/base.asp"-->
<%
dim id:id=Cls.getint(Cls.fget("id",0),0)
Dim rs,sql
Dim Title,Content,Addtime,OpenName,WeiXinName

sql = "select p_Title,p_Addtime,p_Content from [Plug_PicMsg] where p_Type='图文' and p_Id="&id
Set rs = Server.CreateObject("Adodb.Recordset")
rs.Open sql, cls.db.Conn, 1, 1
if rs.eof then
	response.Write "数据出错"
	response.End()
end if
Title=rs("p_Title")
Content=rs("p_Content")
Addtime=cls.format_time(rs("p_Addtime"),1)
rs.close
set rs=nothing

sql = "select OpenName,WeiXinName from [sys_Config] where id=1"
Set rs = Server.CreateObject("Adodb.Recordset")
rs.Open sql, cls.db.Conn, 1, 1
OpenName=rs("OpenName")
WeiXinName=rs("WeiXinName")
rs.close
set rs=nothing
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title><%=Title%></title>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=0">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black">
<meta name="format-detection" content="telephone=no">
<link rel="stylesheet" type="text/css" href="css/client-page1cc06e.css">
<!--[if lt IE 9]>
    <link rel="stylesheet" type="text/css" href="http://res.wx.qq.com/mmbizwap/zh_CN/htmledition/style/pc-page1c98c1.css"/>
    <![endif]-->
<link media="screen and (min-width:1000px)" rel="stylesheet" type="text/css" href="css/pc-page1c98c1.css">
<style>
body {
	-webkit-touch-callout: none;
	-webkit-text-size-adjust: none;
}
#nickname {
	overflow:hidden;
	white-space:nowrap;
	text-overflow:ellipsis;
	max-width:90%;
}
ol, ul {
	list-style-position:inside;
}
#activity-detail .page-content .text {
	font-size:16px;
}
.text{ padding:10px 5px;}
</style>
</head>
<body id="activity-detail">
<div class="page-bizinfo">
  <div class="header">
    <h1 id="activity-name"><%=Title%></h1>
    <p class="activity-info"> <span id="post-date" class="activity-meta no-extra"><%=Addtime%></span> <span class="activity-meta"><%=OpenName%></span> <a href="javascript:viewProfile();" id="post-user" class="activity-meta"> <span class="text-ellipsis"><%=WeiXinName%></span><i class="icon_link_arrow"></i> </a> </p>
  </div>
</div>
<div id="page-content" class="page-content">
  <div id="img-content">
    <div class="text"><%=Content%>
    </div>
  </div>
</div>
</body>
</html>