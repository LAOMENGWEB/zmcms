<!--#include file="../base.asp"-->
<%
dim id:id=enhtml(fget("id",0))
dim tablename:tablename="dy"
Dim rs,sql,rs2,sql2
Dim Title,Content,Addtime,OpenName,WeiXinName

	set rs=server.CreateObject("adodb.recordset")
	sql="select * from ["&tablename&"] where id="&clng(id)
	rs.open sql,conn,1,3
	if rs.eof then
		response.Write("数据出错")
		response.End()
	else
		Title=rs("title")
		Content=rs("content")
		Addtime=rs("adddate")
	end if
	rs.close
	set rs=nothing
	
	set rs2=server.CreateObject("adodb.recordset")
	sql2="select * from [sys_Config] where id=1"
	rs2.open sql2,conn2,1,3
	if not rs2.eof then
		OpenName=rs2("OpenName")
		WeiXinName=rs2("WeiXinName")
	end if
	rs2.close
	set rs2=nothing
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
    <div class="text"><%=nohtml1(nohtml(Content))%></div>
  </div>
</div>
</body>
</html>