<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Inc/sub.asp" -->  

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>提示信息-<%=mc%></title>
<style type="text/css">
body, h1, p, a, img {margin:0;padding:0;}
body{text-align:center;font-family:Arial, Helvetica, sans-serif, "宋体";}
h1,p{padding-left:100px; text-align:left;}
h1{font-size:14px; margin-top:12px;}
p{font-size:12px; line-height:150%;padding:10px 10px 10px 100px;line-height:18px;}
.box_border{margin:180px auto 0;width:420px; padding:20px 0;}
#right{border:1px solid #A6DFA6; background:#EEF9EE url(<%=sitepath%>images/bg_right.gif) no-repeat 10px center;}
#right h1{color:#237E29;}
#right a:link,a:visited{color: #237E29;text-decoration: none;}
#right a:hover,a:active{color: #237E29;text-decoration: underline;}
#wrong{border:1px solid #FDBD77; background:#FFFDD7 url(<%=sitepath%>images/bg_wrong.gif) no-repeat 10px center;}
#wrong h1{color:#f30;}
#wrong a:link,a:visited{color: #f30;text-decoration: none;}
#wrong a:hover,a:active{color: #0066FF;text-decoration: underline;}
#right a{color: #237E29;}
#wrong a{color: #f30;}

</style>
<body>
<%
type1=request.QueryString("type1")
Message=request.QueryString("Message")
url=request.QueryString("url")
m=request.QueryString("c")
%>

<div class="box_border" id="<%=m%>">
<h1><%=Message%></h1>
<p>
<%if type1=1 then%>
<a href="javascript:history.back(-1);" ><%=word(622)%></a>
<script language="javascript">setTimeout(function(){history.back();}, 3000);</script>
<%end if%>
<%if type1=2 then%>
<a href="<%=sitepath%><%=url%>"><%=word(622)%></a>
	<script language="javascript">setTimeout("location.href='<%=sitepath%><%=url%>';",3000);</script> 
<%end if%>
</p>
</div>
</body>
</html>