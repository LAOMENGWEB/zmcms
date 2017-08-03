<!--#include file="../../weixin/Lib/base.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../../weixin/Lib/Page.asp"-->
<%
if Instr(session("AdminPurview"),"|9999,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""../right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
	dim datadb,datadb2
	dim i	
%>
<link href="../../weixin/css/wx.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../../weixin/js/jquery.js"></script>
<script language="javascript" src="../../weixin/js/jquery.artDialog.js?skin=default"></script>
<script language="javascript" src="../../weixin/js/artDialog.iframeTools.js"></script>
<script language="javascript" src="../../weixin/js/base.js"></script>
<script language="javascript" src="../../weixin/js/admin.js"></script>
<script>
function checksearch(a) {
    return "" == a.key.value || "请输入关键字" == a.key.value ? ($.message({
        content: "请输入关键字"
    }), a.key.focus(), !1) : void 0
}
</script>
</head>
<body>

<%
key=Cls.enhtml(Cls.fget("key",0))
user=Cls.enhtml(Cls.fget("user",0))
keyWord=Cls.enhtml(Cls.fget("keyWord",0))
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >关键字匹配记录</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div></td>
 

 </tr>
</table>	




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >

<div class="clear_fixed">
  <div id="menu">
  <div class="search"><form action="?" method="get" onSubmit="return checksearch(this)"><input type="text" name="key"  value="<%if key<>"" then%><%=key%><%else%>请输入关键字<%end if%>" onFocus="if(this.value=='请输入关键字')this.value=''" onBlur="if(this.value=='')this.value='请输入关键字'" /><input type="submit" class="bnt" value="查询" /></form></div>
	<dl>
      <dt><span><a href="KeyWord.asp?act=add">添加关键字</a></span></dt>
      <dt><span><a href="KeyWord.asp">查看关键字</a></span></dt>
    </dl>
</div>
  <form>
 
  </td>
 

 </tr>
</table>
  
  
      <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
      <tr class="wz">
        <td>ID</td>
        <td>微信头像</td>
        <td>微信用户</td>
        <td>匹配关键字</td>
        <td>关键字备注</td>
        <td>添加时间</td>
      </tr>        
        <%
		dim whereStr
		whereStr="where 1=1 "
		If key<>"" then
			whereStr=whereStr&"and b.k_keyWord like '%"&key&"%'"
		end if
		If user<>"" then
			whereStr=whereStr&"and a.Openid = '"&user&"'"
		end if
		If keyWord<>"" then
			whereStr=whereStr&"and b.k_keyWord = '"&keyWord&"'"
		end if
		Set Page = new Page_List
		Page.Con = Cls.db.conn		
		dim SqlStr:SqlStr="Select a.id,a.k_id,a.Openid,a.addtime,b.k_keyWord,b.k_keyWord_beizhu,c.nickname,c.headimgurl from ([sys_KeyWord_hits] a left join [sys_KeyWord] b on a.k_id=b.k_id) left join [sys_User] c ON a.OpenID=c.openid "&whereStr&" order by a.id desc"
		Page.Sql = SqlStr
		Page.PageSize = 20
		Page.Style = 5
		Set Rs = Page.Rs
		
		If not Rs.Bof or Rs.Eof Then
			For i = 1 To Page.PageSize
			if Cls.strlen(rs("headimgurl"))=0 then
				UserAvatar="../images/user.jpg"
			else
				UserAvatar=rs("headimgurl")
			end if
			if Cls.strlen(rs("nickname"))=0 then
				WeixinName=rs("Openid")
			else
				WeixinName=rs("nickname")
			end if
				%>
                
      <tr id="list_<%=rs("id")%>">
        <td align="center" ><%=rs("id")%></td>
        <td align="center" ><img src="<%=UserAvatar%>" width="48"></td>
        <td align="center" ><a href="?user=<%=rs("Openid")%>"><%=WeixinName%></a></td>
        <td align="center" ><a href="?keyWord=<%=rs("k_keyWord")%>"><%=rs("k_keyWord")%></a></td>
        <td align="center" ><a href="?keyWord=<%=rs("k_keyWord")%>"><%=rs("k_keyWord_beizhu")%></a></td>
        <td align="center" ><%=cls.format_time(rs("addtime"),1)%></td>
      </tr><%
				Rs.MoveNext
			Next
		End If
		
		Call Page.ShowPage
		
		Rs.Close
		Set Rs = Nothing
		Set Page = Nothing%>
    </table>
  </form>
</div>
</body>
</html>
