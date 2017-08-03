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

	dim datadb
	dim i
	dim act:act=lcase(Cls.fget("act",0))
	select case act
		case "setdb":setdb
	end select
	
	sub setdb()
		dim Token,OpenName,WeiXinName,WeixinAvatar,WeiXinType,WeiXinOpenID
		Token=Cls.enhtml(Cls.fpost("Token",0))
		OpenName=Cls.enhtml(Cls.fpost("OpenName",0))
		WeiXinName=Cls.enhtml(Cls.fpost("WeiXinName",0))
		'WeixinAvatar=Cls.enhtml(Cls.fpost("WeixinAvatar",0))
		WeiXinType=Cls.enhtml(Cls.fpost("WeiXinType",0))
		WeiXinOpenID=Cls.enhtml(Cls.fpost("WeiXinOpenID",0))
		if Cls.strlen(Token)=0 then Cls.echo "0Token不能为空":exit sub
		if Cls.strlen(OpenName)=0 then Cls.echo "0微信公众平台名称不能为空":exit sub
		if Cls.strlen(WeiXinName)=0 then Cls.echo "0微信帐号不能为空":exit sub
		
		Dim data  'array("字段",t1,字符长度255|数字0,字符过滤1|数字0)
		data=array(array("Token",Token,0,1),array("OpenName",OpenName,0,1),array("WeiXinName",WeiXinName,0,1),array("WeiXinType",WeiXinType,50,1),array("WeiXinOpenID",WeiXinOpenID,50,1))
		if Cls.db.dbupdate("[sys_Config]","id=1",data)=1 then
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if
		cls.die	
	end sub
	
%>
<link href="../../weixin/css/wx.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../../weixin/js/jquery.js"></script>
<script language="javascript" src="../../weixin/js/jquery.artDialog.js?skin=default"></script>
<script language="javascript" src="../../weixin/js/artDialog.iframeTools.js"></script>
<script language="javascript" src="../../weixin/js/base.js"></script>
<script language="javascript" src="../../weixin/js/admin.js"></script>
<script type="text/javascript">
function checkdata(the)
{
	if($.trim(the.Token.value)=="")
	{
		$.message({content:"Token不能为空"});
		the.Token.focus();
		return false;
	}
	if($.trim(the.OpenName.value)=="")
	{
		$.message({content:"微信公众平台名称不能为空"});
		the.OpenName.focus();
		return false;
	}
	if($.trim(the.WeiXinName.value)=="")
	{
		$.message({content:"微信帐号不能为空"});
		the.WeiXinName.focus();
		return false;
	}
	if($.trim(the.WeiXinOpenID.value)=="")
	{
		$.message({content:"微信原始ID不能为空"});
		the.WeiXinOpenID.focus();
		return false;
	}
	var url,data;
	url="?act=setdb";
	data="Token="+encodeURIComponent($.trim(the.Token.value));	
	data+="&OpenName="+encodeURIComponent($.trim(the.OpenName.value));	
	data+="&WeiXinName="+encodeURIComponent($.trim(the.WeiXinName.value));		
	<!--data+="&WeixinAvatar="+encodeURIComponent($.trim(the.WeixinAvatar.value));	-->
	data+="&WeiXinType="+encodeURIComponent($.trim(the.WeiXinType.value));			
	data+="&WeiXinOpenID="+encodeURIComponent($.trim(the.WeiXinOpenID.value));		
	
	$.ajax({
	type:"post",
	cache:false,
	url:url,
	data:data,
	error:function(){$.message({type:"error",content:"服务器错误，操作失败！"});},
	success:function(_)
	{
		var act=_.substring(0,1);
		var info=_.substring(1);
		switch(act)
		{
			case "0":
				$.message({type:"error",content:info});
				break;
			case "1":
				$.message({type:"ok",content:"保存成功"});
				setTimeout("location.href='?'",1000);
				break;
			default:
				alert(_)
				break;
		}
	}
	});
	return false

}
</script>
</head>
<%datadb=Cls.db.dbload("","Token,OpenName,WeiXinName,WeixinAvatar,WeiXinType,WeiXinOpenID,Verify","[sys_Config]","id=1","")%>

<body>
<form onSubmit="return checkdata(this)">
		   
<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >微信系统设置</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div></td>
 

 </tr>
</table>

<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >微信API接口：</td>
<td width="88%">
<input type="text" name="url" size="50" value="http://<%=Request.ServerVariables("SERVER_NAME")%><%=webroot%>Weixin_Api.asp" readonly="readonly"/>
    请将此API填写到微信平台 （自动获取无需更改）
</td>
 </tr>

<tr > 
<td width="10%"   >token：</td>
<td width="88%">
<input type="text" name="Token" size="50" value="<%=datadb(0,0)%>"/>
请将此Token填写到微信平台
</td>
</tr> 

<tr > 
<td width="10%"   >公众号名称：</td>
<td width="88%">
<input type="text" name="OpenName" size="50" value="<%=datadb(1,0)%>" />
</td>
 </tr> 
 
 
 <tr > 
<td width="10%"   >微信账号：</td>
<td width="88%">
<input type="text" name="WeiXinName" size="50" value="<%=datadb(2,0)%>" />
</td>
 </tr> 
 
  <tr > 
<td width="10%"   >微信原始ID：</td>
<td width="88%">
<input type="text" name="WeiXinOpenID" size="50" value="<%=datadb(5,0)%>" /> 例如:gh_d65b7fe5b062
</td>
 </tr> 
 
   <tr > 
<td width="10%"   >公众号类型：</td>
<td width="88%">
  <select name="WeiXinType" id="WeiXinType">
          <option value="订阅号" <%if datadb(4,0)="订阅号" then Cls.echo "selected"%>>订阅号</option>
          <option value="服务号" <%if datadb(4,0)="服务号" then Cls.echo "selected"%>>服务号</option>
        </select> 正确选择公众号类型
</td>
 </tr> 
 <!--
    <tr > 
<td width="10%"   >公众号头像：</td>
<td width="88%">
<input type="text" name="WeixinAvatar" size="50" value="<%=datadb(3,0)%>" /> 可直接填写头像地址
</td>
 </tr> 
 -->
 
     <tr > 
<td width="10%"   >认证状态：</td>
<td width="88%">
<%if datadb(6,0)="未认证" then%><span style="color:#F00; font-size:12px;">未认证</span><%elseif datadb(6,0)="已认证" then%><span style="color:#00F; font-size:12px;">已认证</span><%else%>未知<%end if%>
</td>
 </tr> 
 
      <tr > 
<td width="10%"   >权限说明：</td>
<td width="88%">
您是属于<span style="color:#F00; font-size:12px;"><%=datadb(6,0)%></span>的【<%=datadb(4,0)%>】,
           <%if datadb(4,0)="服务号" and datadb(6,0)="已认证" then%>
           您可以使用【自定义菜单】，【粉丝管理】，【回复粉丝】等高级功能
           <%elseif datadb(4,0)="服务号" and datadb(6,0)="未认证" then%>
           您可以使用【自定义菜单】功能，不可使用<del style="color:#999;">【粉丝管理】</del>，<del style="color:#999;">【回复粉丝】</del>等高级功能
           <%elseif datadb(4,0)="订阅号" and datadb(6,0)="已认证" then%>
           您可以使用【自定义菜单】功能，不可使用<del style="color:#999;">【粉丝管理】</del>，<del style="color:#999;">【回复粉丝】</del>等高级功能
           <%else%>
           您只可以使用基本功能，需要认证后才可以使用【自定义菜单】功能;<br>只有服务号才可以使用【粉丝管理】，【回复粉丝】等高级功能
           <%end if%>
</td>
 </tr> 
   <tr > 
            <td colspan="2"> <input class="bnt" type="submit" name="Submit" value="保存">
            　            </td>
          </tr>
 
</table>

          </form>
</body>
</html>
