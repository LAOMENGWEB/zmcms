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
		dim AppId,AppSecret,Access_Token,Verify
		AppId=Cls.enhtml(Cls.fpost("AppId",0))
		AppSecret=Cls.enhtml(Cls.fpost("AppSecret",0))
		if Cls.strlen(AppId)=0 then Cls.echo "0AppId不能为空":exit sub
		if Cls.strlen(AppSecret)=0 then Cls.echo "0AppSecret不能为空":exit sub		
		
		'远程验证AppId和AppSecret是否正确
		strJson=GetURL(GetTokenUtl&"grant_type=client_credential&appid="&AppId&"&secret="&Appsecret&"")
		if InStr(strJson,"errcode")>0 then Cls.echo "0AppId值与AppSecret校验错误，请重新填写正确接口":exit sub	
		Verify="已认证"	
		Call InitScriptControl:Set objTest = getJSONObject(strJson)
		Access_token=objTest.access_token	'获取新Access_token
		Expires_In=objTest.expires_in	
		'新的Access_token入库
		data=array(array("Access_token",Access_token,0,1),array("Expires_In",Expires_In,0,0),array("Token_Time",now(),50,1))	
		call Cls.db.dbupdate("[sys_Config]","id=1",data)
		
		'验证服务号是否经过认证
		if Cls.db.dbload("","WeiXinType","[sys_Config]","id=1","")(0,0)="服务号" then
			Call InitScriptControl:Set objTest = getJSONObject(strJson)
			Access_Token=objTest.access_token
			strJson=GetURL(GetUserInfoUrl&"access_token="&Access_Token&"&openid=1234567890")
			Call InitScriptControl:Set objTest = getJSONObject(strJson)
			if objTest.errcode="48001" then
				Verify="未认证"
			end if
		end if
				
		datadb=array(array("AppId",AppId,0,1),array("AppSecret",AppSecret,0,1),array("Verify",Verify,50,1))
		if Cls.db.dbupdate("[sys_Config]","id=1",datadb)=1 then
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
	if($.trim(the.AppId.value)=="")
	{
		$.message({content:"AppId不能为空"});
		the.AppId.focus();
		return false;
	}
	if($.trim(the.AppSecret.value)=="")
	{
		$.message({content:"AppSecret不能为空"});
		the.AppSecret.focus();
		return false;
	}
	var url,data;
	url="?act=setdb";
	data="AppId="+encodeURIComponent($.trim(the.AppId.value));	
	data+="&AppSecret="+encodeURIComponent($.trim(the.AppSecret.value));	
	
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
<%datadb=Cls.db.dbload("","AppId,AppSecret","[sys_Config]","id=1","")%>

<body>
<form onSubmit="return checkdata(this)">
<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >高级设置</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div></td>
 

 </tr>
</table>		   
		   
<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >高级接口 AppId：</td>
<td width="88%">
<input type="text" name="AppId" size="50" value="<%=datadb(0,0)%>" /> 
              请到微信平台查看AppId，用于自定义菜单等高级接口，不可为空
</td>
 </tr>	
 
 <tr > 
<td width="12%"   >高级接口 Appsecret：</td>
<td width="88%">
<input type="text" name="AppSecret" size="50" value="<%=datadb(1,0)%>" />
             请到微信平台查看Appsecret，用于自定义菜单等高级接口，不可为空为空
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
