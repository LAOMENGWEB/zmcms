<!--#include file="../../Lib/base.asp"-->
<%
'插件将会传过来下面四个值，请根据值进行操作Session("FromUserName")|Session("ToUserName")|Session("k_plugParam")|Session("Content")
'插件API需输出标准的微信XML

'本接口为单图文接口，可返回XML格式的单图文给微信				
dim k_plugParam:k_plugParam=Session("k_plugParam")
dim FromUserName:FromUserName=Session("FromUserName")
dim ToUserName:ToUserName=Session("ToUserName")
dim Title,Descriptions,PicUrl,Url,Api_Token

'Api_Token用于本插件与其它网站插件进行校对的，可用系统的Session("Token"),也可自己设置一个Token
Api_Token=Session("Token")
'Api_Token="asdfw235lasf3525"

if k_plugParam<>0 then
	datadb=Cls.db.dbload("","p_Id,p_Title,p_Info,p_Pic,p_Type,p_Content,p_UrlType,p_Url","[Plug_PicMsg]","p_Id="&k_plugParam,"")
	if ubound(datadb)>=0 then
		Title=datadb(1,0)
		Descriptions=datadb(2,0)
		
		
		if datadb(3,0)<>"" then
				picurl=weburl&picroot&datadb(3,0)
				else
				picurl1="images/wu.gif"
				picurl=weburl&picroot&picurl1
				end if
		
		
		
		
		
		
		
		if datadb(4,0)="图文" then
			Url=weburl&webroot&"plug/picMsg/show.asp?id="&datadb(0,0)
		else
			if datadb(6,0)="插件" then
				'FromUserName和signature主要是转微信OpenID过去【外部链接型插件】进行调用，属于直接跳转页面,另外，需要做好远程API的Token校验
				Session("times")=now()
				Dim times:times=Session("times")	'可用Times的Session来验证是否非法用户授权
				Url=datadb(7,0)&"?FromUserName="&FromUserName&"&ToUserName="&ToUserName&"&times="&ToUnixTime(times)&"&signature="&md5(fromusername&Api_Token)&""
			else
				Url=datadb(7,0)
			end if
		end if
		cls.echo RequestSendPicText(FromUserName,ToUserName,Title,Descriptions,PicUrl,Url)
	end if
end if

function RequestSendPicText(fromusername,tousername,title,descriptions,picurl,url)
	dim t:t="<xml>"
	t=t&"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>"
	t=t&"<FromUserName><![CDATA["&tousername&"]]></FromUserName>"
	t=t&"<CreateTime>"&now&"</CreateTime>"
	t=t&"<MsgType><![CDATA[news]]></MsgType>"
	t=t&"<ArticleCount>1</ArticleCount>"
	t=t&"<Articles>"
	t=t&"<item>"
	t=t&"<Title><![CDATA["&title&"]]></Title>"
		t=t&"<Description><![CDATA["&descriptions&"]]></Description>"
		t=t&"<PicUrl><![CDATA["&picurl&"]]></PicUrl>"
	t=t&"<Url><![CDATA["&url&"]]></Url>"
	t=t&"</item>"
	t=t&"</Articles>"
	t=t&"</xml>"
	RequestSendPicText=t
end function
%>