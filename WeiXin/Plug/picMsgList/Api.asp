<!--#include file="../../Lib/base.asp"-->
<%
'插件将会传过来下面四个值，请根据值进行操作Session("FromUserName")|Session("ToUserName")|Session("k_plugParam")|Session("Content")
'插件API需输出标准的微信XML

'本接口为多图文接口，可返回XML格式的多图文给微信							
dim k_plugParam:k_plugParam=Session("k_plugParam")
dim FromUserName:FromUserName=Session("FromUserName")
dim ToUserName:ToUserName=Session("ToUserName")
dim Title,Descriptions,PicUrl,Url,Api_Token

'Api_Token用于本插件与其它网站插件进行校对的，可用系统的Session("Token"),也可自己设置一个Token
Api_Token=Session("Token")
'Api_Token="asdfw235lasf3525"

if k_plugParam<>0 then
	datadb=Cls.db.dbload("","p_IdList","[Plug_PicMsgList]","p_Id="&k_plugParam,"")
	if ubound(datadb)>=0 then
		datadb2=Cls.db.dbload("","p_Id,p_Title,p_Info,p_Pic,p_Type,p_Content,p_UrlType,p_Url","[Plug_PicMsg]","p_Id in ("&datadb(0,0)&")","")
		if ubound(datadb2)>=0 then
			dim t:t="<xml>"
			t=t&"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>"
			t=t&"<FromUserName><![CDATA["&tousername&"]]></FromUserName>"
			t=t&"<CreateTime>"&now&"</CreateTime>"
			t=t&"<MsgType><![CDATA[news]]></MsgType>"
			t=t&"<ArticleCount>"&ubound(datadb2,2)+1&"</ArticleCount>"
			t=t&"<Articles>"
			for i=0 to ubound(datadb2,2)
				Title=datadb2(1,i)
				Descriptions=datadb2(2,i)
				
				if datadb2(3,i)<>"" then
				picurl=weburl&picroot&datadb2(3,i)
				else
				picurl1="images/wu.gif"
				picurl=weburl&picroot&picurl1
				end if
				if datadb2(4,i)="图文" then
					Url=weburl&webroot&"plug/picMsg/show.asp?id="&datadb2(0,i)
				else
					if datadb2(6,i)="插件" then
						'FromUserName和signature主要是转微信OpenID过去【外部链接型插件】进行调用，属于直接跳转页面,另外，需要做好远程API的Token校验
						Session("times")=now()
						Dim times:times=Session("times")	'可用Times的Session来验证是否非法用户授权
						Url=datadb(7,0)&"?FromUserName="&FromUserName&"&ToUserName="&ToUserName&"&times="&ToUnixTime(times)&"&signature="&md5(fromusername&Api_Token)&""
					else
						Url=datadb2(7,i)
					end if
				end if
				t=t&"<item>"
				t=t&"<Title><![CDATA["&Title&"]]></Title>"
					t=t&"<Description><![CDATA["&Descriptions&"]]></Description>"
		
					t=t&"<PicUrl><![CDATA["&PicUrl&"]]></PicUrl>"
		
				t=t&"<Url><![CDATA["&Url&"]]></Url>"
				t=t&"</item>"
			next
			t=t&"</Articles>"
			t=t&"</xml>"
			Cls.echo t
			cls.die
		end if
	end if
end if
%>