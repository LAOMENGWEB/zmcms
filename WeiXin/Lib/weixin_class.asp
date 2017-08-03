<%
	'微信接口函数
	Dim GetTokenUtl:GetTokenUtl="https://api.weixin.qq.com/cgi-bin/token?"	'获取access_token接口
	Dim SendMsgUrl:SendMsgUrl="https://api.weixin.qq.com/cgi-bin/message/custom/send?"	'发送消息接口
	Dim SendMenuUrl:SendMenuUrl="https://api.weixin.qq.com/cgi-bin/menu/"	'自定义菜单接口
	Dim GetUserListUrl:GetUserListUrl="https://api.weixin.qq.com/cgi-bin/user/get?"	'获取粉丝列表
	Dim GetUserInfoUrl:GetUserInfoUrl="https://api.weixin.qq.com/cgi-bin/user/info?"	'获取粉丝信息接口
	
	'被动返回文本消息
	function RequestSendText(fromusername,tousername,returnstr)
		RequestSendText="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & dehtml(returnstr) & "]]></Content>" &_
		"</xml>"
	end function
	
	function GetRule(msgtype)
		msgtype=enhtml(msgtype)
		dim IsFoundKey:IsFoundKey=0
		dim sqltime:sqltime=now()
		if msgtype="接收图片信息" or msgtype="接收语音信息" or msgtype="接收视频信息" then
			Set Rs = Server.CreateObject("Adodb.RecordSet")
			SQL = "select top 1 k_id,k_text,k_type,k_plugName,k_plugParam from [sys_KeyWord] where k_keyWord_beizhu='"&msgtype&"' order by k_id desc"
			Rs.Open SQL,Conn,1,3
			If Not (Rs.Eof And Rs.Bof) Then
				IsFoundKey=1
				conn.Execute("update [sys_KeyWord] set k_hits=k_hits+1 where k_id="&rs("k_id"))			'记录匹配关键字的次数
				conn.Execute("insert into [sys_KeyWord_hits](k_id,Openid,Addtime) values("&rs("k_id")&",'"&FromUserName&"','"&sqltime&"')")  '记录每一次关键字匹配的时间
				if Rs("k_type")="文本" then
					GetRule=RequestSendText(FromUserName,ToUserName,Rs("k_text"))	'直接返回文本
				else
					GetRule=GetPlug(Rs("k_plugName"),Rs("k_plugParam"))		'调用插件数据
				end if
			End if
			Rs.Close
		end if
	end function
	
	'查找关键字
	function GetKeyWord(keyword)
		keyword=enhtml(keyword)
		dim IsFoundKey:IsFoundKey=0
		dim sqltime:sqltime=now()
		
		Set Rs = Server.CreateObject("Adodb.RecordSet")
		SQL = "select top 1 k_id,k_text,k_type,k_plugName,k_plugParam from [sys_KeyWord] where k_keyType='完全匹配' and k_keyWord = '"&keyword&"' order by k_id desc"
		Rs.Open SQL,Conn,1,3
		If Not (Rs.Eof And Rs.Bof) Then
			IsFoundKey=1
			conn.Execute("update [sys_KeyWord] set k_hits=k_hits+1 where k_id="&rs("k_id"))			'记录匹配关键字的次数
			conn.Execute("insert into [sys_KeyWord_hits](k_id,Openid,Addtime) values("&rs("k_id")&",'"&FromUserName&"','"&sqltime&"')")  '记录每一次关键字匹配的时间
			if Rs("k_type")="文本" then
				GetKeyWord=RequestSendText(FromUserName,ToUserName,Rs("k_text"))	'直接返回文本
			else
				GetKeyWord=GetPlug(Rs("k_plugName"),Rs("k_plugParam"))		'调用插件数据
			end if
			else
			GetKeyWord=RequestSendText(FromUserName,ToUserName,"获取不到相关的信息")	'直接返回文本
		End if
		Rs.Close
				
		if IsFoundKey=0 then
			SQL = "select k_id,k_text,k_type,k_plugName,k_plugParam from [sys_KeyWord] where k_keyType='左匹配' and k_keyWord like '"&keyword&"%'"
			Rs.Open SQL,Conn,1,3
			If Not (Rs.Eof And Rs.Bof) Then
			IsFoundKey=1
				conn.Execute("update [sys_KeyWord] set k_hits=k_hits+1 where k_id="&rs("k_id"))	
				conn.Execute("insert into [sys_KeyWord_hits](k_id,Openid,Addtime) values("&rs("k_id")&",'"&FromUserName&"','"&sqltime&"')")
				if Rs("k_type")="文本" then
					GetKeyWord=RequestSendText(FromUserName,ToUserName,Rs("k_text"))	
				else
					GetKeyWord=GetPlug(Rs("k_plugName"),Rs("k_plugParam"))		
				end if
				else
				GetKeyWord=RequestSendText(FromUserName,ToUserName,"获取不到相关的信息")	'直接返回文本
			End if
			Rs.Close
		end if
		
		if IsFoundKey=0 then
			SQL = "select k_id,k_text,k_type,k_plugName,k_plugParam from [sys_KeyWord] where k_keyType='右匹配' and k_keyWord like '%"&keyword&"'"
			Rs.Open SQL,Conn,1,3
			If Not (Rs.Eof And Rs.Bof) Then
				IsFoundKey=1
				conn.Execute("update [sys_KeyWord] set k_hits=k_hits+1 where k_id="&rs("k_id"))	
				conn.Execute("insert into [sys_KeyWord_hits](k_id,Openid,Addtime) values("&rs("k_id")&",'"&FromUserName&"','"&sqltime&"')")
				if Rs("k_type")="文本" then
					GetKeyWord=RequestSendText(FromUserName,ToUserName,Rs("k_text"))	
				else
					GetKeyWord=GetPlug(Rs("k_plugName"),Rs("k_plugParam"))		
				end if
				else
				GetKeyWord=RequestSendText(FromUserName,ToUserName,"获取不到相关的信息")	'直接返回文本
			End if
			Rs.Close	
		end if
		
		if IsFoundKey=0 then
			if datatype=true then
				SQL = "select k_id,k_text,k_type,k_plugName,k_plugParam from [sys_KeyWord] where k_keyType='模糊匹配' and instr(k_keyWord,'"&keyword&"')>0"
			else
				SQL = "select k_id,k_text,k_type,k_plugName,k_plugParam from [sys_KeyWord] where k_keyType='模糊匹配' and instr(k_keyWord,'"&keyword&"')>0"
			end if
			Rs.Open SQL,Conn,1,3
			If Not (Rs.Eof And Rs.Bof) Then
				IsFoundKey=1
				conn.Execute("update [sys_KeyWord] set k_hits=k_hits+1 where k_id="&rs("k_id"))	
				conn.Execute("insert into [sys_KeyWord_hits](k_id,Openid,Addtime) values("&rs("k_id")&",'"&FromUserName&"','"&sqltime&"')")
				if Rs("k_type")="文本" then
					GetKeyWord=RequestSendText(FromUserName,ToUserName,Rs("k_text"))
				else
					GetKeyWord=GetPlug(Rs("k_plugName"),Rs("k_plugParam"))
				end if
				else
				GetKeyWord=RequestSendText(FromUserName,ToUserName,"获取不到相关的信息")	'直接返回文本
			End if
			Rs.Close	
		end if
		
		if IsFoundKey=0 then
			'GetKeyWord= "无此关键字"	
			'当无查询不到关键字时，调用关注回复,或者小鸡问答插件
		end if
	end function
	
	
	function GetPlug(p_Name,k_plugParam)
		SQL2 = "select p_Type,p_Api from [Plug] where p_IsLock='开启' and p_Name='"&p_Name&"'"
		Set Rs2 = Server.CreateObject("Adodb.RecordSet")
		Rs2.Open SQL2,Conn,1,3
		If Not (Rs2.Eof And Rs2.Bof) Then
				'将需要传递的参数加入Session
				Session("FromUserName")=FromUserName
				Session("ToUserName")=ToUserName
				Session("k_plugParam")=k_plugParam
				Session("Content")=Content
				Server.Execute(Rs2("p_Api"))	'执行调用插件的API
				Exit Function		
		End if
	end function
	
	'主动发送文本消息
	Function PostMsg(FromUserName,ToUserName,StrMsg)
		Access_token=GetToken()
		Sendtext="{""touser"":"""&ToUserName&""",""msgtype"":""text"",""text"":{""content"":"""&StrMsg&"""}}"
		strJson=PostURL(SendMsgUrl&"&access_token="&Access_token,Sendtext)
		Call InitScriptControl:Set objTest = getJSONObject(strJson)
		if objTest.errcode="0" then	
			'消息入库
			data=array(array("ToUserName",FromUserName,50,1),array("FromUserName",ToUserName,50,1),array("CreateTime",now(),50,1),array("MsgType","text",50,1),array("Content",StrMsg,0,1),array("BeiZhu","主动发送消息",0,1),array("Addtime",now(),50,1),array("reply","1",0,0))	
			Call Cls.db.dbnew("[sys_Msg]",data,"")
			PostMsg="1"
		else
			PostMsg="0发送失败"&objTest.errcode
		end if
	End Function
	
	'获取用户信息并入库
	Function GetUserInfo(id)
		Access_token=GetToken()
		strJson=GetURL(GetUserInfoUrl&"&access_token="&Access_token&"&openid="&id&"")
		Call InitScriptControl:Set objTest = getJSONObject(strJson)
		data=array(array("openid",objTest.openid,255,1),array("nickname",objTest.nickname,200,1),array("sex",objTest.sex,50,1),array("city",objTest.city,0,0),array("country",objTest.country,50,1),array("province",objTest.province,50,1),array("language",objTest.language,50,1),array("headimgurl",objTest.headimgurl,0,1),array("subscribe_time",FromUnixTime(objTest.subscribe_time),50,1))	
		call Cls.db.dbnew("[sys_User]",data,"FromUserName='"&FromUserName&"'")
	End Function
	
	'删除用户数据
	Function DelUser(id)
		Cls.db.dbdel "[sys_User]","openid='"&id&"'"
	End Function
	
	'获取最新Access_token
	Private function GetToken()
		dim datadb
		datadb=Cls.db.dbload("","AppId,Appsecret,Access_token,Token_Time,Expires_In","[sys_Config]","id=1","")	
		AppId=datadb(0,0)
		Appsecret=datadb(1,0)
		Access_token=datadb(2,0)
		Token_Time=datadb(3,0)
		Expires_In=datadb(4,0)
		GetToken=Access_token
		If datediff("s",Token_Time,Now())>Expires_In then '当Access_token过期时，重新获取新Access_token
			strJson=GetURL(GetTokenUtl&"grant_type=client_credential&appid="&AppId&"&secret="&Appsecret&"")
			if InStr(strJson,"errcode")>0 then GetToken="":exit function
			Call InitScriptControl:Set objTest = getJSONObject(strJson)
			Access_token=objTest.access_token	'获取新Access_token
			Expires_In=objTest.expires_in	
			'新的Access_token入库
			data=array(array("Access_token",Access_token,0,1),array("Expires_In",Expires_In,0,0),array("Token_Time",now(),50,1))	
			call Cls.db.dbupdate("[sys_Config]","id=1",data)
			GetToken=Access_token
		End If
	End function
	
	'Post内容，通过xml.http形式获取远程文件
	Function PostURL(url,PostStr)
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "POST", url, false ,"" ,""
			.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			.Send(PostStr)
			PostURL = .responsetext
		End With
		Set Retrieval = Nothing
		'response.Write PostURL
	End Function
	
	'Get内容，通过xml.http形式获取远程文件
	Function GetURL(url)	
		dim http
		set http=server.createobject("microsoft.xmlhttp")
			http.open "get",url,false
			http.setRequestHeader "If-Modified-Since","0"
			http.send()
			GetURL=http.responsetext
		set http=nothing
		'response.Write GetURL	
	End Function
	
	'时间戳转换成普通日期
	Function FromUnixTime(intTime) 
		If IsEmpty(intTime) Or Not IsNumeric(intTime) Then 
			FromUnixTime = Now() 
			Exit Function 
		End If 	
		FromUnixTime = DateAdd("s", intTime, "1970-1-1 0:0:0") 
		FromUnixTime = DateAdd("h", 8, FromUnixTime) 
	End Function
	
	'普通日期转换成时间戳
	Function ToUnixTime(strTime)        
		If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now   
		 ToUnixTime = DateAdd("h",-8,strTime)        
		 ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)        
	End Function  
	
	'解析json
	'Call InitScriptControl
	'Set objTest = getJSONObject(strTest)  
	Dim sc4Json   
	Sub InitScriptControl
		Set sc4Json = Server.CreateObject("MSScriptControl.ScriptControl")    
		sc4Json.Language = "JavaScript"    
		sc4Json.AddCode "var itemTemp=null;function getJSArray(arr, index){itemTemp=arr[index];}"    
	End Sub 
	Function getJSONObject(strJSON)    
		sc4Json.AddCode "var jsonObject = " & strJSON    
		Set getJSONObject = sc4Json.CodeObject.jsonObject    
	End Function 
	Sub getJSArrayItem(objDest,objJSArray,index)    
		On Error Resume Next    
		sc4Json.Run "getJSArray",objJSArray, index    
		Set objDest = sc4Json.CodeObject.itemTemp    
		If Err.number=0 Then Exit Sub    
		objDest = sc4Json.CodeObject.itemTemp    
	End Sub
	
	''转换HTML代码，过滤代码
	public function enhtml(byval t0)
		if isnull(t0) then enhtml="":exit function
		if t0="<p>&nbsp;</p>" then enhtml="":exit function
		t0=replace(t0,"&","&amp;")
		t0=replace(t0,"'","&#39;")
		t0=replace(t0,"""","&#34;")
		t0=replace(t0,"<","&lt;")
		t0=replace(t0,">","&gt;")
		set reg=new regexp
		reg.ignorecase=true
		reg.global=true
		reg.pattern="(w)(here)"
		t0=reg.replace(t0,"$1h&#101;re")
		reg.pattern="(s)(elect)"
		t0=reg.replace(t0,"$1el&#101;ct")
		reg.pattern="(i)(nsert)"
		t0=reg.replace(t0,"$1ns&#101;rt")
		reg.pattern="(c)(reate)"
		t0=reg.replace(t0,"$1r&#101;ate")
		reg.pattern="(d)(rop)"
		t0=reg.replace(t0,"$1ro&#112;")
		reg.pattern="(a)(lter)"
		t0=reg.replace(t0,"$1lt&#101;r")
		reg.pattern="(d)(elete)"
		t0=reg.replace(t0,"$1el&#101;te")
		reg.pattern="(u)(pdate)"
		t0=reg.replace(t0,"$1p&#100;ate")
		reg.pattern="(\s)(or)"
		t0=reg.replace(t0,"$1o&#114;")
		reg.pattern="(java)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		reg.pattern="(j)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		reg.pattern="(vb)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		if instr(t0,"expression")<>0 then
			t0=replace(t0,"expression","e&#173;xpression",1,-1,0)
		end if
		enhtml=t0
	end function
	
	''逆向转换HTML
	function dehtml(byval t0)
		if isnull(t0) then
			dehtml=""
			exit function
		end if
		t0=replace(t0,"&amp;","&")
		t0=replace(t0,"&#39;","'")
		t0=replace(t0,"&#34;","""")
		t0=replace(t0,"&lt;","<")
		t0=replace(t0,"&gt;",">")
		t0=replace(t0,chr(10),vbcrlf)
		dehtml=t0
	end function
%>
