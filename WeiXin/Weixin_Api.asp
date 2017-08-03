<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Config.asp"-->
<!--#include file="lib/sha1.asp"-->
<!--#include file="lib/weixin_class.asp"-->
<%
	session.codepage=65001
	response.charset="utf-8"
	server.scripttimeout=999999
	dim startime:startime=timer()
		
	dim xml_dom,StrSend
	dim ToUserName,FromUserName,CreateTime,MsgType,MsgId
	dim Content,Recognition,MediaId,PicUrl,Format,ThumbMediaId,Location_X,Location_Y,Scale,Label,Title,Descriptions,Url,EventKey,Latitude,Longitude,Precision
	dim BeiZhu
	dim Connstr,Conn,Rs
	
	'定义数据库Conn连接
	on error resume next
	If datatype=true then  '判断是否access
		Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(webroot&datapath&"/"&datadbname)
	Else
		Connstr="Provider=Sqloledb;User ID="&datauser&";Password="&datapass&";Initial CataLog="&datasqlname&";Data Source="&datahost&";"
	End if
	Set Conn=server.createobject("adodb.connection")
	Conn.open Connstr
	
	
	Dim SignaTure, Nonce, TimesTamp, EchoStr
	SignaTure = Request("signature")
	Nonce     = Request("nonce")
	TimesTamp = Request("timestamp")
	EchoStr   = Request("echostr")		
	
	'验证微信接口
	If EchoStr<>"" then
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Sql = "select Token from [sys_Config] where id=1"
		Rs.Open Sql,Conn,1,3
		If Not rs.EOF = True Then
			Token=Rs("Token")	'获取系统中设置的Token
			Session("Token")=Token
		End If	
		Rs.Close
		Set Rs=Nothing
		
		'下面进行Token,TimesTamp,Nonce三个参数的字典排序
		dim str,M
		dim Myarray:Myarray=Sort(Array(Token,TimesTamp,Nonce))
		For M=0 To Ubound(Myarray)
			str=str&Myarray(M)
		Next
		if SignaTure=Lcase(sha1(str)) then
			response.Write EchoStr	'验证成功，返回正确EchoStr给微信，接通接口API
			response.End
		end if
	End if
	
	'接收微信发过来的消息并返回消息到微信，微信与微易程序交互的函数入口
	call GetXmlData()
	response.Write StrSend
	CloseConn()
	
	'获取微信主动发送过来的内容
	Sub GetXmlData
		set xml_dom=Server.CreateObject("MSXML2.DOMDocument")
		if xml_dom.load(request)=false then
			Response.End()	'判断是否由微信Post正确的XML数据过来
		else
			ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text '接收者微信账号。即我们的公众平台账号。
			FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text '发送者微信账号Openid，此Openid是由微信自己加密的，无法获取微信用户的真实微信号，而且同一个微信用户，对不同的公众号，生成的OpenId也是不同的，具有全网唯一值，可用来绑定其它会员系统
			CreateTime=xml_dom.getelementsbytagname("CreateTime").item(0).text
			MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text
			if (MsgType="event") then
				strEventType=xml_dom.getelementsbytagname("Event").item(0).text '微信事件
				if strEventType="subscribe" then '表示订阅微信公众平台
						BeiZhu="用户关注"
						Call AddMsg()
						Content="关注回复"
						StrSend=GetKeyWord(Content)
				ElseIf strEventType="unsubscribe" Then'取消关注
						BeiZhu="取消关注"
						Call AddMsg()
				ElseIf strEventType="CLICK" Then'点击菜单获取关键字，获取
						EventKey=xml_dom.getelementsbytagname("EventKey").item(0).text
						Content=EventKey
						BeiZhu="点击菜单获取关键字"
						Call AddMsg()
						StrSend=GetKeyWord(EventKey)
				ElseIf strEventType="VIEW" Then'点击菜单获取关键字，跳转到链接
						EventKey=xml_dom.getelementsbytagname("EventKey").item(0).text
						Content=EventKey
						BeiZhu="点击菜单跳转网址"
						Call AddMsg()
				ElseIf strEventType="LOCATION" Then'获取用户地理位置，当用户打开对话框时，自动获取微信用户的实时地址。本功能需要配合服务号的LEB接口。
						Latitude=xml_dom.getelementsbytagname("Latitude").item(0).text
						Longitude=xml_dom.getelementsbytagname("Longitude").item(0).text
						Precision=xml_dom.getelementsbytagname("Precision").item(0).text
						'记录用户LEB信息
						SQL = "select * from [sys_User_Location]"
						Set Rs = Server.CreateObject("Adodb.RecordSet")
						Rs.Open SQL,Conn,1,3
						Rs.Addnew
						Rs("OpenId")		= FromUserName
						Rs("Latitude")		= Latitude
						Rs("Longitude")		= Longitude
						Rs("Precision") 	= Precision
						Rs("AddTime")    	= Now()
						Rs.Update
						Rs.Close
						Set Rs = Nothing					
				end if
			else
				MsgId=xml_dom.getelementsbytagname("MsgId").item(0).text
			End If
			If MsgType="text" then
				Content=xml_dom.getelementsbytagname("Content").item(0).text
				'StrSend=RequestSendText(FromUserName,ToUserName,"你这发的是文本消息啊："&Content)
				BeiZhu="接收文本信息"
				Call AddMsg()
				StrSend=GetKeyWord(Content)
			elseif MsgType="image" then
				MediaId=xml_dom.getelementsbytagname("MediaId").item(0).text
				PicUrl=xml_dom.getelementsbytagname("PicUrl").item(0).text
				'StrSend=RequestSendText(FromUserName,ToUserName,"你发的是图片？"&PicUrl)
				Content=PicUrl
				BeiZhu="接收图片信息"
				Call AddMsg()
				StrSend=GetRule(BeiZhu)
			elseif MsgType="voice" then
				MediaId=xml_dom.getelementsbytagname("MediaId").item(0).text
				Format=xml_dom.getelementsbytagname("Format").item(0).text
				Recognition=xml_dom.getelementsbytagname("Recognition").item(0).text
				'StrSend=RequestSendText(FromUserName,ToUserName,"你发的是声音？")
				BeiZhu="接收语音信息"
				Call AddMsg()
				if len(Trim(Recognition))>0 then
					StrSend=GetKeyWord(Recognition)
				else
					StrSend=GetRule(BeiZhu)
				end if
			elseif MsgType="video" then
				MediaId=xml_dom.getelementsbytagname("MediaId").item(0).text
				ThumbMediaId=xml_dom.getelementsbytagname("ThumbMediaId").item(0).text
				'StrSend=RequestSendText(FromUserName,ToUserName,"你发的是视频？")
				BeiZhu="接收视频信息"
				Call AddMsg()
				StrSend=GetRule(BeiZhu)
			elseif MsgType="location" then
				Location_X=xml_dom.getelementsbytagname("Location_X").item(0).text
				Location_Y=xml_dom.getelementsbytagname("Location_Y").item(0).text
				Scale=xml_dom.getelementsbytagname("Scale").item(0).text
				Label=xml_dom.getelementsbytagname("Label").item(0).text
				'StrSend=RequestSendText(FromUserName,ToUserName,"你发的是地址信息？"&Label)
				BeiZhu="接收位置信息"
				Call AddMsg()
				StrSend=GetRule(BeiZhu)
			elseif MsgType="link" then
				Title=xml_dom.getelementsbytagname("Title").item(0).text
				Descriptions=xml_dom.getelementsbytagname("Description").item(0).text
				Url=xml_dom.getelementsbytagname("Url").item(0).text
				'StrSend=RequestSendText(FromUserName,ToUserName,"你发的是链接？")
				Content=Url
				BeiZhu="接收链接信息"
				Call AddMsg()
				StrSend=GetRule(BeiZhu)
			end if	
			set xml_dom=Nothing
		end if				
	End Sub
	
	Sub AddMsg()'消息入库
		Dim StrSQL '该条件解决同一消息多次入库情况，微信的机制是当用户发消息到公众平台时，微信Post标准的XML到微易，如果微信在五秒内收不到微易的响应会断掉连接，并且重新发起请求，总共重试三次，为了避免多次消息入库，加入MsgId的判断
		If MsgId = "" Then
			StrSQL = "1 = 1" 
		Else 
			StrSQL = "MsgId = '"& MsgId &"'"
		End If
		SQL = "select top 1 * from [sys_Msg] where "& StrSQL
		Set Rs = Server.CreateObject("Adodb.RecordSet")
		Rs.Open SQL,Conn,1,3
		If MsgId <> "" And Not (Rs.BOF And Rs.EOF) Then '所入库消息存在，直接退出
			Rs.Close
			Set Rs = nothing
			Exit Sub
		End if
		Rs.addnew
		Rs("ToUserName")   = ToUserName
		Rs("FromUserName") = FromUserName
		Rs("CreateTime")   = FromUnixTime(CreateTime)
		Rs("MsgType")      = MsgType
		Rs("MsgId")        = MsgId
		if len(trim(Recognition))>0 then
			Rs("Content")	= Recognition
		else
			Rs("Content")      = Content
		end if 
		Rs("MediaId")      = MediaId
		Rs("PicUrl")       = PicUrl
		Rs("Format")       = Format
		Rs("ThumbMediaId") = ThumbMediaId
		Rs("Location_X")   = Location_X
		Rs("Location_Y")   = Location_Y
		Rs("Scale")        = Scale
		Rs("Label")        = Label
		Rs("Title")        = Title
		Rs("Url")          = Url
		Rs("BeiZhu")       = BeiZhu
		Rs("AddTime")      = Now()
		Rs.Update
		Rs.Close
		Set Rs = Nothing
	End Sub
		
	'字典排序
	Function Sort(ary)
		Dim KeepChecking,I,FirstValue,SecondValue
		KeepChecking = TRUE 
		Do Until KeepChecking = FALSE 
			KeepChecking = FALSE 
			For I = 0 to UBound(ary) 
				If I = UBound(ary) Then Exit For 
				If ary(I) > ary(I+1) Then 
					FirstValue = ary(I) 
					SecondValue = ary(I+1) 
					ary(I) = SecondValue 
					ary(I+1) = FirstValue 
					KeepChecking = TRUE 
				End If 
			Next 
		Loop 
		Sort = ary 
	End Function 
		
	Sub CloseConn()
		Conn.close
		Set Conn=nothing
	End Sub
	
%>