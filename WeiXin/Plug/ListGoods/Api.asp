<!--#include file="../base.asp"-->
<%
'插件将会传过来下面四个值，请根据值进行操作Session("FromUserName")|Session("ToUserName")|Session("k_plugParam")|Session("Content")
'插件API需输出标准的微信XML

'本接口为多图文接口，可返回XML格式的多图文给微信							
dim k_plugParam:k_plugParam=Session("k_plugParam")
dim FromUserName:FromUserName=Session("FromUserName")
dim ToUserName:ToUserName=Session("ToUserName")
dim Title,Descriptions,PicUrl,Url,Api_Token
dim Rs,Sql
dim Rs2,Sql2
dim tablename:tablename="dy"		'''''''''''''''''''''''''''
dim plugpath:plugpath="ListGoods"		'''''''''''''''''''''''''''

'Api_Token用于本插件与其它网站插件进行校对的，可用系统的Session("Token"),也可自己设置一个Token
Api_Token=Session("Token")
'Api_Token="asdfw235lasf3525"

if k_plugParam<>0 then

	Set Rs2 = Server.CreateObject("Adodb.RecordSet")
	Sql2="select p_IdList from [Plug_"&plugpath&"] where p_Id="&k_plugParam
	Set Rs2 = Conn2.Execute(Sql2)
	If not Rs2.eof then
			if sqlno=1 then
		Sql="select * from ["&tablename&"] where id in("&rs2("p_IdList")&") order by instr('"&rs2("p_IdList")&",',','&id&',')"
		else
		Sql="select * from ["&tablename&"] where id in("&rs2("p_IdList")&") order by charindex(','+rtrim(cast(id as varchar(10)))+',','"&rs2("p_IdList")&",')"
		
		end if
		Set Rs = Server.CreateObject("Adodb.RecordSet")
		Rs.Open Sql,Conn,1,3
		If not Rs.eof then
			dim t:t="<xml>"
			t=t&"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>"
			t=t&"<FromUserName><![CDATA["&tousername&"]]></FromUserName>"
			t=t&"<CreateTime>"&now&"</CreateTime>"
			t=t&"<MsgType><![CDATA[news]]></MsgType>"
			t=t&"<ArticleCount>"&rs.recordcount&"</ArticleCount>"
			t=t&"<Articles>"
			Do While Not Rs.Eof
				Title=rs("Title")
				Descriptions=CutStrX(nohtml1(rs("content")),200)
				if rs("smallimage")<>"" then
				picurl=weburl&picroot&rs("smallimage")
					else
				
				picurl1="images/wu.gif"
				picurl=weburl&picroot&picurl1
				
				
				end if
				
				
				 set rszm=server.CreateObject("adodb.recordset")
                rszm.open "select * from mnclass where id="&rs("zmone")&" ",conn,1,1
                mm=rszm("m")
                rszm.close
                Set rszm=Nothing
				
				
				
				if rs("wlurl")="" then
				if kq=1 then
				Url=weburl&picroot&"wap/showabout/?pone="&rs("zmone")&"&aboutID="&rs("id")&"&m="&mm&"&language="&rs("lang")
				else
				Url=weburl&webroot&"plug/"&plugpath&"/show.asp?id="&rs("id")
				end if
				else
				Url=rs("wlurl")
				end if
				
				
				
				
				
				t=t&"<item>"
				t=t&"<Title><![CDATA["&Title&"]]></Title>"
			
					t=t&"<Description><![CDATA["&Descriptions&"]]></Description>"
		
		
					t=t&"<PicUrl><![CDATA["&PicUrl&"]]></PicUrl>"
		
				t=t&"<Url><![CDATA["&Url&"]]></Url>"
				t=t&"</item>"
			Rs.MoveNext
			Loop
			t=t&"</Articles>"
			t=t&"</xml>"
			response.Write t
		end if
		Rs.Close
		Set Rs = Nothing
	end if
	Rs2.Close
	Set Rs2 = Nothing
end if
%>