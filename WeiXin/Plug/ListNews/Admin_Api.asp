<!--#include file="../base.asp"-->
<%

dim id:id=enhtml(fget("id",0))
dim act:act=enhtml(fget("act",0))
dim tablename:tablename="news"

dim Rs,Sql
dim Rs2,Sql2

dim SqlStr,datadb,oderStr
if len(id)>0 then
	SqlStr="and Id in ("&id&")"
	if sqlno=1 then
	oderStr=" order by instr('"&id&",',','&id&',')"
	else
	oderStr=" order by charindex(','+rtrim(cast(id as varchar(10)))+',','"&id&",')"
	end if
	 
else
	SqlStr=""
	oderStr=" order by id desc"
end if

if act="getlist" then	
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="select id,title from ["&tablename&"] where id>0 "&SqlStr&oderStr
	Set Rs = Conn.Execute(Sql)
	If not Rs.eof then
		Do While Not Rs.Eof
			response.Write "<option value="""&Rs("id")&""">"&Rs("id")&"_"&Rs("Title")&"</option>"
		Rs.MoveNext
		Loop
	end if
	Rs.Close
	Set Rs = Nothing
else
	id=Cint(id)	'将字符串转化为整型	
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Sql2="select p_Id,p_Title from [Plug_ListNews] order by p_Id desc"
	Set Rs2 = Conn2.Execute(Sql2)
	If not Rs2.eof then
		Do While Not Rs2.Eof
			response.Write "<option value="""&Rs2("p_Id")&""""
			if Rs2("p_Id")=id then response.Write " selected"
			response.Write ">"&Rs2("p_Id")&"_"&Rs2("p_Title")&"</option>"
		Rs2.MoveNext
		Loop
	else
		  response.Write ("<option value="""">暂无多图文消息</option>  ")    
	end if
	Rs2.Close
	Set Rs2 = Nothing
end if
%>