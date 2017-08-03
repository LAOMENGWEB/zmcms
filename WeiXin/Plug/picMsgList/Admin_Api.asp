<!--#include file="../../Lib/base.asp"-->
<%
dim id:id=Cls.enhtml(Cls.fget("id",0))
dim act:act=Cls.enhtml(Cls.fget("act",0))
dim SqlStr,datadb
if cls.strlen(id)>0 then
	SqlStr="p_Id in ("&id&")"
else
	SqlStr=""
end if

if act="getlist" then
	datadb=Cls.db.dbload("","p_Id,p_Title","[Plug_PicMsg]",SqlStr,"p_Id desc")
	if ubound(datadb)>=0 then
		for i=0 to ubound(datadb,2)
		  cls.echo "<option value="""&datadb(0,i)&""">"&datadb(0,i)&"_"&datadb(1,i)&"</option>"
		next
	end if
else
	id=Cint(id)	'将字符串转化为整型
	datadb=Cls.db.dbload("","p_Id,p_Title","[Plug_PicMsgList]","","p_Id desc")
	if ubound(datadb)>=0 then
		for i=0 to ubound(datadb,2)
		  Cls.echo ("<option value="""&datadb(0,i)&"""")
		  if datadb(0,i)=id then cls.echo " selected"
		  Cls.echo (">"&datadb(0,i)&"_"&datadb(1,i)&"</option>")
		next
	else  
		  Cls.echo ("<option value="""">暂无多图文消息</option>  ")    
	end if
end if
%>