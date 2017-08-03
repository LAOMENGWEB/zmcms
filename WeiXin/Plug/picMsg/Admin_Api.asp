<!--#include file="../../Lib/base.asp"-->
<%
	dim id:id=Cls.getint(Cls.fget("id",0),0)
	dim datadb:datadb=Cls.db.dbload("","p_Id,p_Title","[Plug_PicMsg]","","p_Id desc")
	if ubound(datadb)>=0 then
		for i=0 to ubound(datadb,2)
		  Cls.echo ("<option value="""&datadb(0,i)&"""")
		  if datadb(0,i)=id then cls.echo " selected"
		  Cls.echo (">"&datadb(0,i)&"_"&datadb(1,i)&"</option>")
		next
	else  
		  Cls.echo ("<option value="""">暂无图文消息</option>  ")    
	end if
%>