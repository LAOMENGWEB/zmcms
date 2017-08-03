<!-- #Include File="../inc/conn.asp" -->
<%
Function addlog(logcontent)

        strsql="INSERT INTO [AdminLog] (userName,glyName,LoginIP,LoginSoft,LoginTime,logcontent) VALUES ('"&session("username")&"','"&Session("glyname")&"','"&Request.ServerVariables("Remote_Addr")&"','"&Request.ServerVariables("Http_USER_AGENT")&"','"&now()&"','"&logcontent&"')"
		conn.execute strsql
End Function

sub mxlb(DateBase,ClassId)
	Sqlp ="select * from "&DateBase&" Where ParentClassID = 0 and lang='"&lang&"' order by px"   
    Set Rsp=server.CreateObject("adodb.recordset")   
    rsp.open sqlp,conn,1,1 
    Response.Write("<option value="""">请选择分类</option>") 
    If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
    Else
      Do while not Rsp.Eof   
         Response.Write("<option")
		 If ClassId=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
				Response.Write(" value=" &Rsp("ID")& " style='color:#0000ff;'")
         		Response.Write(">|-" & Rsp("ClassName") & "</option>") & VbCrLf
		 
		    Sqlpp ="select * from "&DateBase&" Where ParentClassID="&Rsp("ID")&" and lang='"&lang&"' order by px"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option")
				If ClassId=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(" value="""&Rspp("ID")&""" style='color:#0000ff;'>　|-" & Rspp("ClassName") & "")
				Response.Write("</option>" )  & VbCrLf
				
			Sqlppp ="select * from "&DateBase&" Where ParentClassID="&Rspp("ID")&" and lang='"&lang&"' order by px"   
   			Set Rsppp=server.CreateObject("adodb.recordset")   
   			rsppp.open sqlppp,conn,1,1
			Do while not Rsppp.Eof 
				Response.Write("<option")
				If ClassId=Rsppp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(" value="""&Rsppp("ID")&""" style='color:#0000ff;'>　　　|-" & Rsppp("ClassName") & "")
				Response.Write("</option>" )  & VbCrLf	
				Rsppp.Movenext   
      		Loop
			Rspp.Movenext   
      		Loop
      Rsp.Movenext   
      Loop   
   End if
   rsp.close 
   rspp.close
   set rsp=nothing   
   set rspp=nothing 
end sub 

sub showColumn(DateBase,tid,id)
    if id<>"" then
	set rsObj1=server.CreateObject("adodb.recordset")
	Sqlp ="select * from "&DateBase&" Where id="&id&""     
    rsObj1.open sqlp,conn,1,1 
	end if 
	
	set rsObj=server.CreateObject("adodb.recordset")
	
	'Sqlp1 ="select * from Mycolumn where TypeID="&tid&" and lang='"&lang&"' order by Sequence"
	Sqlp1 ="select * from Mycolumn where TypeID="&tid&"  order by Sequence"
	rsObj.open sqlp1,conn,1,1
	if not rsObj.eof then
	do while not rsObj.eof
	if id="" then vstr="" else vstr=rsObj1(""&rsObj("ColumnName")&"")
	
	select case rsObj("ColumnType")
				case 0
					mystr=mystr&" <tr ><td>"&rsObj("ColumnInfo")&"：</td><td colspan=2><input type='text' name='"&rsObj("ColumnName")&"' value='"&vstr&"' /></td></tr>" 
				case 1
					mystr=mystr&"<tr ><td>"&rsObj("ColumnInfo")&"：</td><td colspan=2><input type='text' name='"&rsObj("ColumnName")&"' value='1'  style='width:80px;' /></td></tr>" 
				case 2
					mystr=mystr&"<script>var editor1;KindEditor.ready(function(K) {editor1 = K.create('textarea[id="""&rsObj("ColumnName")&"""]', {filterMode : false, uploadJson : '../zmbjq/asp/upload_json.asp',fileManagerJson : '../zmbjq/asp/file_manager_json.asp',allowFileManager : true,width: '700px',hight: '300px', afterBlur: function(){this.sync();}});});</script><tr ><td>"&rsObj("ColumnInfo")&"：</td><td colspan=2><textarea name='"&rsObj("Columnname")&"' id='"&rsObj("Columnname")&"' style=""width:700px;height:300px;visibility:hidden;"">"&vstr&"</textarea></div></div>" 
			

             case 3
					mystr=mystr&"<tr ><td>"&rsObj("ColumnInfo")&"：</td><td colspan=2><textarea name='"&rsObj("Columnname")&"'  style=""width:700px;height:300px;"">"&vstr&"</textarea></td></tr>" 
end Select



				
			rsObj.movenext
		loop
	end if
	rsObj.close
	set rsObj=nothing
	if id<>"" then
		rsObj1.close
		set rsObj1=nothing
	end if
	 response.Write( mystr)
    end sub 
	
	sub check_sort(a,b,c,d)
		sql="select id,classname,c_child from ["&d&"] where ParentClassID = "& a &" order by px asc"
		Set Rs1=Conn.Execute(sql)
		If Not Rs1.eof Then
			sortdata = rs1.getrows()
			Issortdata = 1
		End If
		Rs1.close : Set Rs1 = Nothing
		If Issortdata = 1 Then
			For iii = 0 To UBound(sortdata,2)
				If a = 0 Then
						Response.write("<option value='"& clng(sortdata(0,iii)) &"'")
				if instr(c,","& sortdata(0,iii) &",") > 0  Then Response.write("Selected")
						Response.write(">+ "& sortdata(1,iii) &"</option>")
				Else
						Response.write("<option value='"& clng(sortdata(0,iii)) &"'")
				If instr(c,","& sortdata(0,iii) &",") > 0 Then Response.write("Selected")
						Response.write(">"& tmp(b,"├ ") &" "& sortdata(1,iii) &"</option>")
				End If
				if sortdata(2,iii)>0 Then
					b=b+1
					check_sort sortdata(0,iii),b,c,d
					b=b-1
				End If
			Next
		End If
end Sub
'20150525 增加后台默认语言
set rslang=server.CreateObject("adodb.recordset")
rslang.open "select * from lang where ht=1 ",conn,1,1
lang=rslang("langbs")
langname=rslang("langname")
hbfhht=rslang("huobi")
hbwzht=rslang("huobi1")
rslang.close
Set rslang=Nothing

if lang<>"" then
set rswzset=server.CreateObject("adodb.recordset")
rswzset.open "select * from wzset where lang='"&lang&"'",conn,1,1
if rswzset.bof and rswzset.eof then response.end
yy=rswzset("yy")
newinfo=rswzset("newinfo")
ProductInfo=rswzset("ProductInfo")
alinfo=rswzset("alinfo")
downinfo=rswzset("downinfo")
tvinfo=rswzset("tvinfo")
zpinfo=rswzset("zpinfo")
lyinfo=rswzset("lyinfo")
wapnewinfo=rswzset("wapnewinfo")
wapProductInfo=rswzset("wapProductInfo")
wapalinfo=rswzset("wapalinfo")
wapdowninfo=rswzset("wapdowninfo")
waptvinfo=rswzset("waptvinfo")
wapzpinfo=rswzset("wapzpinfo")
waplyinfo=rswzset("waplyinfo")
skins=rswzset("skins")
rswzset.close
set rswzset=nothing
dataname1="tagstyle.asp"
SiteDataPath1 = SitePath&"template/"&skins&"/data/"&DataName1 
	on error resume next
	Connstr3="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(SiteDataPath1)
	Set Conn3=server.createobject("adodb.connection")
	Conn3.open Connstr3
    If Err Then
	err.Clear
	Set Conn3 = Nothing
	Response.End
	End If
end if

If lang="ch" then
Newclass="list"        '新闻分类前缀
newname="New"        '新闻详细页前缀
ProductClass="Productlist"        '产品分类前缀
ProductName="Product"        '产品详细页前缀
alClass="photolist"        '案例分类前缀
alName="photo"        '案例详细页前缀
downClass="downlist"        '下载分类前缀
downName="down"        '下载详细页前缀
tvClass="videolist"        '视频分类前缀
tvName="video"        '视频详细页前缀
zpClass="joblist"        '招聘列表前缀
zpName="job"        '招聘详细页前缀
aboutname="about"        '企业详细页前缀
guestname="showguest"        '显示留言页前缀
Separated="_"        '分隔符
Else
Newclass="list"&lang&""        '新闻分类前缀
newname="New"&lang&""        '新闻详细页前缀
ProductClass="Productlist"&lang&""        '产品分类前缀
ProductName="Product"&lang&""        '产品详细页前缀
alClass="photolist"&lang&""        '案例分类前缀
alName="photo"&lang&""        '案例详细页前缀
downClass="downlist"&lang&""        '下载分类前缀
downName="down"&lang&""        '下载详细页前缀
tvClass="videolist"&lang&""        '视频分类前缀
tvName="video"&lang&""        '视频详细页前缀
zpClass="joblist"&lang&""        '招聘列表前缀
zpName="job"&lang&""        '招聘详细页前缀
aboutname="about"&lang&""        '企业详细页前缀
guestname="showguest"&lang&""        '显示留言页前缀
Separated="_"        '分隔符

End if

%>
