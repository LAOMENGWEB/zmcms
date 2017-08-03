<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="sqlzr.asp"-->
<!--#include file="MD5.asp"-->
<!--#include file="page.asp"-->
<%
if yy = 2 then
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
End If
value = replace(Replace(Trim(request("value")),"'","&#180;")," ","")
'**************************************************
'函 数 名：logoimg
'作    用：显示网站logo
'*************************************************
function logoimg
logoimg=loginimg&"<a href="""&ym&""" title="""&mc&""" ><img  src="""&sitepath&""&logo&""" width="""&logowidth&""" height="""&logoheight&""" alt="""&mc&""" title="""&mc&""" /></a>"
end function
'**************************************************
'函 数 名：waplogoimg
'作    用：显示wap网站logo
'*************************************************
function waplogoimg
waplogoimg=waplogoimg&"<a href="""&ym&"/wap/?language="&language&""" title="""&mc&"""><img  src="""&sitepath&""&waplogo&""" width="""&waplogowidth&""" height="""&waplogoheight&""" alt="""&mc&""" title="""&mc&""" /></a>"
end function
'**************************************************
'函 数 名：erwm
'作    用：调用二维码
'*************************************************
function erwm
erwm=erwm&"<img  src="""&sitepath&""&ewm&""" width="""&ewmwidth&""" height="""&ewmheight&""" alt="""&mc&""" title="""&mc&""" />"
end function
'**************************************************
'函 数 名：wzgb
'作    用：网站是否关闭按钮
'*************************************************
function wzgb
if wzkg=1 then  
response.Redirect ""&sitepath&"error/?language="&language&"" 
end if
end function
'**************************************************
'函 数 名：share
'作    用：显示百度分享
'*************************************************
function share
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from share where id=1"
rs.Open sql, conn, 1, 1
if rs.bof and rs.eof then response.end
img=rs("img")
wz2=rs("wz")
top=rs("top")
jl=rs("jl")
if jt=1 then
if jl=1 then
share=share&"<script type=""text/javascript"" id=""bdshare_js"" data=""type=slide&amp;img="&img&"&amp;mini=1&amp;pos="&wz2&"&amp;uid=0"" ></script>"
else
share=share&"<script type=""text/javascript"" id=""bdshare_js"" data=""type=slide&amp;img="&img&"&amp;pos="&wz2&"&amp;uid=0"" ></script>"
end if
share=share&"<script type=""text/javascript"" id=""bdshell_js""></script>"
share=share&"<script type=""text/javascript"">"
share=share&"var bds_config={""bdTop"":"&top&"};"
share=share&"document.getElementById(""bdshell_js"").src = ""http://bdimg.share.baidu.com/static/js/shell_v2.js?cdnversion="" + Math.ceil(new Date()/3600000);"
share=share&"</script>"
else
end if
rs.close
set rs=nothing 
end function
'**************************************************
'函 数 名：dlgg
'作    用：显示对联广告
'*************************************************
function dlgg
if dlad=1 then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ad where Lang='"&Language&"'  order by id desc",conn,1,1
if rs.bof and rs.eof then response.end
Width=rs("Width")
Height=rs("Height")
pic=rs("pic")
Url=rs("url")
pic1=rs("pic1")
Url1=rs("url1")
top=rs("top")
left1=rs("left1")
right1=rs("right1")
lf=rs("lf")
rf=rs("rf")
%>
<script language="javascript" type="text/javascript">
lastScrollY=0;
function heartBeat(){
var diffY;
if (document.documentElement && document.documentElement.scrollTop)
 diffY = document.documentElement.scrollTop;
else if (document.body)
 diffY = document.body.scrollTop
else
    {/*Netscape stuff*/}
//alert(diffY);
percent=.1*(diffY-lastScrollY);
if(percent>0)percent=Math.ceil(percent);
else percent=Math.floor(percent);
document.getElementById("left").style.top=parseInt(document.getElementById
("left").style.top)+percent+"px";
document.getElementById("right").style.top=parseInt(document.getElementById
("left").style.top)+percent+"px";
lastScrollY=lastScrollY+percent;
//alert(lastScrollY);
}
j1="<div id=\"left\" style='overflow: hidden;left:<%=left1%>px;position:absolute;top:<%=top%>px;width:<%=width%>px; z-index:999;display:<%=lf%>'><span style='display:inline-block;float:right;cursor:pointer;' onclick='this.parentNode.style.display=\"none\"'><img src='<%=SitePath%>images/close.gif'/></span><a href='<%=Url%>' target='_blank'><img src='<%=sitepath%><%=pic%>' width='<%=width%>'height='<%=height%>'/></a></div>"
j2="<div id=\"right\"style='overflow: hidden;right:<%=right1%>px;position:absolute;top:<%=top%>px;width:<%=width%>px;z-index:999;display:<%=rf%>'><span style='display:inline-block;float:right;cursor:pointer' onclick='this.parentNode.style.display=\"none\"'><img src='<%=SitePath%>images/close.gif'/></span><a href='<%=Url1%>' target='_blank'><img src='<%=sitepath%><%=pic1%>' width='<%=width%>'height='<%=height%>'/></a></div>"
document.write(j1);
document.write(j2);
window.setInterval("heartBeat()",1);
</script>
<% 

rs.close
set rs=nothing
end if
end function 
'**************************************************
'函 数 名：pfgg
'作    用：显示漂浮广告
'*************************************************
function pfgg
if pfad=1 then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ad where Lang='"&Language&"'  order by id desc",conn,1,1
if rs.bof and rs.eof then response.end
pfWidth=rs("pfWidth")
pfHeight=rs("pfHeight")
pfpic=rs("pfpic")
pfUrl=rs("pfurl")
pfgg=pfgg&"<script type=""text/javascript"">" & vbCrLf
pfgg=pfgg&"$(function(){" & vbCrLf
pfgg=pfgg&"$(""#closepiaofu"").click(function () {" & vbCrLf
pfgg=pfgg&"$(""#wzsse"").hide();" & vbCrLf
pfgg=pfgg&"});" & vbCrLf
pfgg=pfgg&"})" & vbCrLf
pfgg=pfgg&"</script>" & vbCrLf
pfgg=pfgg&"<div id=""wzsse"" style=""position:absolute;z-index:999"">"& vbCrLf
pfgg=pfgg&"<div style=""text-align:right;cursor:pointer;"" id=""closepiaofu"">关闭</div>"& vbCrLf
pfgg=pfgg&"<a href="""&pfUrl&""" target=""_blank""><img src="""&sitepath&""&pfpic&""" border=""0"" width="""&pfwidth&""" height="""&pfheight&""" /></a>"& vbCrLf
pfgg=pfgg&"</div>"& vbCrLf
pfgg=pfgg&"<script type=""text/javascript"" src="""&SitePath&"js/floatadv.js""></script>"& vbCrLf
pfgg=pfgg&"<script type=""text/javascript"">"& vbCrLf
pfgg=pfgg&"var ad1 = new AdMove(""wzsse"", true);"& vbCrLf
pfgg=pfgg&"ad1.Run();"& vbCrLf
pfgg=pfgg&"</script>"& vbCrLf
rs.close
set rs=nothing 
end if
end function
'**************************************************
'函 数 名：模板条件语句!只针对于调用栏目使用，其他无效！
'*************************************************
function ireplace(str)
dim re, tmatch, tmpstr, tstr, m
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(str)
For Each m in tmatch 
    str = replace(str, m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
ireplace = str
end function
function Replace_tpl_If(ta, tb, tc)
dim ValueIf: ValueIf = replace(ta, "'", "")
if Eval(ValueIf) Then
Replace_tpl_If = tb
else
Replace_tpl_If = tc
end if
end function
'**************************************************
'函 数 名：Memu 循环显示顶部，底部，其他导航代码
'*************************************************
'-----一级导航循环
function Memu(str)
set reg=new regexp
reg.IgnoreCase = true 
reg.Global = True 
reg.Pattern = "<Memu:\{(.+?)\}>([\s\S]*?)<Memu>"
set aa=reg.execute(str)
for each match in aa
bq=match.submatches(0)
nr=match.submatches(1)
Row=nbq(bq,"Row")
LX=nbq(bq,"LX")
str=replace(str,match,Getmemu(Row,LX,nr))
next
Memu=str
end function
function Getmemu(Row,LX,zmcms)
  set rs = server.createobject("adodb.recordset")
  if LX=1 then
  sql="select top "&Row&" id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = 0 and xs= 1 and fl=1 and Lang='"&Language&"' order by px asc"
  end if
  if LX=2 then
  sql="select top "&Row&" id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = 0 and xs= 1 and fl=0 and Lang='"&Language&"'  order by px asc"
  end if
  if LX=3 then
  sql="select top "&Row&" id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = 0 and xs= 1 and fl=2 and Lang='"&Language&"' order by px asc"
  end if
    if LX=4 then
  sql="select top "&Row&" id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,bcolor,c_child,m,lang from meun where ParentClassID = 0 and xs= 1 and fl=3 and Lang='"&Language&"' order by px asc"
  end if
  rs.open sql,conn,1,1
  do while not rs.eof
  linenums = linenums + 1
if yy=1 then
zurl=rs(2)
else
zurl=rs(3)
end if
if zurl<>"" then
zurl1=Left(zurl,InstrRev(zurl,"//"))
end if
if zurl1<>"" then
zurl2=left(zurl1,(len(zurl1)-2))
end if
if lx<>4 then
if yy=1 then
if zurl2<>"http" and zurl2<>"https" then
lanmuurl=""&sitepath&""&rs(2)&""
else
lanmuurl=rs(2)
end if
else
if zurl2<>"http" and zurl2<>"https" then
lanmuurl=""&sitepath&""&rs(3)&""
else
lanmuurl=rs(3)
end if
end if
Else

if zurl2<>"http" and zurl2<>"https" then	
lanmuurl=""&sitepath&"wap/"&rs(2)&""
Else
lanmuurl=rs(2)	
End if

end if
if rs(5)<>"" then
lanmuname="<font color="&rs(5)&">"&rs(1)&"</font>"
else
lanmuname=rs(1)
end if
if rs(7)<>"" then
fenleitp=""&sitepath&""&rs(7)&""
else
fenleitp=""&sitepath&"images/demo.jpg"
end if

if request.QueryString("m")<>"" then
rqm=request.QueryString("m")
else
rqm=0
end if

zmcms1=replace(zmcms,"{Memu:i}",linenums)
zmcms2=replace(zmcms1,"{Memu:id}",rs(0))
zmcms3=replace(zmcms2,"{Memu:name}",lanmuname)
zmcms4=replace(zmcms3,"{Memu:url}",lanmuurl)
zmcms5=replace(zmcms4,"{Memu:mopen}",rs(4))
zmcms6=replace(zmcms5,"{Memu:bieming}",rs(6))
zmcms7=replace(zmcms6,"{Memu:img}",fenleitp)
zmcms8=replace(zmcms7,"{Memu:child}",rs(8))
zmcms9=replace(zmcms8,"{Memu:lujing}",""&sitepath&"")
zmcms10=replace(zmcms9,"{Memu:count}",conn.execute("select count(*) from meun where ParentClassID = 0 and Lang='"&Language&"'")(0))
if lx=4 then
zmcms11=replace(zmcms10,"{Memu:m}",rs(10))
zmcms12=replace(zmcms11,"{Memu:lang}",rs(11))
else
zmcms11=replace(zmcms10,"{Memu:m}",rs(9))
zmcms12=replace(zmcms11,"{Memu:lang}",rs(10))
end if





zmcms13=replace(zmcms12,"{Memu:rqm}",rqm)
zmcms14=replace(zmcms13,"{Memu:addlang}",language)




if lx=4 then
if rs(8)<>"" then
zmcms15=replace(zmcms14,"{Memu:bcolor}",rs(8))
else
zmcms15=replace(zmcms14,"{Memu:bcolor}","")
end if

if rs(3)<>"" then
zmcms16=replace(zmcms15,"{Memu:ico}",rs(3))
else
zmcms16=replace(zmcms15,"{Memu:ico}","")
end if
end if
if lx=4 then
mm=mm+zmcms16
else
mm=mm+zmcms14
end if
'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if
rs.movenext
loop
Getmemu=Getmemu&mm
rs.Close
set rs=nothing
end Function

'-----二级导航循环
function Memu1(str)
	set reg=new regexp
	reg.IgnoreCase = True 
	reg.Global = True 
	reg.Pattern = "<Memu1:\{(.+?)\}>([\s\S]*?)<Memu1>"
	set aa=reg.execute(str)
	for each match in aa
	bq=match.submatches(0)
	nr=match.submatches(1)
	Id1=nbq(bq,"Id1")
	LX1=nbq(bq,"LX1")
	str=replace(str,match,Getmemu1(nr,LX1,Id1))
	next
	Memu1=str
end function
function Getmemu1(zmcms,LX1,Id1)
  set rss = server.createobject("adodb.recordset")
  if LX1=1 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = " & Id1 & " and xs= 1 and fl=1 and Lang='"&Language&"' order by px asc"
 end if
   if LX1=2 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = " & Id1 & " and xs= 1 and fl=0 and Lang='"&Language&"' order by px asc"
 end if
   if LX1=3 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = " & Id1 & " and xs= 1 and fl=2 and Lang='"&Language&"' order by px asc"
 end if
    if LX1=4 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,bcolor,c_child,m,lang from meun where ParentClassID = " & Id1 & " and xs= 1 and fl=3 and Lang='"&Language&"' order by px asc"
 end if
  rss.open sql,conn,1,1
do while not rss.eof
linenums = linenums + 1
if yy=1 then
zurl=rss(2)
else
zurl=rss(3)
end if
if zurl<>"" then
zurl1=Left(zurl,InstrRev(zurl,"//"))
end if
if zurl1<>"" then
zurl2=left(zurl1,(len(zurl1)-2))
end if

if lx1<>4 then
if yy=1 then
if zurl2<>"http" and zurl2<>"https" then
lanmuurl1=""&sitepath&""&rss(2)&""
else
lanmuurl1=rss(2)
end if
else
if zurl2<>"http" and zurl2<>"https" then
lanmuurl1=""&sitepath&""&rss(3)&""
else
lanmuurl1=rss(3)
end if
end if
else
if zurl2<>"http" and zurl2<>"https" then	
lanmuurl1=""&sitepath&"wap/"&rss(2)&""
Else
lanmuurl1=rss(2)	
End if



end if

if rss(5)<>"" then
lanmuname1="<font color="&rss(5)&">"&rss(1)&"</font>"
else
lanmuname1=rss(1)
end if

if rss(7)<>"" then
fenleitp=""&sitepath&""&rss(7)&""
else
fenleitp=""&sitepath&"images/demo.jpg"
end if


if request.QueryString("m")<>"" then
rqm=request.QueryString("m")
else
rqm=0
end if





zmcms1=replace(zmcms,"{Memu1:i}",linenums)
zmcms2=replace(zmcms1,"{Memu1:id}",rss(0))
zmcms3=replace(zmcms2,"{Memu1:name}",lanmuname1)
zmcms4=replace(zmcms3,"{Memu1:url}",lanmuurl1)
zmcms5=replace(zmcms4,"{Memu1:mopen}",rss(4))
zmcms6=replace(zmcms5,"{Memu1:bieming}",rss(6))
zmcms7=replace(zmcms6,"{Memu1:img}",fenleitp)
zmcms8=replace(zmcms7,"{Memu1:child}",rss(8))
zmcms9=replace(zmcms8,"{Memu1:lujing}",""&sitepath&"")
zmcms10=replace(zmcms9,"{Memu1:count}",conn.execute("select count(*) from meun where ParentClassID = " & Id1 & " and Lang='"&Language&"'")(0))
if lx1=4 then
zmcms11=replace(zmcms10,"{Memu1:m}",rss(10))
zmcms12=replace(zmcms11,"{Memu1:lang}",rss(11))
else
zmcms11=replace(zmcms10,"{Memu1:m}",rss(9))
zmcms12=replace(zmcms11,"{Memu1:lang}",rss(10))
end if
zmcms13=replace(zmcms12,"{Memu1:rqm}",rqm)
zmcms14=replace(zmcms13,"{Memu1:addlang}",language)

if lx1=4 then
if rss(8)<>"" then
zmcms15=replace(zmcms14,"{Memu1:bcolor}",rss(8))
else
zmcms15=replace(zmcms14,"{Memu1:bcolor}","")
end if

if rss(3)<>"" then
zmcms16=replace(zmcms15,"{Memu1:ico}",rss(3))
else
zmcms16=replace(zmcms15,"{Memu1:ico}","")
end if
end if
if lx1=4 then
mm=mm+zmcms16
else
mm=mm+zmcms14
end if




'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if
rss.movenext
loop
Getmemu1=Getmemu1&mm

rss.Close
set rss=nothing
end Function




'-----三级导航循环
function Memu2(str)
	set reg=new regexp
	reg.IgnoreCase = True 
	reg.Global = True 
	reg.Pattern = "<Memu2:\{(.+?)\}>([\s\S]*?)<Memu2>"
	set aa=reg.execute(str)
	for each match in aa
	bq=match.submatches(0)
	nr=match.submatches(1)
	Id2=nbq(bq,"Id2")
	LX2=nbq(bq,"LX2")
	str=replace(str,match,Getmemu2(nr,LX2,Id2))
	next
	Memu2=str
end function
function Getmemu2(zmcms,LX2,Id2)
  set rsss = server.createobject("adodb.recordset")
  if LX2=1 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = " & Id2 & " and xs= 1 and fl=1 and Lang='"&Language&"' order by px asc"
 end if
   if LX2=2 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = " & Id2 & " and xs= 1 and fl=0 and Lang='"&Language&"' order by px asc"
 end if
   if LX2=3 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,c_child,m,lang from meun where ParentClassID = " & Id2 & " and xs= 1 and fl=2 and Lang='"&Language&"' order by px asc"
 end if
   if LX2=4 then
  sql="select id,classname,url,htmlurl,mopen,mcolor,bieming,SmallImage,bcolor,c_child,m,lang from meun where ParentClassID = " & Id2 & " and xs= 1 and fl=3 and Lang='"&Language&"' order by px asc"
 end if
  rsss.open sql,conn,1,1
 
do while not rsss.eof
linenums = linenums + 1
if yy=1 then
zurl=rsss(2)
else
zurl=rsss(3)
end if
if zurl<>"" then
zurl1=Left(zurl,InstrRev(zurl,"//"))
end if
if zurl1<>"" then
zurl2=left(zurl1,(len(zurl1)-2))
end if
  
if lx2<>4 then
if yy=1 then
if zurl2<>"http" and zurl2<>"https" then
lanmuurl2=""&sitepath&""&rsss(2)&""
else
lanmuurl2=rsss(2)
end if
else
if zurl2<>"http" and zurl2<>"https" then
lanmuurl2=""&sitepath&""&rsss(3)&""
else
lanmuurl2=rsss(3)
end if
end if
else
if zurl2<>"http" and zurl2<>"https" then	
lanmuurl2=""&sitepath&"wap/"&rsss(2)&""
Else
lanmuurl2=rsss(2)	
End if
end if


if rsss(5)<>"" then
lanmuname2="<font color="&rsss(5)&">"&rsss(1)&"</font>"
else
lanmuname2=rsss(1)
end if


if rsss(7)<>"" then
fenleitp=""&sitepath&""&rsss(7)&""
else
fenleitp=""&sitepath&"images/demo.jpg"
end if

if request.QueryString("m")<>"" then
rqm=request.QueryString("m")
else
rqm=0
end if

zmcms1=replace(zmcms,"{Memu2:i}",linenums)
zmcms2=replace(zmcms1,"{Memu2:id}",rsss(0))
zmcms3=replace(zmcms2,"{Memu2:name}",lanmuname2)
zmcms4=replace(zmcms3,"{Memu2:url}",lanmuurl2)
zmcms5=replace(zmcms4,"{Memu2:mopen}",rsss(4))
zmcms6=replace(zmcms5,"{Memu2:bieming}",rsss(6))
zmcms7=replace(zmcms6,"{Memu2:img}",fenleitp)
zmcms8=replace(zmcms7,"{Memu2:child}",rsss(8))
zmcms9=replace(zmcms8,"{Memu2:lujing}",""&sitepath&"")
zmcms10=replace(zmcms9,"{Memu2:count}",conn.execute("select count(*) from meun where ParentClassID = " & Id2 & " and Lang='"&Language&"'")(0))
if lx2=4 then
zmcms11=replace(zmcms10,"{Memu2:m}",rsss(10))
zmcms12=replace(zmcms11,"{Memu2:lang}",rsss(11))
else
zmcms11=replace(zmcms10,"{Memu2:m}",rsss(9))
zmcms12=replace(zmcms11,"{Memu2:lang}",rsss(10))
end if
zmcms13=replace(zmcms12,"{Memu2:rqm}",rqm)
zmcms14=replace(zmcms13,"{Memu2:addlang}",language)
if lx2=4 then
if rsss(8)<>"" then
zmcms15=replace(zmcms14,"{Memu2:bcolor}",rsss(8))
else
zmcms15=replace(zmcms14,"{Memu2:bcolor}","")
end if

if rsss(3)<>"" then
zmcms16=replace(zmcms15,"{Memu2:ico}",rsss(3))
else
zmcms16=replace(zmcms15,"{Memu2:ico}","")
end if
end if
if lx2=4 then
mm=mm+zmcms16
else
mm=mm+zmcms14
end if

'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if
rsss.movenext
loop
Getmemu2=Getmemu2&mm

rsss.Close
set rsss=nothing
end Function





'**************************************************
'函 数 名：lanmu 循环显示模块分类代码
'*************************************************
'-----一级栏目循环
function lanmu(str)
set reg=new regexp
reg.IgnoreCase = True 
reg.Global = True 
reg.Pattern = "<lanmu:\{(.+?)\}>([\s\S]*?)<lanmu>"
set aa=reg.execute(str)
for each match in aa
bq=match.submatches(0)
nr=match.submatches(1)
Row=nbq(bq,"Row")
LX=nbq(bq,"LX")
str=replace(str,match,Getlanmu(Row,LX,nr))
next
lanmu=str
end function

function Getlanmu(Row,LX,zmcms)
  set rs = server.createobject("adodb.recordset")
  if LX=1 then
  sql="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from newsclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX=2 then
  sql="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from cpclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX=3 then
  sql="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from alclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX=4 then
  sql="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from spclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX=5 then
  sql="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from xzclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  rs.open sql,conn,1,1
  do while not rs.eof
  linenums = linenums + 1
  if LX=1 THEN
  lanmuurl1=""&SitePath&"wap/info/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=2 THEN
  lanmuurl1=""&SitePath&"wap/product/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=3 THEN
  lanmuurl1=""&SitePath&"wap/photo/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=4 THEN
  lanmuurl1=""&SitePath&"wap/video/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=5 THEN
  lanmuurl1=""&SitePath&"wap/down/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  if yy=1 then
  if LX=1 THEN
  lanmuurl=""&SitePath&""&iinfo&"/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=2 THEN
  lanmuurl=""&SitePath&""&product&"/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=3 THEN
  lanmuurl=""&SitePath&""&photo&"/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=4 THEN
  lanmuurl=""&SitePath&""&video&"/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  IF LX=5 THEN
  lanmuurl=""&SitePath&""&down&"/?pone="&rs(0)&"&m="&rs(5)&"&Language="&rs(6)&""
  END IF
  ELSE
  IF LX=1 THEN
  lanmuurl=""&SitePath&""&rs(4)&"/"&newClass&""&Separated&""&rs("id")&""&Separated&"1.html"
  END IF
  IF LX=2 THEN
  lanmuurl=""&SitePath&""&rs(4)&"/"&productClass&""&Separated&""&rs("id")&""&Separated&"1.html"
  END IF
  IF LX=3 THEN
  lanmuurl=""&SitePath&""&rs(4)&"/"&alClass&""&Separated&""&rs("id")&""&Separated&"1.html"
  END IF
  IF LX=4 THEN
  lanmuurl=""&SitePath&""&rs(4)&"/"&tvClass&""&Separated&""&rs("id")&""&Separated&"1.html"
  END IF
  IF LX=5 THEN
  lanmuurl=""&SitePath&""&rs(4)&"/"&downClass&""&Separated&""&rs("id")&""&Separated&"1.html"
  END IF
  END IF
  
  if lx=1 then
  countnum=Conn.Execute("Select count(*) from news where zmone="&rs(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx=2 then
  countnum=Conn.Execute("Select count(*) from product where zmone="&rs(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx=3 then
  countnum=Conn.Execute("Select count(*) from al where zmone="&rs(0)&" and Lang='"&Language&"' ")(0)
  end if
  
   if lx=4 then
  countnum=Conn.Execute("Select count(*) from zxsp where zmone="&rs(0)&" and Lang='"&Language&"' ")(0)
  end If
   if lx=5 then
  countnum=Conn.Execute("Select count(*) from download where zmone="&rs(0)&" and Lang='"&Language&"' ")(0)
  end If


  if lx=1 then
  count=Conn.Execute("Select count(*) from newsclass where ParentClassID = 0 and Lang='"&Language&"' ")(0)
  end if
  if lx=2 then
  count=Conn.Execute("Select count(*) from cpclass where ParentClassID = 0 and Lang='"&Language&"' ")(0)
  end if
  if lx=3 then
  count=Conn.Execute("Select count(*) from alclass where ParentClassID = 0 and Lang='"&Language&"' ")(0)
  end if
  
   if lx=4 then
  count=Conn.Execute("Select count(*) from spclass where ParentClassID = 0 and Lang='"&Language&"' ")(0)
  end If
   if lx=5 then
  count=Conn.Execute("Select count(*) from xzclass where ParentClassID = 0 and Lang='"&Language&"' ")(0)
  end If



  
  
  
  if rs(2)<>"" then
fenleitp=""&sitepath&""&rs(2)&""
else
fenleitp=""&sitepath&"images/demo.jpg"
end if

'更新栏目高亮 更新于20150507

if request.QueryString("pone")="" then
po=0
else
po=request.QueryString("pone")
end if
'更新栏目高亮 更新于20150507
  
  zmcms1=replace(zmcms,"{lanmu:i}",linenums)
  zmcms2=replace(zmcms1,"{lanmu:id}",rs(0))
  zmcms3=replace(zmcms2,"{lanmu:name}",rs(1))
  zmcms4=replace(zmcms3,"{lanmu:url}",lanmuurl)
  zmcms5=replace(zmcms4,"{lanmu:img}",fenleitp)
  zmcms6=replace(zmcms5,"{lanmu:wapurl}",lanmuurl1)
  zmcms7=replace(zmcms6,"{lanmu:child}",rs(3))
  zmcms8=replace(zmcms7,"{lanmu:num}",countnum)
  zmcms9=replace(zmcms8,"{lanmu:mulu}",rs(4))
  zmcms10=replace(zmcms9,"{lanmu:lujing}",""&sitepath&"")
  zmcms11=replace(zmcms10,"{lanmu:po}",po)
  zmcms12=replace(zmcms11,"{lanmu:lang}",rs(6))
  zmcms13=replace(zmcms12,"{lanmu:addlang}",language)
  zmcms14=replace(zmcms13,"{lanmu:count}",count)
  mm=mm+zmcms14
'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if
rs.movenext
loop
Getlanmu=Getlanmu&mm
rs.Close
set rs=nothing
end Function

'-----二级导航循环
function lanmu1(str)
	set reg=new regexp
	reg.IgnoreCase = True 
	reg.Global = True 
	reg.Pattern = "<lanmu1:\{(.+?)\}>([\s\S]*?)<lanmu1>"
	set aa=reg.execute(str)
	for each match in aa
	bq=match.submatches(0)
	nr=match.submatches(1)
	Id1=nbq(bq,"Id1")
	LX1=nbq(bq,"LX1")
	str=replace(str,match,Getlanmu1(nr,LX1,Id1))
	next
	lanmu1=str
end function
function Getlanmu1(zmcms,LX1,Id1)
  set rss = server.createobject("adodb.recordset")
  if LX1=1 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where ParentClassID = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX1=2 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where ParentClassID = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX1=3 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where ParentClassID = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX1=4 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where ParentClassID = "&Id1&" and xs= 1 and Lang='"&Language&"' order by px asc"
  end if
   if LX1=5 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where ParentClassID = "&Id1&" and xs= 1 and Lang='"&Language&"' order by px asc"
  end if

  rss.open sql,conn,1,1
 
  do while not rss.eof
  linenums = linenums + 1
  if LX1=1 THEN
  lanmuurl1=""&SitePath&"wap/info/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=2 THEN
  lanmuurl1=""&SitePath&"wap/product/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=3 THEN
  lanmuurl1=""&SitePath&"wap/photo/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=4 THEN
  lanmuurl1=""&SitePath&"wap/video/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=5 THEN
  lanmuurl1=""&SitePath&"wap/down/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  '获取上一级栏目的名称
   set rss1 = server.createobject("adodb.recordset")
  if LX1=1 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where id = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX1=2 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where id = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX1=3 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where id = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX1=4 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where id = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX1=5 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where id = "&Id1&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rss1.open sql1,conn,1,1


  if yy=1 then
  if LX1=1 THEN
  lanmuurl=""&SitePath&""&iinfo&"/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=2 THEN
  lanmuurl=""&SitePath&""&product&"/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=3 THEN
  lanmuurl=""&SitePath&""&photo&"/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=4 THEN
  lanmuurl=""&SitePath&""&video&"/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  IF LX1=5 THEN
  lanmuurl=""&SitePath&""&down&"/?pone="&Id1&"&ptwo="&rss(0)&"&m="&rss(5)&"&Language="&rss(6)&""
  END IF
  ELSE
  IF LX1=1 THEN
  lanmuurl=""&SitePath&""&rss1(4)&"/"&rss(4)&"/"&newClass&""&Separated&""&rss(0)&""&Separated&"1.html"
  END IF
  IF LX1=2 THEN
  lanmuurl=""&SitePath&""&rss1(4)&"/"&rss(4)&"/"&productClass&""&Separated&""&rss(0)&""&Separated&"1.html"
  END IF
  IF LX1=3 THEN
  lanmuurl=""&SitePath&""&rss1(4)&"/"&rss(4)&"/"&alClass&""&Separated&""&rss(0)&""&Separated&"1.html"
  END IF
  IF LX1=4 THEN
  lanmuurl=""&SitePath&""&rss1(4)&"/"&rss(4)&"/"&tvClass&""&Separated&""&rss(0)&""&Separated&"1.html"
  END IF
  IF LX1=5 THEN
  lanmuurl=""&SitePath&""&rss1(4)&"/"&rss(4)&"/"&downClass&""&Separated&""&rss(0)&""&Separated&"1.html"
  END IF
  END IF
  
  
  if lx1=1 then
  countnum=Conn.Execute("Select count(*) from news where zmone="&Id1&" and zmtwo="&rss(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx1=2 then
  countnum=Conn.Execute("Select count(*) from product where zmone="&Id1&" and zmtwo="&rss(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx1=3 then
  countnum=Conn.Execute("Select count(*) from al where zmone="&Id1&" and zmtwo="&rss(0)&" and Lang='"&Language&"' ")(0)
  end If
   if lx1=4 then
  countnum=Conn.Execute("Select count(*) from zxsp where zmone="&Id1&" and zmtwo="&rss(0)&" and Lang='"&Language&"' ")(0)
  end if
   if lx1=5 then
  countnum=Conn.Execute("Select count(*) from download where zmone="&Id1&" and zmtwo="&rss(0)&" and Lang='"&Language&"'")(0)
  end if
 
  
    if lx1=1 then
  count=Conn.Execute("Select count(*) from newsclass where ParentClassID = "&Id1&" and Lang='"&Language&"' ")(0)
  end if
  if lx1=2 then
  count=Conn.Execute("Select count(*) from cpclass where ParentClassID = "&Id1&" and Lang='"&Language&"' ")(0)
  end if
  if lx1=3 then
  count=Conn.Execute("Select count(*) from alclass where ParentClassID = "&Id1&" and Lang='"&Language&"' ")(0)
  end if
  
   if lx1=4 then
  count=Conn.Execute("Select count(*) from spclass where ParentClassID = "&Id1&" and Lang='"&Language&"' ")(0)
  end If
   if lx1=5 then
  count=Conn.Execute("Select count(*) from xzclass where ParentClassID = "&Id1&" and Lang='"&Language&"' ")(0)
  end If
  
  
    if rss(2)<>"" then
fenleitp=""&sitepath&""&rss(2)&""
else
fenleitp=""&sitepath&"images/demo.jpg"
end if

if request.QueryString("ptwo")="" then
po=0
else
po=request.QueryString("ptwo")
end if
  
  zmcms1=replace(zmcms,"{lanmu1:i}",linenums)
  zmcms2=replace(zmcms1,"{lanmu1:id}",rss(0))
  zmcms3=replace(zmcms2,"{lanmu1:name}",rss(1))
  zmcms4=replace(zmcms3,"{lanmu1:url}",lanmuurl)
  zmcms5=replace(zmcms4,"{lanmu1:img}",fenleitp)
  zmcms6=replace(zmcms5,"{lanmu1:wapurl}",lanmuurl1)
  zmcms7=replace(zmcms6,"{lanmu1:child}",rss(3))
  zmcms8=replace(zmcms7,"{lanmu1:num}",countnum)
  zmcms9=replace(zmcms8,"{lanmu1:mulu}",rss(4))
  zmcms10=replace(zmcms9,"{lanmu1:lujing}",""&sitepath&"")
  zmcms11=replace(zmcms10,"{lanmu1:po}",po)
  zmcms12=replace(zmcms11,"{lanmu1:lang}",rss(6))
  zmcms13=replace(zmcms12,"{lanmu1:addlang}",language)
  zmcms14=replace(zmcms13,"{lanmu1:count",count)
  mm=mm+zmcms14

'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if
rss.movenext
loop
Getlanmu1=Getlanmu1&mm
rss.Close
set rss=nothing
end Function

'-----三级导航循环
function lanmu2(str)
	set reg=new regexp
	reg.IgnoreCase = True 
	reg.Global = True 
	reg.Pattern = "<lanmu2:\{(.+?)\}>([\s\S]*?)<lanmu2>"
	set aa=reg.execute(str)
	for each match in aa
	bq=match.submatches(0)
	nr=match.submatches(1)
	Id2=nbq(bq,"Id2")
	Id3=nbq(bq,"Id3")
	LX2=nbq(bq,"LX2")
	str=replace(str,match,Getlanmu2(nr,LX2,Id2,Id3))
	next
	lanmu2=str
end function
function Getlanmu2(zmcms,LX2,Id2,Id3)
  set rsss = server.createobject("adodb.recordset")
  if LX2=1 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where ParentClassID = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX2=2 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where ParentClassID = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX2=3 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where ParentClassID = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX2=4 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where ParentClassID = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX2=5 then
  sql="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where ParentClassID = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rsss.open sql,conn,1,1

  do while not rsss.eof
linenums = linenums + 1

if LX2=1 THEN
  lanmuurl1=""&SitePath&"wap/info/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=2 THEN
  lanmuurl1=""&SitePath&"wap/product/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=3 THEN
  lanmuurl1=""&SitePath&"wap/photo/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=4 THEN
  lanmuurl1=""&SitePath&"wap/video/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=5 THEN
  lanmuurl1=""&SitePath&"wap/down/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  
  '获取上一级菜单的目录名称
  
  
  set rsss1 = server.createobject("adodb.recordset")
  if LX2=1 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX2=2 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX2=3 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX2=4 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX2=5 then
  sql1="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rsss1.open sql1,conn,1,1
  
  
    set rsss2 = server.createobject("adodb.recordset")
  if LX2=1 then
  sql2="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where id = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX2=2 then
  sql2="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where id = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX2=3 then
  sql2="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where id = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX2=4 then
  sql2="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where id = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX2=5 then
  sql2="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where id = "&Id3&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rsss2.open sql2,conn,1,1
  
  
  
  
  
  
  
  
  
  if yy=1 then
  if LX2=1 THEN
  lanmuurl=""&SitePath&""&iinfo&"/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=2 THEN
  lanmuurl=""&SitePath&""&product&"/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=3 THEN
  lanmuurl=""&SitePath&""&photo&"/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=4 THEN
  lanmuurl=""&SitePath&""&video&"/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  IF LX2=5 THEN
  lanmuurl=""&SitePath&""&down&"/?pone="&Id2&"&ptwo="&Id3&"&pthree="&rsss(0)&"&m="&rsss(5)&"&Language="&rsss(6)&""
  END IF
  ELSE
  IF LX2=1 THEN
  lanmuurl=""&SitePath&""&rsss1(4)&"/"&rsss2(4)&"/"&rsss(4)&"/"&newClass&""&Separated&""&rsss(0)&""&Separated&"1.html"
  END IF
  IF LX2=2 THEN
  lanmuurl=""&SitePath&""&rsss1(4)&"/"&rsss2(4)&"/"&rsss(4)&"/"&productClass&""&Separated&""&rsss(0)&""&Separated&"1.html"
  END IF
  IF LX2=3 THEN
  lanmuurl=""&SitePath&""&rsss1(4)&"/"&rsss2(4)&"/"&rsss(4)&"/"&alClass&""&Separated&""&rsss(0)&""&Separated&"1.html"
  END IF
  IF LX2=4 THEN
  lanmuurl=""&SitePath&""&rsss1(4)&"/"&rsss2(4)&"/"&rsss(4)&"/"&tvClass&""&Separated&""&rsss(0)&""&Separated&"1.html"
  END IF
  IF LX2=5 THEN
  lanmuurl=""&SitePath&""&rsss1(4)&"/"&rsss2(4)&"/"&rsss(4)&"/"&downClass&""&Separated&""&rsss(0)&""&Separated&"1.html"
  END IF
  END IF
  
    if lx2=1 then
  countnum=Conn.Execute("Select count(*) from news where zmone="&Id2&" and zmtwo="&Id3&" and zmthree="&rsss(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx2=2 then
  countnum=Conn.Execute("Select count(*) from product where zmone="&Id2&" and zmtwo="&Id3&" and zmthree="&rsss(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx2=3 then
  countnum=Conn.Execute("Select count(*) from al where zmone="&Id2&" and zmtwo="&Id3&" and zmthree="&rsss(0)&" and Lang='"&Language&"' ")(0)
  end if

   if lx2=4 then
  countnum=Conn.Execute("Select count(*) from zxsp where zmone="&Id2&" and zmtwo="&Id3&" and zmthree="&rsss(0)&" and Lang='"&Language&"' ")(0)
  end If
     if lx2=5 then
  countnum=Conn.Execute("Select count(*) from download where zmone="&Id2&" and zmtwo="&Id3&" and zmthree="&rsss(0)&" and Lang='"&Language&"'")(0)
  end if
  





  if lx2=1 then
  count=Conn.Execute("Select count(*) from newsclass where ParentClassID = "&Id3&" and Lang='"&Language&"' ")(0)
  end if
  if lx2=2 then
  count=Conn.Execute("Select count(*) from cpclass where ParentClassID = "&Id3&" and Lang='"&Language&"' ")(0)
  end if
  if lx2=3 then
  count=Conn.Execute("Select count(*) from alclass where ParentClassID = "&Id3&" and Lang='"&Language&"' ")(0)
  end if
  
   if lx2=4 then
  count=Conn.Execute("Select count(*) from spclass where ParentClassID = "&Id3&" and Lang='"&Language&"' ")(0)
  end If
   if lx2=5 then
  count=Conn.Execute("Select count(*) from xzclass where ParentClassID = "&Id3&" and Lang='"&Language&"' ")(0)
  end If












  
if request.QueryString("pthree")="" then
po=0
else
po=request.QueryString("pthree")
end if
  
  
  
  if rsss(2)<>"" then
  fenleitp=""&sitepath&""&rsss(2)&""
  else
  fenleitp=""&sitepath&"images/demo.jpg"
  end if
  zmcms1=replace(zmcms,"{lanmu2:i}",linenums)
  zmcms2=replace(zmcms1,"{lanmu2:id}",rsss(0))
  zmcms3=replace(zmcms2,"{lanmu2:name}",rsss(1))
  zmcms4=replace(zmcms3,"{lanmu2:url}",lanmuurl)
  zmcms5=replace(zmcms4,"{lanmu2:img}",fenleitp)
  zmcms6=replace(zmcms5,"{lanmu2:wapurl}",lanmuurl1)
  zmcms7=replace(zmcms6,"{lanmu2:child}",rsss(3))
  zmcms8=replace(zmcms7,"{lanmu2:num}",countnum)
  zmcms9=replace(zmcms8,"{lanmu2:lanmu}",rsss(4))
  zmcms10=replace(zmcms9,"{lanmu2:lujing}",""&sitepath&"")
  zmcms11=replace(zmcms10,"{lanmu2:po}",po)
  zmcms12=replace(zmcms11,"{lanmu2:lang}",rsss(6))
  zmcms13=replace(zmcms12,"{lanmu2:addlang}",language)
  zmcms14=replace(zmcms13,"{lanmu2:count}",count)
  mm=mm+zmcms14
'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if
rsss.movenext
loop
Getlanmu2=Getlanmu2&mm
rsss.Close
set rsss=nothing
end Function


'**************************************************
'函 数 名：lmmc 根据不同的ID 显示不同的一级栏目名称，在调用同一模块时配合使用
'*************************************************
function lmmc(lx)
lmid=request.QueryString("pone")
if lmid<>"" then
set rs=server.CreateObject("adodb.recordset")
IF LX=1  THEN
rs.open "select * from newsclass where id=" & lmid & " and Lang='"&Language&"' order by px asc",conn,1,1
END IF
IF LX=2  THEN
rs.open "select * from cpclass where id=" & lmid & " and Lang='"&Language&"' order by px asc",conn,1,1
END IF
IF LX=3  THEN
rs.open "select * from alclass where id=" & lmid & " and Lang='"&Language&"' order by px asc",conn,1,1
END IF
IF LX=4  THEN
rs.open "select * from spclass where id=" & lmid & " and Lang='"&Language&"' order by px asc",conn,1,1
END IF
IF LX=5  THEN
rs.open "select * from xzclass where id=" & lmid & " and Lang='"&Language&"' order by px asc",conn,1,1
END IF
IF LX=6  THEN
rs.open "select * from mnclass where id=" & lmid & " and Lang='"&Language&"' order by px asc",conn,1,1
END IF
lmmc=rs("classname")
rs.close
set rs=nothing
else
if lx=1 then
lmmc=""&word(25)&""
end if
if lx=2 then
lmmc=""&word(26)&""
end if
if lx=3 then
lmmc=""&word(27)&""
end if
if lx=4 then
lmmc=""&word(28)&""
end if
if lx=5 then
lmmc=""&word(29)&""
end if



end if
end function
'**************************************************
'函 数 名：lanmu3 4  只显示二级或者是三级栏目
'*************************************************
function lanmu3(str)
set reg=new regexp
reg.IgnoreCase = True 
reg.Global = True 
reg.Pattern = "<lanmu3:\{(.+?)\}>([\s\S]*?)<lanmu3>"
set aa=reg.execute(str)
for each match in aa
bq=match.submatches(0)
nr=match.submatches(1)
Row=nbq(bq,"Row")
LX3=nbq(bq,"LX3")
str=replace(str,match,Getlanmu3(Row,LX3,nr))
next
lanmu3=str
end function
function Getlanmu3(Row,LX3,zmcms)
lmid=request.QueryString("pone")
  set rs33 = server.createobject("adodb.recordset")
  if LX3=1 then
  sql33="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from newsclass where ParentClassID = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX3=2 then
  sql33="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from cpclass where ParentClassID = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX3=3 then
  sql33="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from alclass where ParentClassID = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX3=4 then
  sql33="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from spclass where ParentClassID = "&lmid&" and xs= 1 and Lang='"&Language&"' order by px asc"
  end if
   if LX3=5 then
  sql33="select top "&Row&" id,classname,SmallImage,c_child,mulu,m,lang from xzclass where ParentClassID = "&lmid&" and xs= 1 and Lang='"&Language&"' order by px asc"
  end if
  rs33.open sql33,conn,1,1
  do while not rs33.eof
  
  if LX3=1 THEN
  lanmuurl1=""&SitePath&"wap/info/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=2 THEN
  lanmuurl1=""&SitePath&"wap/product/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=3 THEN
  lanmuurl1=""&SitePath&"wap/photo/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=4 THEN
  lanmuurl1=""&SitePath&"wap/video/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=5 THEN
  lanmuurl1=""&SitePath&"wap/down/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  
  
    '获取上一级栏目的名称
   set rs333 = server.createobject("adodb.recordset")
  if LX3=1 then
  sql333="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where id = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX3=2 then
  sql333="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where id = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX3=3 then
  sql333="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where id = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX3=4 then
  sql333="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where id = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX3=5 then
  sql333="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where id = "&lmid&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rs333.open sql333,conn,1,1
 
  if yy=1 then
  if LX3=1 THEN
  lanmuurl=""&SitePath&""&iinfo&"/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=2 THEN
  lanmuurl=""&SitePath&""&product&"/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=3 THEN
  lanmuurl=""&SitePath&""&photo&"/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=4 THEN
  lanmuurl=""&SitePath&""&video&"/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  IF LX3=5 THEN
  lanmuurl=""&SitePath&""&down&"/?pone="&lmid&"&ptwo="&rs33(0)&"&m="&rs33(5)&"&Language="&rs33(6)&""
  END IF
  ELSE
  IF LX3=1 THEN
  lanmuurl=""&SitePath&""&rs333(4)&"/"&rs33(4)&"/"&newClass&""&Separated&""&rs33("id")&""&Separated&"1.html"
  END IF
  IF LX3=2 THEN
  lanmuurl=""&SitePath&""&rs333(4)&"/"&rs33(4)&"/"&productClass&""&Separated&""&rs33("id")&""&Separated&"1.html"
  END IF
  IF LX3=3 THEN
  lanmuurl=""&SitePath&""&rs333(4)&"/"&rs33(4)&"/"&alClass&""&Separated&""&rs33("id")&""&Separated&"1.html"
  END IF
  IF LX3=4 THEN
  lanmuurl=""&SitePath&""&rs333(4)&"/"&rs33(4)&"/"&tvClass&""&Separated&""&rs33("id")&""&Separated&"1.html"
  END IF
  IF LX3=5 THEN
  lanmuurl=""&SitePath&""&rs333(4)&"/"&rs33(4)&"/"&downClass&""&Separated&""&rs33("id")&""&Separated&"1.html"
  END IF
  END IF
  
  
  if lx3=1 then
  countnum=Conn.Execute("Select count(*) from news where zmone="&lmid&" and zmtwo="&rs33(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx3=2 then
  countnum=Conn.Execute("Select count(*) from product where zmone="&lmid&" and zmtwo="&rs33(0)&" and Lang='"&Language&"'  ")(0)
  end if
  if lx3=3 then
  countnum=Conn.Execute("Select count(*) from al where zmone="&lmid&" and zmtwo="&rs33(0)&" and Lang='"&Language&"'  ")(0)
  end if
   if lx3=5 then
  countnum=Conn.Execute("Select count(*) from download where zmone="&lmid&" and zmtwo="&rs33(0)&" and Lang='"&Language&"' ")(0)
  end if
   if lx3=4 then
  countnum=Conn.Execute("Select count(*) from zxsp where zmone="&lmid&" and zmtwo="&rs33(0)&" and Lang='"&Language&"' ")(0)
  end if
  
  


   if lx3=1 then
  count=Conn.Execute("Select count(*) from newsclass where ParentClassID = "&lmid&" and Lang='"&Language&"' ")(0)
  end if
  if lx3=2 then
  count=Conn.Execute("Select count(*) from cpclass where ParentClassID = "&lmid&" and Lang='"&Language&"' ")(0)
  end if
  if lx3=3 then
  count=Conn.Execute("Select count(*) from alclass where ParentClassID = "&lmid&" and Lang='"&Language&"' ")(0)
  end if
  
   if lx3=4 then
  count=Conn.Execute("Select count(*) from spclass where ParentClassID = "&lmid&" and Lang='"&Language&"' ")(0)
  end If
   if lx3=5 then
  count=Conn.Execute("Select count(*) from xzclass where ParentClassID = "&lmid&" and Lang='"&Language&"' ")(0)
  end If








  
  linenums = linenums + 1
  
  
  
if request.QueryString("ptwo")="" then
po=0
else
po=request.QueryString("ptwo")
end if
  
  
  
  
    if rs33(2)<>"" then
fenleitp=""&sitepath&""&rs33(2)&""
else
fenleitp=""&sitepath&"images/demo.jpg"
end if
  zmcms1=replace(zmcms,"{lanmu3:i}",linenums)
  zmcms2=replace(zmcms1,"{lanmu3:id}",rs33(0))
  zmcms3=replace(zmcms2,"{lanmu3:name}",rs33(1))
  zmcms4=replace(zmcms3,"{lanmu3:url}",lanmuurl)
  zmcms5=replace(zmcms4,"{lanmu3:img}",fenleitp)
  zmcms6=replace(zmcms5,"{lanmu3:wapurl}",lanmuurl1)
  zmcms7=replace(zmcms6,"{lanmu3:child}",rs33(3))
  zmcms8=replace(zmcms7,"{lanmu3:num}",countnum)
  zmcms9=replace(zmcms8,"{lanmu3:lanmu}",rs33(4))
  zmcms10=replace(zmcms9,"{lanmu3:lujing}",""&sitepath&"")
  zmcms11=replace(zmcms10,"{lanmu3:po}",po)
  zmcms12=replace(zmcms11,"{lanmu3:lang}",rs33(6))
  zmcms13=replace(zmcms12,"{lanmu3:addlang}",language)
   zmcms14=replace(zmcms13,"{lanmu3:count}",count)
  mm=mm+zmcms14
 

  
'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if
rs33.movenext
loop
Getlanmu3=Getlanmu3&mm
rs33.Close
set rs=nothing
end Function
  
  
 
function lanmu4(str)
	set reg=new regexp
	reg.IgnoreCase = True 
	reg.Global = True 
	reg.Pattern = "<lanmu4:\{(.+?)\}>([\s\S]*?)<lanmu4>"
	set aa=reg.execute(str)
	for each match in aa
	bq=match.submatches(0)
	nr=match.submatches(1)
	Id4=nbq(bq,"Id4")
	LX4=nbq(bq,"LX4")
	str=replace(str,match,Getlanmu4(nr,LX4,Id4))
	next
	lanmu4=str
end function
function Getlanmu4(zmcms,LX4,Id4)
lmid4=request.QueryString("pone")
  set rs44 = server.createobject("adodb.recordset")
  if LX4=1 then
  sql44="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where ParentClassID = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX4=2 then
  sql44="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where ParentClassID = "&Id4&" and xs= 1  and Lang='"&Language&"' order by px asc"
  end if
  if LX4=3 then
  sql44="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where ParentClassID = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX4=4 then
  sql44="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where ParentClassID = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX4=5 then
  sql44="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where ParentClassID = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rs44.open sql44,conn,1,1


  do while not rs44.eof
linenums = linenums + 1

 if LX4=1 THEN
  lanmuurl1=""&SitePath&"wap/info/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=2 THEN
  lanmuurl1=""&SitePath&"wap/product/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=3 THEN
  lanmuurl1=""&SitePath&"wap/photo/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=4 THEN
  lanmuurl1=""&SitePath&"wap/video/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=5 THEN
  lanmuurl1=""&SitePath&"wap/down/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF


 '获取上一级菜单的目录名称
  
  
  set rs441 = server.createobject("adodb.recordset")
  if LX4=1 then
  sql441="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX4=2 then
  sql441="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX4=3 then
  sql441="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX4=4 then
  sql441="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX4=5 then
  sql441="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where ParentClassID = 0 and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rs441.open sql441,conn,1,1
  
  
    set rs442 = server.createobject("adodb.recordset")
  if LX4=1 then
  sql442="select id,classname,SmallImage,c_child,mulu,m,lang from newsclass where id = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX4=2 then
  sql442="select id,classname,SmallImage,c_child,mulu,m,lang from cpclass where id = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
  if LX4=3 then
  sql442="select id,classname,SmallImage,c_child,mulu,m,lang from alclass where id = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX4=4 then
  sql442="select id,classname,SmallImage,c_child,mulu,m,lang from spclass where id = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if
   if LX4=5 then
  sql442="select id,classname,SmallImage,c_child,mulu,m,lang from xzclass where id = "&Id4&" and xs= 1 and Lang='"&Language&"'  order by px asc"
  end if

  rs442.open sql442,conn,1,1
  
if yy=1 then
  if LX4=1 THEN
  lanmuurl=""&SitePath&""&iinfo&"/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=2 THEN
  lanmuurl=""&SitePath&""&product&"/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=3 THEN
  lanmuurl=""&SitePath&""&photo&"/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=4 THEN
  lanmuurl=""&SitePath&""&video&"/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  IF LX4=5 THEN
  lanmuurl=""&SitePath&""&down&"/?pone="&lmid4&"&ptwo="&Id4&"&pthree="&rs44(0)&"&m="&rs44(5)&"&Language="&rs44(6)&""
  END IF
  ELSE
  IF LX4=1 THEN
  lanmuurl=""&SitePath&""&rs441(4)&"/"&rs442(4)&"/"&rs44(4)&"/"&newClass&""&Separated&""&rs44(0)&""&Separated&"1.html"
  END IF
  IF LX4=2 THEN
  lanmuurl=""&SitePath&""&rs441(4)&"/"&rs442(4)&"/"&rs44(4)&"/"&productClass&""&Separated&""&rs44(0)&""&Separated&"1.html"
  END IF
  IF LX4=3 THEN
  lanmuurl=""&SitePath&""&rs441(4)&"/"&rs442(4)&"/"&rs44(4)&"/"&alClass&""&Separated&""&rs44(0)&""&Separated&"1.html"
  END IF
  IF LX4=4 THEN
  lanmuurl=""&SitePath&""&rs441(4)&"/"&rs442(4)&"/"&rs44(4)&"/"&tvClass&""&Separated&""&rs44(0)&""&Separated&"1.html"
  END IF
  IF LX4=5 THEN
  lanmuurl=""&SitePath&""&rs441(4)&"/"&rs442(4)&"/"&rs44(4)&"/"&downClass&""&Separated&""&rs44(0)&""&Separated&"1.html"
  END IF
  END IF
  
 if lx4=1 then
  countnum=Conn.Execute("Select count(*) from news where  zmone="&lmid4&" and zmtwo="&Id4&" and zmthree="&rs44(0)&" and Lang='"&Language&"' ")(0)
  end if
  if lx4=2 then
  countnum=Conn.Execute("Select count(*) from product where zmone="&lmid4&" and zmtwo="&Id4&" and zmthree="&rs44(0)&" and Lang='"&Language&"'  ")(0)
  end if
  if lx4=3 then
  countnum=Conn.Execute("Select count(*) from al where zmone="&lmid4&" and zmtwo="&Id4&" and zmthree="&rs44(0)&" and Lang='"&Language&"'  ")(0)
  end if
   if lx4=5 then
  countnum=Conn.Execute("Select count(*) from download where zmone="&lmid4&" and zmtwo="&Id4&" and zmthree="&rs44(0)&" and Lang='"&Language&"' ")(0)
  end if
   if lx4=4 then
  countnum=Conn.Execute("Select count(*) from zxsp where zmone="&lmid4&" and zmtwo="&Id4&" and zmthree="&rs44(0)&" and Lang='"&Language&"' ")(0)
  end If




    if lx4=1 then
  count=Conn.Execute("Select count(*) from newsclass where ParentClassID = "&Id4&" and Lang='"&Language&"' ")(0)
  end if
  if lx4=2 then
  count=Conn.Execute("Select count(*) from cpclass where ParentClassID = "&Id4&" and Lang='"&Language&"' ")(0)
  end if
  if lx4=3 then
  count=Conn.Execute("Select count(*) from alclass where ParentClassID = "&Id4&" and Lang='"&Language&"' ")(0)
  end if
  
   if lx4=4 then
  count=Conn.Execute("Select count(*) from spclass where ParentClassID = "&Id4&" and Lang='"&Language&"' ")(0)
  end If
   if lx4=5 then
  count=Conn.Execute("Select count(*) from xzclass where ParentClassID = "&Id4&" and Lang='"&Language&"' ")(0)
  end If
  
  if request.QueryString("pthree")="" then
po=0
else
po=request.QueryString("pthree")
end if
  
  
    if rs44(2)<>"" then
fenleitp=""&sitepath&""&rs44(2)&""
else
fenleitp=""&sitepath&"images/demo.jpg"
end if
  
  zmcms1=replace(zmcms,"{lanmu4:i}",linenums)
  zmcms2=replace(zmcms1,"{lanmu4:id}",rs44(0))
  zmcms3=replace(zmcms2,"{lanmu4:name}",rs44(1))
  zmcms4=replace(zmcms3,"{lanmu4:url}",lanmuurl)
  zmcms5=replace(zmcms4,"{lanmu4:img}",fenleitp)
  zmcms6=replace(zmcms5,"{lanmu4:wapurl}",lanmuurl1)
  zmcms7=replace(zmcms6,"{lanmu4:child}",rs44(3))
  zmcms8=replace(zmcms7,"{lanmu4:num}",countnum)
  zmcms9=replace(zmcms8,"{lanmu4:lanmu}",rs44(4))
  zmcms10=replace(zmcms9,"{lanmu4:lujing}",""&sitepath&"")
  zmcms11=replace(zmcms10,"{lanmu4:po}",po)
  zmcms12=replace(zmcms11,"{lanmu4:lang}",rs44(6))
  zmcms13=replace(zmcms12,"{lanmu4:addlang}",language)
  zmcms14=replace(zmcms13,"{lanmu4:count}",count)
  mm=mm+zmcms14
 

'{FH:2:twostyle}
if RegExpTest(mm,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(mm,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
mm = ReplaceTest(mm,"{FH:\d:([^}]*)}","")
end if
end if

rs44.movenext
loop
Getlanmu4=Getlanmu4&mm
rs44.Close
set rs44=nothing
end Function



'**************************************************
'函 数 名：diyfoucs
'作    用：用于自定定义幻灯样式接口 ，可自由制作不同的幻灯样式
'*************************************************-
function diyfoucs(style,PID,msql,p)
dim N
N=hdsl
set RSZM=server.createobject("adodb.recordset")
ZM="select top "&N&" * from banner where  zmone="&PID&" and Lang='"&Language&"' "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
RSZM.open ZM,conn,1,1	
 	
For i = 1 To RSZM.RecordCount
linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{幻灯ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{幻灯名称}",RSZM("banner_name"))
tempListStr = replace(tempListStr,"{幻灯地址}",RSZM("banner_url"))
if rszm("uploadfile")<>"" then
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&RSZM("uploadfile")&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}","")
end if
tempListStr = replace(tempListStr,"{幻灯说明}",RSZM("content"))
returnStr = returnStr & tempListStr
RSZM.movenext
next
RSZM.close 
set RSZM=nothing
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)
diyfoucs = returnStr
end function

'**************************************************
'函 数 名：lan
'作    用：用于显示前台语言
'*************************************************-
function lan(style,w,N,msql,p)
set RSZM=server.createobject("adodb.recordset")
ZM="select top "&N&" * from lang where  xs=1  "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
RSZM.open ZM,conn,1,1	
 	
For i = 1 To RSZM.RecordCount


	
linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
if request.QueryString("language")="" then
l=langqt
else
l=request.QueryString("language")
end If
	      
	 
set RSZM1=server.createobject("adodb.recordset")
ZM1="select * from lang where langbs='"&l&"' "		 
RSZM1.open ZM1,conn,1,1	
lid=RSZM1("id")
RSZM1.close 
Set RSZM1=Nothing

MyComp = StrComp(l,rszm("langbs"),1)  
if w=1 then
if yy=1 then
ljdz=""&sitepath&"?language="&RSZM("langbs")&""
else
ljdz=""&sitepath&"default"&RSZM("langbs")&".html"
end if
else
ljdz=""&sitepath&"wap/?language="&RSZM("langbs")&""
end if

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{语言ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{地址栏ID}",lid)
tempListStr = replace(tempListStr,"{是否静态}",yy)
tempListStr = replace(tempListStr,"{语言名称}",RSZM("langname"))
tempListStr = replace(tempListStr,"{语言标识}",RSZM("langbs"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{语言图片}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{语言图片}","")
end if
tempListStr = replace(tempListStr,"{条件对比}",MyComp)
tempListStr = replace(tempListStr,"{链接地址}",ljdz)
returnStr = returnStr & tempListStr
RSZM.movenext
next
RSZM.close 
set RSZM=nothing
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)
lan = returnStr
end function


'**************************************************
'函 数 名：diyqq
'*************************************************-
function diyqq(lx,style,msql,p)

set RSZM=server.createobject("adodb.recordset")
if lx=1 then
ZM="select * from kf where Lang='"&Language&"'  "
end if
if lx=2 then
ZM="select * from ww where Lang='"&Language&"' "
end if

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""



tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
RSZM.open ZM,conn,1,1	
 	
For i = 1 To RSZM.RecordCount
linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
if lx=1 then
tempListStr = replace(tempListStr,"{QQID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{QQ号码}",RSZM("qq"))
tempListStr = replace(tempListStr,"{QQ说明}",RSZM("qqsm"))
end if
if lx=2 then
tempListStr = replace(tempListStr,"{WWID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{WW号码}",RSZM("ww"))
tempListStr = replace(tempListStr,"{WW说明}",RSZM("wwsm"))
end if
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
returnStr = returnStr & tempListStr
RSZM.movenext
next
RSZM.close 
set RSZM=nothing
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)
diyqq = returnStr
end function

'-------自定义幻灯结束--------

'通用信息函数，可以在首页，或者是列表页，或者是内容页，可以在网站的任何地方调用！
function info(LX,w,style,msql,gl,N,FL,F,FR,TZ,CZ,D,D1,P)
set rszm=server.createobject("adodb.recordset")
'单页信息调用
if LX=1 then
zm="select top "&N&" * from dy where Lang='"&Language&"' "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
'新闻信息调用
IF LX=2 THEN
if sqlno=1 then
ZM="select top "&N&" * from news where adddate<= now() and Lang='"&Language&"'"
else
ZM="select top "&N&" * from news where adddate<= getdate() and Lang='"&Language&"'"
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF
'产品信息调用
IF LX=3 THEN
if sqlno=1 then
ZM="select top "&N&" * from product where adddate<= now() and Lang='"&Language&"'"
else
ZM="select top "&N&" * from product where adddate<= getdate() and Lang='"&Language&"'"
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF
'图片信息调用
IF LX=4 THEN
if sqlno=1 then
ZM="select top "&N&" * from al where adddate<= now() and Lang='"&Language&"'"
else
ZM="select top "&N&" * from al where adddate<= getdate() and Lang='"&Language&"'"
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF
'视频信息调用
IF LX=5 THEN
if sqlno=1 then
ZM="select top "&N&" * from zxsp where adddate<= now() and Lang='"&Language&"'"
else
ZM="select top "&N&" * from zxsp where adddate<= getdate() and Lang='"&Language&"'"
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF
'下载信息调用
IF LX=6 THEN
if sqlno=1 then
ZM="select top "&N&" * from download where adddate<= now() and Lang='"&Language&"'"
else
ZM="select top "&N&" * from download where adddate<= getdate() and Lang='"&Language&"'"
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF
'招聘信息调用
IF LX=7 THEN
if sqlno=1 then
ZM="select top "&N&" * from job where ksdate<= now() and Lang='"&Language&"'"
else
ZM="select top "&N&" * from job where ksdate<= getdate() and Lang='"&Language&"'"
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF
'友情链接调用
IF LX=8 THEN
zm="select top "&N&" * from yqlj where Lang='"&Language&"'"
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF

'营销网络调用
IF LX=9 THEN
zm="select top "&N&" * from NetWork where Lang='"&Language&"' "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
END IF
'----------------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'----------------------调用样式结束
rszm.open zm,conn,1,1
If not( rszm.bof and rszm.EOF) Then
For i = 1 To rszm.RecordCount
'LX 模块 所属分类带链接开始
if lx=1 or lx=7 or lx=8 or lx=9  then
else
dim rsc
set rsc=server.CreateObject("adodb.recordset")
if lx=2 then
rsc.open "select * from newsclass where Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc.open "select * from cpclass where Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc.open "select * from alclass where Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc.open "select * from spclass where Lang='"&Language&"'",conn,1,1
end if
if lx=6 then
rsc.open "select * from xzclass where Lang='"&Language&"'",conn,1,1
end if
while not rsc.eof
if rszm("zmtwo")=0  and rsc("id") =rszm("zmone") and rszm("zmthree")=0 then
if lx=2 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rsc("id")&">&m="&rsc("m")&"&Language="&rsc("lang")&"" & rsc("classname") & "</a>"
end if

' 获取m的值
mqh=rsc("m")


end if

if lx=3 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if

if lx=4 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if

if lx=5 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if
if lx=6 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
end if
mqh=rsc("m")
end if


if rszm("zmone")<>0  and  rszm("zmtwo")= rsc("id") and rszm("zmthree")=0 then


set rsc1=server.CreateObject("adodb.recordset")
if lx=2 then
rsc1.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if lx=3 then
rsc1.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc1.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc1.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=6 then
rsc1.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if lx=2 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if

' 获取m的值
mqh=rsc("m")
end if
if lx=3 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if
if lx=4 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if
if lx=5 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if

if lx=6 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
end if
mqh=rsc("m")
end if
if rszm("zmone")<>0  and  rszm("zmtwo")<>0  and rszm("zmthree")=rsc("id") then

set rsc3=server.CreateObject("adodb.recordset")
if lx=2 then
rsc3.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc3.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc3.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc3.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=6 then
rsc3.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

set rsc4=server.CreateObject("adodb.recordset")
if lx=2 then
rsc4.open "select * from newsclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc4.open "select * from cpclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc4.open "select * from alclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc4.open "select * from spclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=6 then
rsc4.open "select * from xzclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if





if lx=2 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if

' 获取m的值
mqh=rsc("m")
end if
if lx=3 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if
if lx=4 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if
if lx=5 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
mqh=rsc("m")
end if
if lx=6 then
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&rsc("m")&"&Language="&rsc("lang")&">" & rsc("classname") & "</a>"
end if
end if
mqh=rsc("m")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
end if
'LX 模块 所属分类带链接结束


'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束



'定义链接地址路径开始

if lx=1 then
if w=1 then
tempPage=""&sitepath&""&aabout&"/"
else
tempPage=""&sitepath&"wap/about/"
end if
end if
if lx=2 then
if w=1 then
tempPage=""&sitepath&""&showinfo&"/"
else
tempPage=""&sitepath&"wap/showinfo/"
end if
end if
if lx=3 then
if w=1 then
tempPage=""&sitepath&""&showproduct&"/"
else
tempPage=""&sitepath&"wap/showproduct/"
end if
end if
if lx=4 then
if w=1 then
tempPage=""&sitepath&""&showphoto&"/"
else
tempPage=""&sitepath&"wap/showphoto/"
end if
end if
if lx=5 then
if w=1 then
tempPage=""&sitepath&""&showvideo&"/"
else
tempPage=""&sitepath&"wap/showvideo/"
end if
end if
if lx=6 then
if w=1 then
tempPage=""&sitepath&""&showdown&"/"
else
tempPage=""&sitepath&"wap/showdown/"
end if
end if
if lx=7 then
if w=1 then
tempPage=""&sitepath&""&showjob&"/"
else
tempPage=""&sitepath&"wap/showjob/"
tempPage2=""&sitepath&"wap/"
end if
end if


if lx=9 then
if w=1 then
tempPage=""&sitepath&""&yyxwlxx&"/"
else
tempPage=""&sitepath&"wap/yxwlxx/"
end if
end if






'--------------------定义链接地址路径结束

'-------------tags标签定义开始------------------
if lx=2 then
tt=1
end if
if lx=3 then
tt=2
end if
if lx=4 then
tt=3
end if
if lx=5 then
tt=4
end if
if lx=6 then
tt=5
end if
if lx=1 or lx=7 or lx=8 or lx=9 then
else
keyword=rszm("keyword")
keyword1=Split(keyword,"|")
keyword2=ubound(keyword1)
for m=0 to keyword2
if keyword2=0 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
end if
end if
if keyword2=1 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
end if
end if
if keyword2=2 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
end if
end if
if keyword2=3 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
end if
end if
if keyword2=4 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
end if
end if
if keyword2=5 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
end if
end if
if keyword2=6 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
end if
end if
if keyword2=7 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
end if
end if
if keyword2=8 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
end if
end if
if keyword2=9 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
end if
end if
if keyword2=10 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
end if
end if
if keyword2=11 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
end if
end if
if keyword2=12 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
end if
end if
if keyword2=13 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
end if
end if
if keyword2=14 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
end if
end if
if keyword2=15 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
end if
end if
if keyword2=16 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
end if
end if
if keyword2=17 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
end if
end if
if keyword2=18 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
end if
end if
if keyword2=19 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
end if
end if
if keyword2=20 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
m20="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(20))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(20)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
m20="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(20))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(20)&"</a>"
end if
end if
if keyword2=0 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&FR&"")
end if
if keyword2=1 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&FR&"")
end if
if keyword2=2 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&FR&"")
end if
if keyword2=3 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&FR&"")
end if
if keyword2=4 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&FR&"")
end if
if keyword2=5 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&FR&"")
end if

if keyword2=6 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&FR&"")
end if
if keyword2=7 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&FR&"")
end if
if keyword2=8 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&FR&"")
end if
if keyword2=9 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&FR&"")
end if
if keyword2=10 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&FR&"")
end if
if keyword2=11 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&FR&"")
end if
if keyword2=12 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&FR&"")
end if
if keyword2=13 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&FR&"")
end if
if keyword2=14 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&FR&"")
end if
if keyword2=15 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&FR&"")
end if
if keyword2=16 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&FR&"")
end if
if keyword2=17 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&FR&"")
end if
if keyword2=18 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&FR&"")
end if
if keyword2=19 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&F&""&m19&""&FR&"")
end if
if keyword2=20 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&F&""&m19&""&F&""&m20&""&FR&"")
end if
next

end if

'-------------tags标签定义结束-------------------
'单页内容替换开始
if lx=1 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{单页标题}",cutstrx(rszm("title"),TZ))
tempListStr = replace(tempListStr,"{提示标题}",rszm("title"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=1 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=1  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{单页ID}",rszm("id"))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{发布时间}",rszm("adddate"))
tempListStr = replace(tempListStr,"{单页缩略图地址}",RSZM("SmallImage"))
if request.QueryString("aboutid")<>"" then
tempListStr = replace(tempListStr,"{单页AboutId}",request.QueryString("aboutid"))
else
tempListStr = replace(tempListStr,"{单页AboutId}",0)
end if
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
tempListStr = replace(tempListStr,"{单页内容}",cutstrx(rszm("content"),CZ))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{单页缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{单页缩略图}",""&sitepath&"images/demo.jpg")
end if

if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if



if rszm("wlurl")<>"" then
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
else
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&aboutID=" & rszm("id")&"&m="&gl&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&aboutname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&aboutID=" & rszm("id")&"&m="&gl&"&Language="&Language&"")
end if
end if
end if
'----------------------单页内容替换结束

'----------------------新闻模块替换开始
if lx=2  then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{新闻ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{是否头条}",rsZM("gg"))
tempListStr = replace(tempListStr,"{是否幻灯}",rsZM("tj"))
tempListStr = replace(tempListStr,"{是否置顶}",rsZM("tt"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("sytj"))
'tempListStr = replace(tempListStr,"{是否头条}",rsZM("ft"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{新闻标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{新闻标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=2 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=2  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{信息来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{新闻作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
tempListStr = replace(tempListStr,"{新闻缩略图地址}",RSZM("SmallImage")) '仅仅图片地址
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage")) '仅仅幻灯图片地址

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if

if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" ")
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if


if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then

if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")<>0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if RSZM("BFNR")="" THEN
tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
ELSE
tempListStr = replace(tempListStr,"{内容简介}",rszm("bfnr"))
END IF
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
END IF

'----------------------新闻模块替换结束

'--------产品模块替换开始
if lx=3 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{产品ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{产品库存}",RSZM("kucun1"))
tempListStr = replace(tempListStr,"{产品销量}",RSZM("sellbuys"))
tempListStr = replace(tempListStr,"{剩余库存}",RSZM("kucun1")-RSZM("sellbuys"))

tempListStr = replace(tempListStr,"{产品编号}",RSZM("Product_Id"))
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{产品标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{产品标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
tempListStr = replace(tempListStr,"{是否新品}",RSZM("Newproduct"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{产品来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{产品作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{标准价格}",formatNumber2(RSZM("scprice")))
tempListStr = replace(tempListStr,"{优惠价格}",formatNumber2(RSZM("hyprice")))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage")) '仅仅地址
tempListStr = replace(tempListStr,"{幻灯图地址}",RSZM("IcoImage")) '仅仅地址
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if
if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if

tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=3 and Lang='"&Language&"'  order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=3 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束

tempListStr = replace(tempListStr,"{订购地址}",""&sitepath&"member/ProductBuy/?step=1&id="&RSZM("id")&"&m="&gl&"&Language="&Language&"")
end if

'--------产品模块替换结束






'--------图片模块替换开始
if lx=4 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{图片ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",rsZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{图片标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{图片标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))


tempListStr = replace(tempListStr,"{图片来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{图片作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if
tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=4 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=4 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
end if
'--------图片模块替换结束
'--------视频模块替换开始
if lx=5 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{单视频地址}",RSZM("spurl"))
tempListStr = replace(tempListStr,"{多视频地址}",RSZM("arrspurl"))
tempListStr = replace(tempListStr,"{wap视频地址}",RSZM("wapspurl"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{视频ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{视频标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{视频标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{视频来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{视频作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if



if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=5 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=5 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
end if

'-----------------视频模块替换结束

'-----------------下载模块替换开始
if lx=6 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{下载ID}",rsZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{下载标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{下载标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{下载来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{下载作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{适用系统}",RSZM("System"))
tempListStr = replace(tempListStr,"{软件语言}",RSZM("Language"))
tempListStr = replace(tempListStr,"{软件类型}",RSZM("Softclass"))
tempListStr = replace(tempListStr,"{下载地址一}",RSZM("DownloadUrl"))
tempListStr = replace(tempListStr,"{下载地址二}",RSZM("DownloadUrl1"))
tempListStr = replace(tempListStr,"{下载地址三}",RSZM("DownloadUrl2"))

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=6 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=6 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{文件大小}",RSZM("zdy1"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&"images/demo.jpg")
end if


if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&" " )
end if
end if
tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
end if
'-----------------下载模块替换结束

'-----------------招聘模块替换开始
if lx=7 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)

tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{招聘颜色}",RSZM("btcolor"))
tempListStr = replace(tempListStr,"{招聘ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("ksDate"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{职位名称}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("HrName"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{职位名称}",cutstrx(rsZM("HrName"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("HrName"))
tempListStr = replace(tempListStr,"{需求人数}",RSZM("HrRequireNum"))
tempListStr = replace(tempListStr,"{工作地点}",RSZM("HrAddress"))
tempListStr = replace(tempListStr,"{每月工资}",RSZM("HrSalary"))
if DateDiff("d",""&rszm("ksDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{开始时间}","<font color="&sjys&">"&FormatTime(RSZM("ksDate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{开始时间}",FormatTime(RSZM("ksDate"),D))
end If
if DateDiff("d",""&rszm("ksDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{开始时间1}","<font color="&sjys&">"&FormatTime(RSZM("ksDate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{开始时间1}",FormatTime(RSZM("ksDate"),D1))
end If
if DateDiff("d",""&rszm("jsDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{结束时间}","<font color="&sjys&">"&FormatTime(RSZM("jsDate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{结束时间}",FormatTime(RSZM("jsDate"),D))
end if
if DateDiff("d",""&rszm("jsDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{结束时间1}","<font color="&sjys&">"&FormatTime(RSZM("jsDate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{结束时间1}",FormatTime(RSZM("jsDate"),D1))
end if
tempListStr = replace(tempListStr,"{发布结束时间}",RSZM("jsDate"))
tempListStr = replace(tempListStr,"{用人单位}",RSZM("zpdw"))
tempListStr = replace(tempListStr,"{联系电话}",RSZM("dh"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=7 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=7 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{单位地址}",RSZM("dz"))
tempListStr = replace(tempListStr,"{投递邮箱}",RSZM("emaill"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?jobID=" & rszm("id")&"&m="&gl&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & ""&zpname&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?jobID=" & rszm("id")&"&m="&gl&"&Language="&Language&""  )
end if
if w=1 then
tempListStr = replace(tempListStr,"{应聘地址}",tempPage & ""&ttjjl&"/?Quarters=" & server.urlencode(rszm("HrName"))&"&m="&gl&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{应聘地址}",tempPage2 & "tjjl/?Quarters=" & server.urlencode(rszm("HrName"))&"&m="&gl&"&Language="&Language&"" )
end if

if rszm("jsDate")>now() then
tempListStr = replace(tempListStr,"{招聘状态}",""&word(1)&"")
else
tempListStr = replace(tempListStr,"{招聘状态}",""&word(2)&"")
end if
end if
'-----------------招聘模块替换结束
'-----------------友情链接替换开始
if lx=8 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{链接ID}",rsZM("id"))
tempListStr = replace(tempListStr,"{链接名称}",rsZM("mc"))
if rszm("logourl")<>"" then
tempListStr = replace(tempListStr,"{链接logo}",""&sitepath&""&RSZM("logourl")&"")
else
tempListStr = replace(tempListStr,"{链接logo}",""&sitepath&"images/demo.jpg")
end if
tempListStr = replace(tempListStr,"{链接地址}",RSZM("dz"))
end if
'-----------------友情链接替换结束

'-----------------营销网络替换开始
if lx=9 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{营销网络ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{所属省份}",RSZM("Province"))
tempListStr = replace(tempListStr,"{公司名称}",RSZM("Company"))
tempListStr = replace(tempListStr,"{公司地址}",RSZM("Address"))
tempListStr = replace(tempListStr,"{联系人}",RSZM("User"))
tempListStr = replace(tempListStr,"{联系电话}",RSZM("Tel"))
tempListStr = replace(tempListStr,"{公司形象图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))

tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{公司形象图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{公司形象图}",""&sitepath&"images/demo.jpg")
end if

if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

tempListStr = replace(tempListStr,"{传真号码}",RSZM("Fax"))
tempListStr = replace(tempListStr,"{邮政编码}",RSZM("Zip"))
tempListStr = replace(tempListStr,"{公司网址}",RSZM("gsUrl"))
tempListStr = replace(tempListStr,"{公司简介}",cutstrx(rsZM("CONTENT"),CZ))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=8 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=8 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束





if w=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?yxwlID=" & rszm("id")&"&m="&gl&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?yxwlID=" & rszm("id")&"&m="&gl&"&Language="&Language&"" )
end if
end if
'---------营销网络调用结束






returnStr = returnStr & tempListStr
RSZM.movenext
next
RSZM.close 
set RSZM=nothing
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)



Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)

INFO = returnStr
else
INFO = word(0)
end if
end function






'**************************************************
'函 数 名：guidang (信息按照月份归档)
'*************************************************-
'----------模块列表页调用开始
function guidang(LX,n,w,style,m,l)


pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
set RSZM=server.createobject("adodb.recordset")
if l=0 then

if lx=1 then
zm="select year(adddate),month(adddate),count(id) from news where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=2 then
zm="select year(adddate),month(adddate),count(id) from product where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=3 then
zm="select year(adddate),month(adddate),count(id) from al where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=4 then
zm="select year(adddate),month(adddate),count(id) from zxsp where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=5 then
zm="select year(adddate),month(adddate),count(id) from download where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if

else

if pone="" then
'-----------打开不同类型的数据库
if lx=1 then
zm="select year(adddate),month(adddate),count(id) from news where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=2 then
zm="select year(adddate),month(adddate),count(id) from product where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=3 then
zm="select year(adddate),month(adddate),count(id) from al where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=4 then
zm="select year(adddate),month(adddate),count(id) from zxsp where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=5 then
zm="select year(adddate),month(adddate),count(id) from download where Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if

end if

if pone<>"" then
'-----------打开不同类型的数据库
if lx=1 then
zm="select year(adddate),month(adddate),count(id) from news where zmone="&pone&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=2 then
zm="select year(adddate),month(adddate),count(id) from product where zmone="&pone&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=3 then
zm="select year(adddate),month(adddate),count(id) from al where zmone="&pone&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=4 then
zm="select year(adddate),month(adddate),count(id) from zxsp where zmone="&pone&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=5 then
zm="select year(adddate),month(adddate),count(id) from download where zmone="&pone&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if

end if


if pone<>"" and ptwo<>"" then
'-----------打开不同类型的数据库
if lx=1 then
zm="select year(adddate),month(adddate),count(id) from news where zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=2 then
zm="select year(adddate),month(adddate),count(id) from product where zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=3 then
zm="select year(adddate),month(adddate),count(id) from al where zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=4 then
zm="select year(adddate),month(adddate),count(id) from zxsp where zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=5 then
zm="select year(adddate),month(adddate),count(id) from download where zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if

end if
end if


if pone<>"" and ptwo<>"" and pthree<>"" then
'-----------打开不同类型的数据库
if lx=1 then
zm="select year(adddate),month(adddate),count(id) from news where zmone="&pone&" and zmtwo="&ptwo&" and pthree="&pthree&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=2 then
zm="select year(adddate),month(adddate),count(id) from product where zmone="&pone&" and zmtwo="&ptwo&" and pthree="&pthree&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=3 then
zm="select year(adddate),month(adddate),count(id) from al where zmone="&pone&" and zmtwo="&ptwo&" and pthree="&pthree&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=4 then
zm="select year(adddate),month(adddate),count(id) from zxsp where zmone="&pone&" and zmtwo="&ptwo&" and pthree="&pthree&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if
if lx=5 then
zm="select year(adddate),month(adddate),count(id) from download where zmone="&pone&" and zmtwo="&ptwo&" and pthree="&pthree&" and Lang='"&Language&"' group by month(adddate),year(adddate) order By year(adddate) desc,month(adddate) desc"
end if

end if

if lx=1 then
if w=1 then
tempp= iinfo
else
tempp= "wap/info"
end if
end if
if lx=2 then
if w=1 then
tempp= product
else
tempp= "wap/product"
end if
end if
if lx=3 then
if w=1 then
tempp= photo
else
tempp= "wap/photo"
end if
end if
if lx=4 then
if w=1 then
tempp= video
else
tempp= "wap/video"
end if
end if
if lx=5 then
if w=1 then
tempp= down
else
tempp= "wap/down"
end if
end if

'-----------打开不同模块的类型结束
'--------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""&word(649)&""

end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'--------------调用样式结束
RSZM.open ZM,conn,1,1

If not( rszm.bof and rszm.EOF) Then
rszm.PageSize = n
iCount = rszm.RecordCount 
iPageSize = rszm.PageSize
maxpage = rszm.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rszm.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If 		
For i = 1 To x

'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束
if l=0 then
if request.QueryString("m")<>"" then
wz=""&sitepath&""&tempp&"/?m="&request.QueryString("m")&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&Language="&Language&""
else
wz=""&sitepath&""&tempp&"/?m="&m&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&Language="&Language&""
end if

else

if pone="" then
if request.QueryString("m")<>"" then
wz=""&sitepath&""&tempp&"/?m="&request.QueryString("m")&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&Language="&Language&""
else
wz=""&sitepath&""&tempp&"/?m="&m&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&Language="&Language&""
end if

end if


if pone<>"" then
if request.QueryString("m")<>"" then
wz=""&sitepath&""&tempp&"/?m="&request.QueryString("m")&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&pone="&pone&"&Language="&Language&""
else
wz=""&sitepath&""&tempp&"/?m="&m&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&pone="&pone&"&Language="&Language&""
end if

end if


if pone<>"" and ptwo<>"" then
if request.QueryString("m")<>"" then
wz=""&sitepath&""&tempp&"/?m="&request.QueryString("m")&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&pone="&pone&"&ptwo="&ptwo&"&Language="&Language&""
else
wz=""&sitepath&""&tempp&"/?m="&m&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&pone="&pone&"&ptwo="&ptwo&"&Language="&Language&""
end if

end if


if pone<>"" and ptwo<>"" and pthree<>"" then
if request.QueryString("m")<>"" then
wz=""&sitepath&""&tempp&"/?m="&request.QueryString("m")&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&Language="&Language&""
else
wz=""&sitepath&""&tempp&"/?m="&m&"&logdate="&rszm(0)&"-"&Right(100+rszm(1),2)&"&pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&Language="&Language&""
end if

end if







end if

if request.QueryString("logdate")<>"" then

rqlogdate=request.QueryString("logdate")
else
rqlogdate=0
end if

gdrq=rszm(0)-Right(100+rszm(1),2)

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{所属语言}",language)
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{归档日期月}",Right(100+rszm(1),2))
tempListStr = replace(tempListStr,"{归档日期月1EN}",MonthToShortEN(Right(100+rszm(1),2)))
tempListStr = replace(tempListStr,"{归档日期月2EN}",MonthToLongEN(Right(100+rszm(1),2)))
tempListStr = replace(tempListStr,"{归档日期年}",rszm(0))
tempListStr = replace(tempListStr,"{归档日期}",gdrq)
tempListStr = replace(tempListStr,"{归档数量}",rszm(2))
tempListStr = replace(tempListStr,"{归档日期rq}",rqlogdate)
tempListStr = replace(tempListStr,"{链接地址}",wz)
returnStr = returnStr & tempListStr
RSZM.movenext
next
rszm.Close
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)
guidang = returnStr
else
guidang = word(0)
end if
end function






















'----------模块列表页调用开始
function infolist(LX,w,style,msql,gl,FL,F,FR,TZ,D,D1,CZ,P)

set RSZM=server.createobject("adodb.recordset")
'-----------打开不同类型的数据库
if lx=6 then
if sqlno=1 then
zm = "select * from job where ksdate<=now() and Lang='"&Language&"'   "
else
zm = "select * from job where ksdate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=7 then
if ly=2 then
zm = "select * from ly where Lang='"&Language&"'  order by "&p&" "
else
zm = "select * from ly where sh=1 and Lang='"&Language&"' order by "&p&" "
end if
end if
if lx=8 then
Key = Request.QueryString("Key")
If Key = "" Then 
response.Redirect ".."&sitepath&""
end if
zm="select * from NetWork where Province = '" & Key& " ' and Lang='"&Language&"'"
end if
if lx=1 or lx=2 or lx=3 or lx=4 or lx=5 then
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")


if pone="" then
if lx=1 then
if sqlno=1 then
zm = "select * from news where adddate<=now() and Lang='"&Language&"' "
else
zm = "select * from news where adddate<=getdate() and Lang='"&Language&"' "
end if

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""

end if
if lx=2 then
if sqlno=1 then
zm = "select * from product where adddate<=now() and Lang='"&Language&"'   "
else
zm = "select * from product where adddate<=getdate() and Lang='"&Language&"'   "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""


end if
if lx=3 then
if sqlno=1 then
zm = "select * from al where adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from al where adddate<=getdate() and Lang='"&Language&"'  "
end if

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""



end if
if lx=4 then
if sqlno=1 then
zm = "select * from zxsp where adddate<=now() and Lang='"&Language&"'   "
else
zm = "select * from zxsp where adddate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=5 then
if sqlno=1 then
zm = "select * from download where adddate<=now() and Lang='"&Language&"'   "
else
zm = "select * from download where adddate<=getdate() and Lang='"&Language&"'   "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
end if
if pone<>""  then 
if lx=1 then
if sqlno=1 then
zm = "select * from news where zmone="&Pone&" and adddate<=now() and Lang='"&Language&"' "
else
zm = "select * from news where zmone="&Pone&" and adddate<=getdate() and Lang='"&Language&"' "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=2 then
if sqlno=1 then
zm = "select * from product where zmone="&Pone&" and adddate<=now() and Lang='"&Language&"'   "
else
zm = "select * from product where zmone="&Pone&" and adddate<=getdate() and Lang='"&Language&"'  "
end if


'筛选属性条件开始
If value <> "" Then 
zm=zm+" and ("
arr_value = Split(value,"|")
For iiiiii = 0 To Ubound(arr_value)
If Right(Trim(arr_value(iiiiii)),2) <> "_0" Then 
If sqlno =1 Then 
zm=zm+" instr(p_SelectValue,'_"&Split(arr_value(iiiiii),"_")(1)&"')>0"
else
zm=zm+" charindex('_"&Split(arr_value(iiiiii),"_")(1)&"',p_SelectValue)>0"
end if
Else
zm=zm+" 1=1"
End If 
If iiiiii < UBound(arr_value) Then zm=zm+" and "
Next
zm=zm+")"
End If 
'筛选属性条件结束
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
if sqlno=1 then
zm = "select * from al where zmone="&Pone&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from al where zmone="&Pone&" and adddate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=4 then
if sqlno=1 then
zm = "select * from zxsp where zmone="&Pone&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from zxsp where zmone="&Pone&" and adddate<=getdate() and Lang='"&Language&"'  "
end if

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""

end if
if lx=5 then
if sqlno=1 then
zm = "select * from download where zmone="&Pone&" and adddate<=now() and Lang='"&Language&"'   "
else
zm = "select * from download where zmone="&Pone&" and adddate<=getdate() and Lang='"&Language&"'   "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
end if
if ptwo<>"" and pone<>""  then
if lx=1 then
if sqlno=1 then
zm = "select * from news where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from news where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=getdate() and Lang='"&Language&"'  "
end if

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""


end if
if lx=2 then
if sqlno=1 then
zm = "select * from product where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=now() and Lang='"&Language&"'    "
else
zm = "select * from product where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=getdate() and Lang='"&Language&"'   "
end if

'筛选属性条件开始
If value <> "" Then 
zm=zm+" and ("
arr_value = Split(value,"|")
For iiiiii = 0 To Ubound(arr_value)
If Right(Trim(arr_value(iiiiii)),2) <> "_0" Then 
If sqlno =1 Then 
zm=zm+" instr(p_SelectValue,'_"&Split(arr_value(iiiiii),"_")(1)&"')>0"
else
zm=zm+" charindex('_"&Split(arr_value(iiiiii),"_")(1)&"',p_SelectValue)>0"
end if
Else
zm=zm+" 1=1"
End If 
If iiiiii < UBound(arr_value) Then zm=zm+" and "
Next
zm=zm+")"
End If 
'筛选属性条件结束
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
if sqlno=1 then
zm = "select * from al where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from al where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=4 then
if sqlno=1 then
zm = "select * from zxsp where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from zxsp where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=5 then
if sqlno=1 then
zm = "select * from download where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from download where zmone="&Pone&" and zmtwo="&Ptwo&" and adddate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
end if

if ptwo<>"" and pone<>"" and pthree<>""  then
if lx=1 then
if sqlno=1 then
zm = "select * from news where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=now() and Lang='"&Language&"' "
else
zm = "select * from news where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&"  and adddate<=getdate() and Lang='"&Language&"' "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=2 then
if sqlno=1 then
zm = "select * from product where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=now() and Lang='"&Language&"'   "
else
zm = "select * from product where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=getdate() and Lang='"&Language&"'   "
end if
'筛选属性条件开始
If value <> "" Then 
zm=zm+" and ("
arr_value = Split(value,"|")
For iiiiii = 0 To Ubound(arr_value)
If Right(Trim(arr_value(iiiiii)),2) <> "_0" Then 
If sqlno =1 Then 
zm=zm+" instr(p_SelectValue,'_"&Split(arr_value(iiiiii),"_")(1)&"')>0"
else
zm=zm+" charindex('_"&Split(arr_value(iiiiii),"_")(1)&"',p_SelectValue)>0"
end if
Else
zm=zm+" 1=1"
End If 
If iiiiii < UBound(arr_value) Then zm=zm+" and "
Next
zm=zm+")"
End If 
'筛选属性条件结束
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
if sqlno=1 then
zm = "select * from al where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from al where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=getdate() and Lang='"&Language&"'   "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=4 then
if sqlno=1 then
zm = "select * from zxsp where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from zxsp where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=5 then
if sqlno=1 then
zm = "select * from download where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=now() and Lang='"&Language&"'  "
else
zm = "select * from download where zmone="&Pone&" and zmtwo="&Ptwo&" and zmthree="&Pthree&" and adddate<=getdate() and Lang='"&Language&"'  "
end if
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
end if








end if
'-----------打开不同模块的类型结束
'--------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""&word(649)&""

end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'--------------调用样式结束
RSZM.open ZM,conn,1,1

If not( rszm.bof and rszm.EOF) Then

if lx=1 then
if w=1 then
rszm.PageSize = newinfo
else
rszm.PageSize = wapnewinfo
end if
end if
if lx=2 then
if w=1 then
rszm.PageSize = productInfo
else
rszm.PageSize = wapproductInfo
end if
end if
if lx=3 then
if w=1 then
rszm.PageSize = alInfo
else
rszm.PageSize = wapalInfo
end if
end if
if lx=4 then
if w=1 then
rszm.PageSize = tvInfo
else
rszm.PageSize = waptvInfo
end if
end if
if lx=5 then
if w=1 then
rszm.PageSize = downInfo
else
rszm.PageSize = wapdownInfo
end if
end if
if lx=6 then
if w=1 then
rszm.PageSize = zpInfo
else
rszm.PageSize = wapzpInfo
end if
end if

if lx=7 then
if w=1 then
rszm.PageSize = lyinfo
else
rszm.PageSize = waplyinfo
end if
end if

if lx=8 then
if w=1 then
rszm.PageSize = yxwlinfo
else
rszm.PageSize = wapyxwlinfo
end if
end if
iCount = rszm.RecordCount 
iPageSize = rszm.PageSize
maxpage = rszm.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rszm.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If 		
For i = 1 To x
'---------所属分类开始
if lx=6 or lx=7 or lx=8 then
else
dim rsc
set rsc=server.CreateObject("adodb.recordset")
if lx=1 then
rsc.open "select * from newsclass where Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc.open "select * from cpclass where Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc.open "select * from alclass where Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc.open "select * from spclass where Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc.open "select * from xzclass where Lang='"&Language&"'",conn,1,1
end if
while not rsc.eof
if rszm("zmtwo")=0  and rsc("id") =rszm("zmone") and rszm("zmthree")=0 then
if lx=1 then
mqh=rsc("m")
if w=1 then
if yy=1 then

ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=2 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=3 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=4 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=5 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
end if


if rszm("zmone")<>0  and  rszm("zmtwo")= rsc("id") and rszm("zmthree")=0 then


set rsc1=server.CreateObject("adodb.recordset")
if lx=1 then
rsc1.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if lx=2 then
rsc1.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc1.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc1.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc1.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if lx=1 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=2 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=3 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=4 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=5 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
end if
if rszm("zmone")<>0  and  rszm("zmtwo")<>0  and rszm("zmthree")=rsc("id") then

set rsc3=server.CreateObject("adodb.recordset")
if lx=1 then
rsc3.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc3.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc3.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc3.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc3.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

set rsc4=server.CreateObject("adodb.recordset")
if lx=1 then
rsc4.open "select * from newsclass where id="&rszm("zmtwo")&"and Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc4.open "select * from cpclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc4.open "select * from alclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc4.open "select * from spclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc4.open "select * from xzclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'<p></p>",conn,1,1
end if





if lx=1 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=2 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=3 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=4 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=5 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
end if
'LX 模块 所属分类带链接结束

'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束

'定义链接地址路径开始
if lx<>7  then

if lx=1 then
if w=1 then
tempPage=""&sitepath&""&showinfo&"/"
else
tempPage=""&sitepath&"wap/showinfo/"
end if
end if
if lx=2 then
if w=1 then
tempPage=""&sitepath&""&showproduct&"/"
else
tempPage=""&sitepath&"wap/showproduct/"
end if
end if
if lx=3 then
if w=1 then
tempPage=""&sitepath&""&showphoto&"/"
else
tempPage=""&sitepath&"wap/showphoto/"
end if
end if
if lx=4 then
if w=1 then
tempPage=""&sitepath&""&showvideo&"/"
else
tempPage=""&sitepath&"wap/showvideo/"
end if
end if
if lx=5 then
if w=1 then
tempPage=""&sitepath&""&showdown&"/"
else
tempPage=""&sitepath&"wap/showdown/"
end if
end if
if lx=6 then
if w=1 then
tempPage=""&sitepath&""&showjob&"/"
else
tempPage=""&sitepath&"wap/showjob/"
tempPage2=""&sitepath&"wap/"
end if
end if
if lx=8 then
if w=1 then
tempPage=""&sitepath&""&yyxwlxx&"/"
else
tempPage=""&sitepath&"wap/yxwlxx/"
end if
end if
end if
'--------------------定义链接地址路径结束

'-------------tags标签定义开始------------------
if lx=1 then
tt=1
end if
if lx=2 then
tt=2
end if
if lx=3 then
tt=3
end if
if lx=4 then
tt=4
end if
if lx=5 then
tt=5
end if
if lx=6 or lx=7 or lx=8 then
else
keyword=rszm("keyword")
keyword1=Split(keyword,"|")
keyword2=ubound(keyword1)
for m=0 to keyword2
if keyword2=0 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
end if
end if
if keyword2=1 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
end if
end if
if keyword2=2 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
end if
end if
if keyword2=3 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
end if
end if
if keyword2=4 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
end if
end if
if keyword2=5 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
end if
end if
if keyword2=6 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
end if
end if
if keyword2=7 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
end if
end if
if keyword2=8 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
end if
end if
if keyword2=9 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
end if
end if
if keyword2=10 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
end if
end if
if keyword2=11 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
end if
end if
if keyword2=12 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
end if
end if
if keyword2=13 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
end if
end if
if keyword2=14 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
end if
end if
if keyword2=15 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
end if
end if
if keyword2=16 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
end if
end if
if keyword2=17 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
end if
end if
if keyword2=18 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
end if
end if
if keyword2=19 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
end if
end if
if keyword2=20 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
m20="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(20))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(20)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
m20="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(20))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(20)&"</a>"
end if
end if
if keyword2=0 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&FR&"")
end if
if keyword2=1 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&FR&"")
end if
if keyword2=2 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&FR&"")
end if
if keyword2=3 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&FR&"")
end if
if keyword2=4 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&FR&"")
end if
if keyword2=5 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&FR&"")
end if

if keyword2=6 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&FR&"")
end if
if keyword2=7 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&FR&"")
end if
if keyword2=8 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&FR&"")
end if
if keyword2=9 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&FR&"")
end if
if keyword2=10 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&FR&"")
end if
if keyword2=11 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&FR&"")
end if
if keyword2=12 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&FR&"")
end if
if keyword2=13 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&FR&"")
end if
if keyword2=14 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&FR&"")
end if
if keyword2=15 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&FR&"")
end if
if keyword2=16 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&FR&"")
end if
if keyword2=17 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&FR&"")
end if
if keyword2=18 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&FR&"")
end if
if keyword2=19 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&F&""&m19&""&FR&"")
end if
if keyword2=20 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&F&""&m19&""&F&""&m20&""&FR&"")
end if
next

end if
'----------------------tags标签结束
'----------------------新闻模块替换开始
if lx=1 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{新闻ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{是否头条}",rsZM("gg"))
tempListStr = replace(tempListStr,"{是否幻灯}",rsZM("tj"))
tempListStr = replace(tempListStr,"{是否置顶}",rsZM("tt"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("sytj"))
'tempListStr = replace(tempListStr,"{是否头条}",rsZM("ft"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{新闻标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{新闻标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=2 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=2 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{信息来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{新闻作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
tempListStr = replace(tempListStr,"{新闻缩略图地址}",RSZM("SmallImage")) '仅仅图片地址
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage")) '仅仅幻灯图片地
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then

tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if


if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then

if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")<>0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if RSZM("BFNR")="" THEN
tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
ELSE
tempListStr = replace(tempListStr,"{内容简介}",rszm("bfnr"))
END IF
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
END IF
'----------------------新闻模块替换结束

'--------产品模块替换开始
if lx=2 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{产品ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{产品库存}",RSZM("kucun1"))
tempListStr = replace(tempListStr,"{产品销量}",RSZM("sellbuys"))
tempListStr = replace(tempListStr,"{剩余库存}",RSZM("kucun1")-RSZM("sellbuys"))
tempListStr = replace(tempListStr,"{产品编号}",RSZM("Product_Id"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{产品标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{产品标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
tempListStr = replace(tempListStr,"{是否新品}",RSZM("Newproduct"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{产品来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{产品作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{标准价格}",formatNumber2(RSZM("scprice")))
tempListStr = replace(tempListStr,"{优惠价格}",formatNumber2(RSZM("hyprice")))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage")) '仅仅地址
tempListStr = replace(tempListStr,"{幻灯图地址}",RSZM("IcoImage")) '仅仅地址
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then

sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )

else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=3 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=3 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{订购地址}",""&sitepath&"member/ProductBuy/?step=1&id="&RSZM("id")&"&m="&gl&"")
end if

'--------产品模块替换结束

'--------图片模块替换开始
if lx=3 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{图片ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",rsZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{图片标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{图片标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))

tempListStr = replace(tempListStr,"{图片来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{图片作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)


'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=4 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=4 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
end if
'--------图片模块替换结束
'--------视频模块替换开始
if lx=4  then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{单视频地址}",RSZM("spurl"))
tempListStr = replace(tempListStr,"{多视频地址}",RSZM("arrspurl"))
tempListStr = replace(tempListStr,"{wap视频地址}",RSZM("wapspurl"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{视频ID}",RSZM("id"))

tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{视频标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{视频标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))

tempListStr = replace(tempListStr,"{视频来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{视频作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if



if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=5 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=5 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束

tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
end if

'-----------------视频模块替换结束

'-----------------下载模块替换开始
if lx=5 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{下载ID}",rsZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",rsZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{下载标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{下载标题}",cutstrx(rsZM("title"),TZ))
END IF

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&"images/demo.jpg")
end if

if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if


tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{下载来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{下载作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{适用系统}",RSZM("System"))
tempListStr = replace(tempListStr,"{软件语言}",RSZM("Language"))
tempListStr = replace(tempListStr,"{软件类型}",RSZM("Softclass"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=6 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=6 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{下载地址一}",RSZM("DownloadUrl"))
tempListStr = replace(tempListStr,"{下载地址二}",RSZM("DownloadUrl1"))
tempListStr = replace(tempListStr,"{下载地址三}",RSZM("DownloadUrl2"))
tempListStr = replace(tempListStr,"{文件大小}",RSZM("zdy1"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&"images/demo.jpg")
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if





tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

end if
'-----------------下载模块替换结束

'-----------------招聘模块替换开始
if lx=6 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{招聘颜色}",RSZM("btcolor"))
tempListStr = replace(tempListStr,"{招聘ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("ksDate"))
tempListStr = replace(tempListStr,"{职位名称}","<font color="&rszm("btcolor")&">"&cutstrx(rsZM("HrName"),TZ)&"</font>")


IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{职位名称}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("HrName"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{职位名称}",cutstrx(rsZM("HrName"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("HrName"))

tempListStr = replace(tempListStr,"{需求人数}",RSZM("HrRequireNum"))
tempListStr = replace(tempListStr,"{工作地点}",RSZM("HrAddress"))
tempListStr = replace(tempListStr,"{每月工资}",RSZM("HrSalary"))
if DateDiff("d",""&rszm("ksDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{开始时间}","<font color="&sjys&">"&FormatTime(RSZM("ksDate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{开始时间}",FormatTime(RSZM("ksDate"),D))
end If
if DateDiff("d",""&rszm("ksDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{开始时间1}","<font color="&sjys&">"&FormatTime(RSZM("ksDate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{开始时间1}",FormatTime(RSZM("ksDate"),D1))
end if

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=7 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=7 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
if DateDiff("d",""&rszm("jsDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{结束时间}","<font color="&sjys&">"&FormatTime(RSZM("jsDate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{结束时间}",FormatTime(RSZM("jsDate"),D))
end if
if DateDiff("d",""&rszm("jsDate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{结束时间1}","<font color="&sjys&">"&FormatTime(RSZM("jsDate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{结束时间1}",FormatTime(RSZM("jsDate"),D1))
end if
tempListStr = replace(tempListStr,"{发布结束时间}",RSZM("jsDate"))
tempListStr = replace(tempListStr,"{用人单位}",RSZM("zpdw"))
tempListStr = replace(tempListStr,"{联系电话}",RSZM("dh"))
tempListStr = replace(tempListStr,"{单位地址}",RSZM("dz"))
tempListStr = replace(tempListStr,"{投递邮箱}",RSZM("emaill"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
tempListStr = replace(tempListStr,"{wap链接地址}",tempPage1 & "?jobID=" & rszm("id")&"&m="&request.QueryString("m")&"&Language="&Language&"" )
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?jobID=" & rszm("id")&"&m="&request.QueryString("m")&"&Language="&Language&""  )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&zpname&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?jobID=" & rszm("id")&"&m="&request.QueryString("m")&"&Language="&Language&""  )
end if
if w=1 then
tempListStr = replace(tempListStr,"{应聘地址}",tempPage & ""&ttjjl&"/?Quarters=" & server.urlencode(rszm("HrName"))&"&m="&request.QueryString("m")&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{应聘地址}",tempPage2 & "tjjl/?Quarters=" & server.urlencode(rszm("HrName"))&"&m="&request.QueryString("m")&"&Language="&Language&"" )
end if
if rszm("jsDate")>now() then
tempListStr = replace(tempListStr,"{招聘状态}",""&word(1)&"")
else
tempListStr = replace(tempListStr,"{招聘状态}",""&word(2)&"")
end if
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))

end if
'-----------------招聘模块替换结束

'---------留言模块调用开始
if lx=7 then

hqlydz= gethttp("http://whois.pconline.com.cn/ip.jsp?ip="&RSZM("lyip")&"","")

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)

tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{留言ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("sj"))
tempListStr = replace(tempListStr,"{留言用户}",RSZM("title"))
tempListStr = replace(tempListStr,"{电子邮件}",RSZM("emaill"))
tempListStr = replace(tempListStr,"{联系电话}",RSZM("phone"))
tempListStr = replace(tempListStr,"{联系QQ}",RSZM("qq"))
tempListStr = replace(tempListStr,"{留言详细地址}",hqlydz)
if DateDiff("d",""&rszm("sj")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{留言时间}","<font color="&sjys&">"&FormatTime(RSZM("sj"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{留言时间}",FormatTime(RSZM("sj"),D))
end if

 if rszm("qqh")=1 then
    set rs10 = server.createobject("adodb.recordset")
    sql10="select * from Members where ID="&rszm("MemID")&""
    rs10.open sql10,conn,1,1
    if rs10("MemName")=session("MemName") and session("MemLogin")="Succeed" then
	  MessageContent=rszm("Content")
	else
	  MessageContent=""&word(3)&""
	end if
    rs10.close
    set rs10=nothing
  else
    MessageContent=rszm("Content")
  end if
tempListStr = replace(tempListStr,"{留言内容}",MessageContent)

if rszm("qqh")=1 then

set rs11 = server.createobject("adodb.recordset")
sql11="select * from Members where ID="&rszm("MemID")&" and Lang='"&Language&"'"
rs11.open sql11,conn,1,1
if rs11("MemName")=session("MemName") and session("MemLogin")="Succeed" then
if IsNull(rszm("hf")) or trim(rszm("hf")&"")="" then
replayContent=""&word(4)&""
else
replayContent=rszm("hf")
end if
else
replayContent=""&word(3)&""
end if
rs11.close
set rs11=nothing
else
   
if trim(rszm("hf")&"")="" then
replayContent=""&word(4)&""
else
replayContent=rszm("hf")
end if
end if

tempListStr = replace(tempListStr,"{回复内容}",replayContent)
if trim(rszm("ReplyTime")&"")="" then
replaytime=""&word(4)&""
else
replaytime=rszm("ReplyTime")
end if
tempListStr = replace(tempListStr,"{回复时间}",replaytime)
end if



'---------留言模块调用结束

'---------营销网络调用开始
if lx=8 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))

tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{营销网络ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{所属省份}",RSZM("Province"))
tempListStr = replace(tempListStr,"{公司名称}",RSZM("Company"))
tempListStr = replace(tempListStr,"{公司地址}",RSZM("Address"))
tempListStr = replace(tempListStr,"{联系人}",RSZM("User"))
tempListStr = replace(tempListStr,"{联系电话}",RSZM("Tel"))
tempListStr = replace(tempListStr,"{公司形象图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{公司形象图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{公司形象图}",""&sitepath&"images/demo.jpg")
end if

if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if


tempListStr = replace(tempListStr,"{传真号码}",RSZM("Fax"))
tempListStr = replace(tempListStr,"{邮政编码}",RSZM("Zip"))
tempListStr = replace(tempListStr,"{公司网址}",RSZM("gsUrl"))
tempListStr = replace(tempListStr,"{公司简介}",cutstrx(rsZM("CONTENT"),CZ))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if


'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=8 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=8 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束




if w=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?yxwlID=" & rszm("id")&"&m="&request.QueryString("m")&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?yxwlID=" & rszm("id")&"&m="&request.QueryString("m")&"&Language="&Language&"" )
end if
end if
'---------营销网络调用结束
returnStr = returnStr & tempListStr
RSZM.movenext
next
rszm.Close
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)

'分页改动采用js分页！更新于2015年05.07
'更新人:追梦工作室
if instr(returnStr,"{分页}") > 0 then
if w=1 then
returnStr = replace(returnStr,"{分页}","<div id=""pagepc"" class=""page""></div>")
else
returnStr = replace(returnStr,"{分页}","<div id=""pagewap"" class=""page""></div>")
end if
end if 

if w=1 then
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagepc',"& vbCrLf
returnStr=returnStr&  "pages: "&maxpage&","& vbCrLf
returnStr=returnStr&  "skin: '"&word(475)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(476)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(477)&","& vbCrLf
returnStr=returnStr&  "next: "&word(478)&","& vbCrLf
if word(479)<>"" then
returnStr=returnStr&  "last: "&word(479)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(480)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(481)&","& vbCrLf
returnStr=returnStr&"curr: "&page&","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
if yy=1 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
else
'当pone为空时
if pone="" and ptwo="" and pthree="" then
if lx=1 then
returnStr=returnStr&"location.href ='"&Newclass&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=2 then
returnStr=returnStr&"location.href ='"&ProductClass&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=3 then
returnStr=returnStr&"location.href ='"&alClass&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=4 then
returnStr=returnStr&"location.href ='"&tvClass&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=5 then
returnStr=returnStr&"location.href ='"&downClass&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=6 then
returnStr=returnStr&"location.href ='"&zpClass&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=7 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if

if lx=8 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if
end if
'当pone不为空时
if pone<>"" and ptwo="" and pthree="" then
if lx=1 then
returnStr=returnStr&"location.href ='"&Newclass&""&Separated&""&pone&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=2 then
returnStr=returnStr&"location.href ='"&ProductClass&""&Separated&""&pone&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=3 then
returnStr=returnStr&"location.href ='"&alClass&""&Separated&""&pone&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=4 then
returnStr=returnStr&"location.href ='"&tvClass&""&Separated&""&pone&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=5 then
returnStr=returnStr&"location.href ='"&downClass&""&Separated&""&pone&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=6 then
returnStr=returnStr&"location.href ='"&zpClass&""&Separated&""&pone&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=7 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if

if lx=8 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if
end if
 '当pone  并且 ptwo 不为空时
if pone<>"" and ptwo<>"" and pthree="" then
if lx=1 then
returnStr=returnStr&"location.href ='"&Newclass&""&Separated&""&ptwo&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=2 then
returnStr=returnStr&"location.href ='"&ProductClass&""&Separated&""&ptwo&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=3 then
returnStr=returnStr&"location.href ='"&alClass&""&Separated&""&ptwo&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=4 then
returnStr=returnStr&"location.href ='"&tvClass&""&Separated&""&ptwo&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=5 then
returnStr=returnStr&"location.href ='"&downClass&""&Separated&""&ptwo&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=6 then
returnStr=returnStr&"location.href ='"&zpClass&""&Separated&""&ptwo&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=7 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if

if lx=8 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if
end if
 '当pone  并且 ptwo 并且 pthree 不为空时
if pone<>"" and ptwo<>"" and pthree<>"" then
if lx=1 then
returnStr=returnStr&"location.href ='"&Newclass&""&Separated&""&pthree&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=2 then
returnStr=returnStr&"location.href ='"&ProductClass&""&Separated&""&pthree&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=3 then
returnStr=returnStr&"location.href ='"&alClass&""&Separated&""&pthree&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=4 then
returnStr=returnStr&"location.href ='"&tvClass&""&Separated&""&pthree&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=5 then
returnStr=returnStr&"location.href ='"&downClass&""&Separated&""&pthree&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=6 then
returnStr=returnStr&"location.href ='"&zpClass&""&Separated&""&pthree&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if lx=7 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if

if lx=8 then
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if
end if


end if
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
else
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagewap',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(482)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(483)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(484)&","& vbCrLf
returnStr=returnStr&  "next: "&word(485)&","& vbCrLf
if word(486)<>"" then
returnStr=returnStr&  "last: "&word(486)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(496)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(497)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
end if

'分页改动采用js分页！更新于2015年05.07
'更新人:追梦工作室
INFOlist = returnStr
else
INFOlist = word(0)
end if
end function




'归档列表页



function guidanginfolist(LX,w,l,style,gl,msql,FL,F,FR,TZ,D,D1,CZ,P)
logdate=request.QueryString("logdate")
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
dim y,m
y=year(logdate)
m=month(logdate)

set RSZM=server.createobject("adodb.recordset")
'-----------打开不同类型的数据库

if l=0 then


if lx=1 then
zm="select * from news where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=2 then
zm="select * from product where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
zm="select * from al where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=4 then
zm="select * from zxsp where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=5 then
zm="select * from download where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

else

if pone="" then

if lx=1 then
zm="select * from news where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=2 then
zm="select * from product where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
zm="select * from al where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=4 then
zm="select * from zxsp where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=5 then
zm="select * from download where month(adddate)="&m&" and year(adddate)="&y&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

end if



if pone<>"" then

if lx=1 then
zm="select * from news where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=2 then
zm="select * from product where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
zm="select * from al where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=4 then
zm="select * from zxsp where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=5 then
zm="select * from download where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

end if



if pone<>"" and ptwo<>"" then

if lx=1 then
zm="select * from news where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=2 then
zm="select * from product where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
zm="select * from al where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=4 then
zm="select * from zxsp where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=5 then
zm="select * from download where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' "
		if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

end if


if pone<>"" and ptwo<>"" and pthree<>"" then

if lx=1 then
zm="select * from news where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=2 then
zm="select * from product where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3 then
zm="select * from al where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=4 then
zm="select * from zxsp where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

if lx=5 then
zm="select * from download where month(adddate)="&m&" and year(adddate)="&y&" and zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and Lang='"&Language&"' "
	if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if

end if

















end if





'-----------打开不同模块的类型结束
'--------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""&word(649)&""

end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'--------------调用样式结束
RSZM.open ZM,conn,1,1

If not( rszm.bof and rszm.EOF) Then

if lx=1 then
if w=1 then
rszm.PageSize = newinfo
else
rszm.PageSize = wapnewinfo
end if
end if
if lx=2 then
if w=1 then
rszm.PageSize = productInfo
else
rszm.PageSize = wapproductInfo
end if
end if
if lx=3 then
if w=1 then
rszm.PageSize = alInfo
else
rszm.PageSize = wapalInfo
end if
end if
if lx=4 then
if w=1 then
rszm.PageSize = tvInfo
else
rszm.PageSize = waptvInfo
end if
end if
if lx=5 then
if w=1 then
rszm.PageSize = downInfo
else
rszm.PageSize = wapdownInfo
end if
end if

iCount = rszm.RecordCount 
iPageSize = rszm.PageSize
maxpage = rszm.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rszm.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If 		
For i = 1 To x
'---------所属分类开始
if lx=6 or lx=7 or lx=8 then
else
dim rsc
set rsc=server.CreateObject("adodb.recordset")
if lx=1 then
rsc.open "select * from newsclass where Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc.open "select * from cpclass where Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc.open "select * from alclass where Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc.open "select * from spclass where Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc.open "select * from xzclass where Lang='"&Language&"'",conn,1,1
end if
while not rsc.eof
if rszm("zmtwo")=0  and rsc("id") =rszm("zmone") and rszm("zmthree")=0 then
if lx=1 then
mqh=rsc("m")
if w=1 then
if yy=1 then

ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=2 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=3 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=4 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else

ssfl="<a href="&sitepath&"wap/video/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=5 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
end if


if rszm("zmone")<>0  and  rszm("zmtwo")= rsc("id") and rszm("zmthree")=0 then


set rsc1=server.CreateObject("adodb.recordset")
if lx=1 then
rsc1.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if lx=2 then
rsc1.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc1.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc1.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc1.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if lx=1 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=2 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=3 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=4 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if

if lx=5 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
end if
if rszm("zmone")<>0  and  rszm("zmtwo")<>0  and rszm("zmthree")=rsc("id") then

set rsc3=server.CreateObject("adodb.recordset")
if lx=1 then
rsc3.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc3.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc3.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc3.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc3.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

set rsc4=server.CreateObject("adodb.recordset")
if lx=1 then
rsc4.open "select * from newsclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc4.open "select * from cpclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc4.open "select * from alclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc4.open "select * from spclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc4.open "select * from xzclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if





if lx=1 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=2 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=3 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=4 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
if lx=5 then
mqh=rsc("m")
if w=1 then
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
else
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
end if
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
end if
'LX 模块 所属分类带链接结束

'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束

'定义链接地址路径开始


if lx=1 then
if w=1 then
tempPage=""&sitepath&""&showinfo&"/"
else
tempPage=""&sitepath&"wap/showinfo/"
end if
end if
if lx=2 then
if w=1 then
tempPage=""&sitepath&""&showproduct&"/"
else
tempPage=""&sitepath&"wap/showproduct/"
end if
end if
if lx=3 then
if w=1 then
tempPage=""&sitepath&""&showphoto&"/"
else
tempPage=""&sitepath&"wap/showphoto/"
end if
end if
if lx=4 then
if w=1 then
tempPage=""&sitepath&""&showvideo&"/"
else
tempPage=""&sitepath&"wap/showvideo/"
end if
end if
if lx=5 then
if w=1 then
tempPage=""&sitepath&""&showdown&"/"
else
tempPage=""&sitepath&"wap/showdown/"
end if
end if


'--------------------定义链接地址路径结束

'-------------tags标签定义开始------------------
if lx=1 then
tt=1
end if
if lx=2 then
tt=2
end if
if lx=3 then
tt=3
end if
if lx=4 then
tt=4
end if
if lx=5 then
tt=5
end if

keyword=rszm("keyword")
keyword1=Split(keyword,"|")
keyword2=ubound(keyword1)
for m=0 to keyword2
if keyword2=0 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&"&Language="&Language&""">"&keyword1(0)&"</a>"
end if
end if
if keyword2=1 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
end if
end if
if keyword2=2 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
end if
end if
if keyword2=3 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
end if
end if
if keyword2=4 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
end if
end if
if keyword2=5 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
end if
end if
if keyword2=6 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
end if
end if
if keyword2=7 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
end if
end if
if keyword2=8 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
end if
end if
if keyword2=9 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
end if
end if
if keyword2=10 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
end if
end if
if keyword2=11 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
end if
end if
if keyword2=12 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
end if
end if
if keyword2=13 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
end if
end if
if keyword2=14 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
end if
end if
if keyword2=15 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
end if
end if
if keyword2=16 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
end if
end if
if keyword2=17 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
end if
end if
if keyword2=18 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
end if
end if
if keyword2=19 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
end if
end if
if keyword2=20 then
if w=1 then
m0="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
m20="<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(keyword1(20))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(20)&"</a>"
else
m0="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(0))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(0)&"</a>"
m1="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(1))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(1)&"</a>"
m2="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(2))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(2)&"</a>"
m3="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(3))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(3)&"</a>"
m4="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(4))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(4)&"</a>"
m5="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(5))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(5)&"</a>"
m6="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(6))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(6)&"</a>"
m7="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(7))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(7)&"</a>"
m8="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(8))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(8)&"</a>"
m9="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(9))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(9)&"</a>"
m10="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(10))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(10)&"</a>"
m11="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(11))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(11)&"</a>"
m12="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(12))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(12)&"</a>"
m13="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(13))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(13)&"</a>"
m14="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(14))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(14)&"</a>"
m15="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(15))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(15)&"</a>"
m16="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(16))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(16)&"</a>"
m17="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(17))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(17)&"</a>"
m18="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(18))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(18)&"</a>"
m19="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(19))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(19)&"</a>"
m20="<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(keyword1(20))&"&tt="&tt&"&m="&mqh&"&Language="&Language&""">"&keyword1(20)&"</a>"
end if
end if
if keyword2=0 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&FR&"")
end if
if keyword2=1 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&FR&"")
end if
if keyword2=2 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&FR&"")
end if
if keyword2=3 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&FR&"")
end if
if keyword2=4 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&FR&"")
end if
if keyword2=5 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&FR&"")
end if

if keyword2=6 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&FR&"")
end if
if keyword2=7 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&FR&"")
end if
if keyword2=8 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&FR&"")
end if
if keyword2=9 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&FR&"")
end if
if keyword2=10 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&FR&"")
end if
if keyword2=11 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&FR&"")
end if
if keyword2=12 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&FR&"")
end if
if keyword2=13 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&FR&"")
end if
if keyword2=14 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&FR&"")
end if
if keyword2=15 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&FR&"")
end if
if keyword2=16 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&FR&"")
end if
if keyword2=17 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&FR&"")
end if
if keyword2=18 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&FR&"")
end if
if keyword2=19 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&F&""&m19&""&FR&"")
end if
if keyword2=20 then
tempListStr = replace(tempListStr,"{TAGS标签}",""&FL&""&m0&""&F&""&m1&""&F&""&m2&""&F&""&m3&""&F&""&m4&""&F&""&m5&""&F&""&m6&""&F&""&m7&""&F&""&m8&""&F&""&m9&""&F&""&m10&""&F&""&m11&""&F&""&m12&""&F&""&m13&""&F&""&m14&""&F&""&m15&""&F&""&m16&""&F&""&m17&""&F&""&m18&""&F&""&m19&""&F&""&m20&""&FR&"")
end if
next

'----------------------tags标签结束



tempListStr = replace(tempListStr,"{归档日期}",logdate)


'----------------------新闻模块替换开始
if lx=1 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{新闻ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{是否头条}",rsZM("gg"))
tempListStr = replace(tempListStr,"{是否幻灯}",rsZM("tj"))
tempListStr = replace(tempListStr,"{是否置顶}",rsZM("tt"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("sytj"))
'tempListStr = replace(tempListStr,"{是否头条}",rsZM("ft"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{新闻标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{新闻标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=2 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=2 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{信息来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{新闻作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
tempListStr = replace(tempListStr,"{新闻缩略图地址}",RSZM("SmallImage")) '仅仅图片地址
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage")) '仅仅幻灯图片地
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then

tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if


if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then

if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")<>0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if RSZM("BFNR")="" THEN
tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
ELSE
tempListStr = replace(tempListStr,"{内容简介}",rszm("bfnr"))
END IF
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
END IF
'----------------------新闻模块替换结束

'--------产品模块替换开始
if lx=2 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{产品ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{产品库存}",RSZM("kucun1"))
tempListStr = replace(tempListStr,"{产品销量}",RSZM("sellbuys"))
tempListStr = replace(tempListStr,"{剩余库存}",RSZM("kucun1")-RSZM("sellbuys"))
tempListStr = replace(tempListStr,"{产品编号}",RSZM("Product_Id"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{产品标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{产品标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
tempListStr = replace(tempListStr,"{是否新品}",RSZM("Newproduct"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{产品来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{产品作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{标准价格}",formatNumber2(RSZM("scprice")))
tempListStr = replace(tempListStr,"{优惠价格}",formatNumber2(RSZM("hyprice")))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage")) '仅仅地址
tempListStr = replace(tempListStr,"{幻灯图地址}",RSZM("IcoImage")) '仅仅地址
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then

sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )

else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=3 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=3 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{订购地址}",""&sitepath&"member/ProductBuy/?step=1&id="&RSZM("id")&"&m="&gl&"&Language="&Language&"")
end if

'--------产品模块替换结束

'--------图片模块替换开始
if lx=3 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{图片ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",rsZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{图片标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{图片标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))

tempListStr = replace(tempListStr,"{图片来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{图片作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)


'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=4 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=4 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
end if
'--------图片模块替换结束
'--------视频模块替换开始
if lx=4  then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{单视频地址}",RSZM("spurl"))
tempListStr = replace(tempListStr,"{多视频地址}",RSZM("arrspurl"))
tempListStr = replace(tempListStr,"{wap视频地址}",RSZM("wapspurl"))
tempListStr = replace(tempListStr,"{视频ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{视频标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{视频标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))

tempListStr = replace(tempListStr,"{视频来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{视频作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if



if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=5 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=5 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束

tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
end if

'-----------------视频模块替换结束

'-----------------下载模块替换开始
if lx=5 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{下载ID}",rsZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",rsZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{下载标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{下载标题}",cutstrx(rsZM("title"),TZ))
END IF

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&"images/demo.jpg")
end if

if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if


tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{下载来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{下载作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{适用系统}",RSZM("System"))
tempListStr = replace(tempListStr,"{软件语言}",RSZM("Language"))
tempListStr = replace(tempListStr,"{软件类型}",RSZM("Softclass"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=6 and Lang='"&Language&"'  order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=6 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{下载地址一}",RSZM("DownloadUrl"))
tempListStr = replace(tempListStr,"{下载地址二}",RSZM("DownloadUrl1"))
tempListStr = replace(tempListStr,"{下载地址三}",RSZM("DownloadUrl2"))
tempListStr = replace(tempListStr,"{文件大小}",RSZM("zdy1"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&"images/demo.jpg")
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}"," "&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if





tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

end if
'-----------------下载模块替换结束




returnStr = returnStr & tempListStr
RSZM.movenext
next
rszm.Close
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)

'分页改动采用js分页！更新于2015年05.07
'更新人:追梦工作室
if instr(returnStr,"{分页}") > 0 then
if w=1 then
returnStr = replace(returnStr,"{分页}","<div id=""pagepcgui"" class=""page""></div>")
else
returnStr = replace(returnStr,"{分页}","<div id=""pagewapgui"" class=""page""></div>")
end if
end if 

if w=1 then
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagepcgui',"& vbCrLf
returnStr=returnStr&  "pages: "&maxpage&","& vbCrLf
returnStr=returnStr&  "skin: '"&word(475)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(476)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(477)&","& vbCrLf
returnStr=returnStr&  "next: "&word(478)&","& vbCrLf
if word(479)<>"" then
returnStr=returnStr&  "last: "&word(479)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(480)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(481)&","& vbCrLf
returnStr=returnStr&"curr: "&page&","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
else
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagewapgui',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(482)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(483)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(484)&","& vbCrLf
returnStr=returnStr&  "next: "&word(485)&","& vbCrLf
if word(486)<>"" then
returnStr=returnStr&  "last: "&word(486)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(496)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(497)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
end if

'分页改动采用js分页！更新于2015年05.07
'更新人:追梦工作室
guidangINFOlist = returnStr
else
guidangINFOlist = word(0)
end if
end function


'PC端搜索框调用
function search(lx,ss)

search=search&"<script language=""javascript"">" & vbCrLf
search=search&"function Checkss()" & vbCrLf
search=search&"{" & vbCrLf
search=search&"if (document.form1.key.value==""""){" & vbCrLf
search=search&"alert ('"&word(5)&"');" & vbCrLf
search=search&"document.form1.key.focus();" & vbCrLf
search=search&"return false;" & vbCrLf
search=search&"}" & vbCrLf
search=search&"if (document.form1.tt.value==""""){ " & vbCrLf
search=search&"alert('"&word(6)&"'); " & vbCrLf
search=search&"document.form1.tt.focus(); " & vbCrLf
search=search&"return false;"  & vbCrLf
search=search&"} " & vbCrLf
search=search&"}" & vbCrLf
search=search&"$(document).ready(function(){" & vbCrLf
search=search&"$("".select"").hover(function(){" & vbCrLf
search=search&"$(""#search_all_category"").show();" & vbCrLf
search=search&"},function(){" & vbCrLf
search=search&"$(""#search_all_category"").hide();" & vbCrLf
search=search&"})" & vbCrLf
search=search&"$("".txt"").focus(function(){" & vbCrLf
search=search&"var txt_value=$(this).val()" & vbCrLf
search=search&"$(this).css({color:""#333333""})" & vbCrLf
search=search&"if(txt_value==this.defaultValue){" & vbCrLf
search=search&"$(this).val("""");" & vbCrLf
search=search&"}" & vbCrLf
search=search&"})" & vbCrLf
search=search&"$("".txt"").blur(function(){" & vbCrLf
search=search&"var txt_value=$(this).val()" & vbCrLf
search=search&"if(txt_value==""""){" & vbCrLf
search=search&"$(this).css({color:""#999999""})" & vbCrLf
search=search&"$(this).val(this.defaultValue);" & vbCrLf
search=search&"}" & vbCrLf
search=search&"})" & vbCrLf
search=search&"$("".select a"").click(function(){" & vbCrLf
search=search&"$txt=$(this).text();" & vbCrLf
search=search&"$(""#Show_Category_Name"").text($txt);" & vbCrLf
search=search&"$(""#search_all_category"").hide();" & vbCrLf
search=search&"$(""#vlu"").val($txt);" & vbCrLf
search=search&"})" & vbCrLf
search=search&"});" & vbCrLf
search=search&"</script>" & vbCrLf
search=search&"<div id=""searchbk"">" & vbCrLf
search=search&"<div id=""seach"">" & vbCrLf
search=search&"<form action="""&sitepath&""&ssearch&"/"" method=""get"" onSubmit=""return Checkss()"" name=""form1"">" & vbCrLf

search=search&"<input name=""language"" type=""hidden"" class=""txt"" value="""&language&"""/>" & vbCrLf
if lx=1 or lx=2 then
search=search&"<input name=""key"" type=""text"" class=""txt"" value="""&word(7)&"""/>" & vbCrLf
end if
if lx=3 then
search=search&"<input name=""key"" type=""text"" class=""txt"" value="""&word(8)&"""/>" & vbCrLf
end if
if lx=4 then
search=search&"<input name=""key"" type=""text"" class=""txt"" value="""&word(9)&"""/>" & vbCrLf
end if
if LX=1 THEN
search=search&"<span class=""select""><span id=""Show_Category_Name"">"&word(10)&"</span>" & vbCrLf
search=search&"<div id=""search_all_category"" style=""display:none;"">" & vbCrLf
search=search&"<a href=""javascript:void(0);""><span>"&word(11)&"</span></a>" & vbCrLf
search=search&"<a href=""javascript:void(0);""><span>"&word(12)&"</span></a>" & vbCrLf
search=search&"<a href=""javascript:void(0);""><span>"&word(13)&"</span></a>" & vbCrLf
search=search&"<a href=""javascript:void(0);""><span>"&word(14)&"</span></a>" & vbCrLf
search=search&"<a href=""javascript:void(0);""><span>"&word(15)&"</span></a>" & vbCrLf
search=search&"</div>" & vbCrLf
search=search&"</span>" & vbCrLf
END IF
if LX=2 THEN
search=search&"<span class=""select""><span id=""Show_Category_Name"">"&word(10)&"</span>" & vbCrLf
search=search&"<div id=""search_all_category"" style=""display:none;"">" & vbCrLf
search=search&"<a href=""javascript:void(0);""><span>"&word(11)&"</span></a>" & vbCrLf
search=search&"<a href=""javascript:void(0);""><span>"&word(12)&"</span></a>" & vbCrLf
search=search&"</div>" & vbCrLf
search=search&"</span>" & vbCrLf
END IF
if LX=3 THEN
search=search&"" & vbCrLf
END IF
if LX=4 THEN
search=search&"" & vbCrLf
END IF
search=search&"<input name="""" type=""submit"" class=""sub"" value="""&ss&"""/>" & vbCrLf
if lx=1 or lx=2 then
search=search&"<input id=""vlu"" type=""hidden"" name=""tt"" value="""">" & vbCrLf
end if
if lx=3 then
search=search&"<input  type=""hidden"" name=""tt"" value=""1"">" & vbCrLf
end if
if lx=4 then
search=search&"<input  type=""hidden"" name=""tt"" value=""2"">" & vbCrLf
end if
search=search&"</form>" & vbCrLf
search=search&"</div>" & vbCrLf
search=search&"</div>" & vbCrLf
end function

'手机端搜索框
function wapsearch(ss)
wapsearch=wapsearch&"<SCRIPT language=""javaScript"">" & vbCrLf
wapsearch=wapsearch&"function Checksearch()" & vbCrLf
wapsearch=wapsearch&"{" & vbCrLf
wapsearch=wapsearch&"if (document.myform.key.value==""""){" & vbCrLf
wapsearch=wapsearch&"alert ("""&word(9)&""");" & vbCrLf
wapsearch=wapsearch&"document.myform.key.focus();" & vbCrLf
wapsearch=wapsearch&"return false;" & vbCrLf
wapsearch=wapsearch&"}" & vbCrLf
wapsearch=wapsearch&"}" & vbCrLf
wapsearch=wapsearch&"</SCRIPT>" & vbCrLf
wapsearch=wapsearch&"<form method=""get"" name=""myform"" action="""&sitepath&"wap/search/"" onSubmit=""return Checksearch()"">" & vbCrLf
wapsearch=wapsearch&"<input type=""text"" name=""key"" class=""text"" placeholder="""&word(9)&""" />" & vbCrLf
wapsearch=wapsearch&"<input type=""submit"" name=""Submit"" value="""&ss&""" class=""submit"" />" & vbCrLf
wapsearch=wapsearch&"<input type=""hidden"" name=""Language"" class=""text"" value="""&language&""" />" & vbCrLf
wapsearch=wapsearch&"<input  type=""hidden"" name=""tt"" value=""2"">" & vbCrLf
wapsearch=wapsearch&"</form>"
end function
tt=request.QueryString("tt") 
key=request.QueryString("key") 
if tt=""&word(11)&"" then
tt=1
end if
if tt=""&word(12)&"" then
tt=2
end if
if tt=""&word(13)&"" then
tt=3
end if
if tt=""&word(14)&"" then
tt=4
end if
if tt=""&word(15)&"" then
tt=5
end if
'-------信息搜索结果显示
function showso(TZ,w,style,msql,gl,D,D1,CZ,L,R,Y,P)

if key<>"" and tt<>"" then
Set rs=Server.CreateObject("ADODB.RecordSet") 
if tt=1 then
sql="select * from hotkey where keywords='"&key&"' and ssfl=1"
end if
if tt=2 then
sql="select * from hotkey where keywords='"&key&"' and ssfl=2"
end if
if tt=3 then
sql="select * from hotkey where keywords='"&key&"' and ssfl=3"
end if
if tt=4 then
sql="select * from hotkey where keywords='"&key&"' and ssfl=4"
end if
if tt=5 then
sql="select * from hotkey where keywords='"&key&"' and ssfl=5"
end if
rs.open sql,conn,1,3
if not (rs.eof and rs.bof) then
rs("hits")=rs("hits")+1
rs("adddate")=now()
rs("sip")=Request.ServerVariables("REMOTE_ADDR")
else
rs.addnew
rs("keywords")=key
rs("adddate")=now()
rs("ssfl")=tt
rs("hits")=1
rs("lang")=language
rs("sip")=Request.ServerVariables("REMOTE_ADDR")
end if
rs.update
rs.close
set rs=nothing


set rszm=server.createobject("adodb.recordset") 
if tt=1 then
zm="select * from [news] where title like '%"&key&"%' and Lang='"&Language&"'  " 
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if tt=2  then
zm="select * from [product] where title like '%"&key&"%' and Lang='"&Language&"'   " 
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if tt=3  then
zm="select * from [al] where title like '%"&key&"%' and Lang='"&Language&"'    " 
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if tt=4  then
zm="select * from [zxsp] where title like '%"&key&"%' and Lang='"&Language&"'   " 
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if tt=5  then
zm="select * from [download] where title like '%"&key&"%' and Lang='"&Language&"'  " 
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
'--------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""

end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'--------------调用样式结束

RSZM.open ZM,conn,1,1
If not( rszm.bof and rszm.EOF) Then
rcount=rszm.RecordCount
returnStr=returnStr& ""&L&""&word(16)&" <span "&Y&">"&key&"</span> "&word(17)&" <span "&Y&">"&rcount&"</span> "&word(18)&""&R&""



IF tt=1 then
RSZM.PageSize = newinfo
end if
IF tt=2 then
RSZM.PageSize = productInfo
end if
IF tt=3 then
RSZM.PageSize = alInfo
end if
IF tt=4 then
RSZM.PageSize = tvInfo
end if
IF tt=5 then
RSZM.PageSize = downInfo
end if

iCount = rszm.RecordCount 
iPageSize = rszm.PageSize
maxpage = rszm.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rszm.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If 		
For i = 1 To x

dim rsc
set rsc=server.CreateObject("adodb.recordset")
if Tt=1 then
rsc.open "select * from newsclass where Lang='"&Language&"'",conn,1,1
end if
if Tt=2 then
rsc.open "select * from cpclass where Lang='"&Language&"'",conn,1,1
end if
if Tt=3 then
rsc.open "select * from alclass where Lang='"&Language&"'",conn,1,1
end if
if Tt=4 then
rsc.open "select * from spclass where Lang='"&Language&"'",conn,1,1
end if
if Tt=5 then
rsc.open "select * from xzclass where Lang='"&Language&"'",conn,1,1
end if
while not rsc.eof
if rszm("zmtwo")=0  and rsc("id") =rszm("zmone") and rszm("zmthree")=0 then
if Tt=1 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/info/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if Tt=2 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/product/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if Tt=3 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/photo/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if


if Tt=4 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/video/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if Tt=5 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/down/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

end if
if rszm("zmone")<>0  and  rszm("zmtwo")= rsc("id") and rszm("zmthree")=0 then

set rsc1=server.CreateObject("adodb.recordset")
if Tt=1 then
rsc1.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if Tt=2 then
rsc1.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=3 then
rsc1.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=4 then
rsc1.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=5 then
rsc1.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if












if Tt=1 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if Tt=2 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if Tt=3 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if Tt=4 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if Tt=5 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

end if


if rszm("zmone")<>0  and  rszm("zmtwo")<>0  and rszm("zmthree")=rsc("id") then



set rsc3=server.CreateObject("adodb.recordset")
if Tt=1 then
rsc3.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=2 then
rsc3.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=3 then
rsc3.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=4 then
rsc3.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=5 then
rsc3.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

set rsc4=server.CreateObject("adodb.recordset")
if Tt=1 then
rsc4.open "select * from newsclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=2 then
rsc4.open "select * from cpclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=3 then
rsc4.open "select * from alclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=4 then
rsc4.open "select * from spclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if Tt=5 then
rsc4.open "select * from xzclass where id="&rszm("zmtwo")&"",conn,1,1
end if

if Tt=1 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&info&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if Tt=2 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if Tt=3 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if Tt=4 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"//"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if Tt=5 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

end if









rsc.movenext
wend
rsc.close
set rsc=nothing

'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListSegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListSttr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,Rr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束

if Tt=1 then
if w=1 then
tempPage=""&sitepath&""&showinfo&"/"
else
tempPage=""&sitepath&"wap/showinfo/"
end if
end if
if Tt=2 then
if w=1 then
tempPage=""&sitepath&""&showproduct&"/"
else
tempPage=""&sitepath&"wap/showproduct/"
end if
end if
if Tt=3 then
if w=1 then
tempPage=""&sitepath&""&showphoto&"/"
else
tempPage=""&sitepath&"wap/showphoto/"
end if
end if
if Tt=4 then
if w=1 then
tempPage=""&sitepath&""&showvideo&"/"
else
tempPage=""&sitepath&"wap/showvideo/"
end if
end if
if Tt=5 then
if w=1 then
tempPage=""&sitepath&""&showdown&"/"
else
tempPage=""&sitepath&"wap/showdown/"
end if
end if
'----------------------新闻模块替换开始
if Tt=1 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{新闻ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{是否头条}",rsZM("gg"))
tempListStr = replace(tempListStr,"{是否幻灯}",rsZM("tj"))
tempListStr = replace(tempListStr,"{是否置顶}",rsZM("tt"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("sytj"))
'tempListStr = replace(tempListStr,"{是否头条}",rsZM("ft"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{新闻缩略图地址}",RSZM("SmallImage")) '仅仅图片地址
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage")) '仅仅幻灯图片地
IF RSZM("BTCOLOR")<>"" THEN
title="<font color="&rszm("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>"
ELSE
title=""&cutstrx(rsZM("title"),TZ)&""
END IF
title=Replace(title, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{新闻标题}",title)
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{信息来源}",RSZM("ly"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=2 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=2 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{新闻作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if






if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if


if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then

if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")<>0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if



content=cutstrx(rsZM("CONTENT"),CZ)
content=Replace(content, key, "<span "&Y&">"&key&"</span>")
bfnr=rsZM("bfnr")
bfnr=Replace(bfnr, key, "<span "&Y&">"&key&"</span>")

if RSZM("BFNR")="" THEN
tempListStr = replace(tempListStr,"{内容简介}",content)
ELSE
tempListStr = replace(tempListStr,"{内容简介}",bfnr)
END IF
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
END IF
'----------------------新闻模块替换结束

'--------产品模块替换开始
if Tt=2 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{产品ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{产品编号}",RSZM("Product_Id"))

IF RSZM("BTCOLOR")<>"" THEN
title="<font color="&rszm("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>"
ELSE
title=""&cutstrx(rsZM("title"),TZ)&""
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
title=Replace(title, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{产品标题}",title)

tempListStr = replace(tempListStr,"{产品颜色}",RSZM("btcolor"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
tempListStr = replace(tempListStr,"{产品库存}",RSZM("kucun1"))
tempListStr = replace(tempListStr,"{产品销量}",RSZM("sellbuys"))
tempListStr = replace(tempListStr,"{剩余库存}",RSZM("kucun1")-RSZM("sellbuys"))
tempListStr = replace(tempListStr,"{是否新品}",RSZM("Newproduct"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage")) '仅仅地址
tempListStr = replace(tempListStr,"{幻灯图地址}",RSZM("IcoImage")) '仅仅地址
tempListStr = replace(tempListStr,"{产品来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{产品作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{标准价格}",formatNumber2(RSZM("scprice")))
tempListStr = replace(tempListStr,"{优惠价格}",formatNumber2(RSZM("hyprice")))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if




if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )

else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if





content=cutstrx(rsZM("CONTENT"),CZ)
content=Replace(content, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{内容简介}",content)
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=3 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=3 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{订购地址}",""&sitepath&"member/ProductBuy/?step=1&id="&RSZM("id")&"&m="&gl&"")
end if

'--------产品模块替换结束

'--------图片模块替换开始
if Tt=3 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{图片ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",rsZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
title="<font color="&rszm("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>"
ELSE
title=""&cutstrx(rsZM("title"),TZ)&""
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
title=Replace(title, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{图片标题}",title)

tempListStr = replace(tempListStr,"{图片来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{图片作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if



if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"")
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if




content=cutstrx(rsZM("CONTENT"),CZ)
content=Replace(content, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{内容简介}",content)
tempListStr = replace(tempListStr,"{所属分类}",ssfl)


'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=4 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=4 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
end if
'--------图片模块替换结束
'--------视频模块替换开始
if Tt=4  then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{单视频地址}",RSZM("spurl"))
tempListStr = replace(tempListStr,"{多视频地址}",RSZM("arrspurl"))
tempListStr = replace(tempListStr,"{wap视频地址}",RSZM("wapspurl"))
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{视频ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
title="<font color="&rszm("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>"
ELSE
title=""&cutstrx(rsZM("title"),TZ)&""
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
title=Replace(title, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{视频标题}",title)

tempListStr = replace(tempListStr,"{视频来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{视频作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=5 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=5 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

content=cutstrx(rsZM("CONTENT"),CZ)
content=Replace(content, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{内容简介}",content)
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

end if

'-----------------视频模块替换结束

'-----------------下载模块替换开始
if Tt=5 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{下载ID}",rsZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))
IF RSZM("BTCOLOR")<>"" THEN
title="<font color="&rszm("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>"
ELSE
title=""&cutstrx(rsZM("title"),TZ)&""
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
title=Replace(title, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{下载标题}",title)

tempListStr = replace(tempListStr,"{下载来源}",RSZM("ly"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=6 &Language="&Language&" order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=6 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{下载作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{适用系统}",RSZM("System"))
tempListStr = replace(tempListStr,"{软件语言}",RSZM("Language"))
tempListStr = replace(tempListStr,"{软件类型}",RSZM("Softclass"))
tempListStr = replace(tempListStr,"{下载地址一}",RSZM("DownloadUrl"))
tempListStr = replace(tempListStr,"{下载地址二}",RSZM("DownloadUrl1"))
tempListStr = replace(tempListStr,"{下载地址三}",RSZM("DownloadUrl2"))
tempListStr = replace(tempListStr,"{文件大小}",RSZM("zdy1"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&"images/demo.jpg")
end if

if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if



if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&SitePath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



content=cutstrx(rsZM("CONTENT"),CZ)
content=Replace(content, key, "<span "&Y&">"&key&"</span>")
tempListStr = replace(tempListStr,"{内容简介}",content)
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

end if
'-----------------下载模块替换结束


returnStr = returnStr & tempListStr
RSZM.movenext
next

rszm.Close
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)

if instr(returnStr,"{分页}") > 0 then
if w=1 then
returnStr = replace(returnStr,"{分页}","<div id=""pagepc"" class=""page""></div>")
else
returnStr = replace(returnStr,"{分页}","<div id=""pagewap"" class=""page""></div>")
end if
end if 

if w=1 then
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagepc',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(475)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(476)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(477)&","& vbCrLf
returnStr=returnStr&  "next: "&word(478)&","& vbCrLf
if word(479)<>"" then
returnStr=returnStr&  "last: "&word(479)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(480)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(481)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
else
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagewap',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(482)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(483)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(484)&","& vbCrLf
returnStr=returnStr&  "next: "&word(485)&","& vbCrLf
if word(486)<>"" then
returnStr=returnStr&  "last: "&word(486)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(496)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(497)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
end if


showso = returnStr

else
ero="<div style=""text-align:center;"">"&word(19)&"<span "&Y&"> "&key&" </span>"&word(20)&"</div>"
showso = ero
end if



'---------搜索结果显示结束

end if
end function

'--------订阅emaill框
function demaill(D,TX,T,U)
demaill=demaill&"<div id=""emaill"">" & vbCrLf
demaill=demaill&"<form name=""youjian"" method=""post"" action=""#"">" & vbCrLf
demaill=demaill&"<input name=""emaill"" type=""text"" id=ymail class=""txt""  onfocus=""this.value=''"" value="""&word(21)&""" size=""21"">" & vbCrLf
demaill=demaill&"<INPUT name=Submit2 onclick=""DoMail('1',this.form.ymail.value,'"&U&"');"" type=button value="&d&" class=""sub"">" & vbCrLf
if TX=1 THEN
demaill=demaill&"<input name=Submit4 onclick=""DoMail('0',this.form.ymail.value,'"&U&"');"" type=button value="&t&" class=""sub1"" />" & vbCrLf
end if
demaill=demaill&"</div>" & vbCrLf
demaill=demaill&"<SCRIPT>" & vbCrLf
demaill=demaill&"function checkEmail(str)" & vbCrLf
demaill=demaill&"{" & vbCrLf
demaill=demaill&"if(str == """")" & vbCrLf
demaill=demaill&"return false;" & vbCrLf
demaill=demaill&"if (str.charAt(0) == ""."" || str.charAt(0) == ""@"" || str.indexOf('@', 0) == -1" & vbCrLf
demaill=demaill&"   || str.indexOf('.', 0) == -1 || str.lastIndexOf(""@"") == str.length-1 || str.lastIndexOf(""."") == str.length-1)" & vbCrLf
demaill=demaill&"return false;" & vbCrLf
demaill=demaill&"else" & vbCrLf
demaill=demaill&"return true;" & vbCrLf
demaill=demaill&"}" & vbCrLf
demaill=demaill&"function DoMail(ntype,mail,U)" & vbCrLf
demaill=demaill&"{" & vbCrLf
demaill=demaill&" if(!checkEmail(mail))" & vbCrLf
demaill=demaill&"{" & vbCrLf
demaill=demaill&"alert("""&word(22)&""");" & vbCrLf
demaill=demaill&"return;" & vbCrLf
demaill=demaill&"}" & vbCrLf
demaill=demaill&"window.open("""&sitepath&"inc/yjdy/?Do=""+ntype+""&U=""+U+""&mail=""+mail+""&Language="&Language&""","""","""");" & vbCrLf
demaill=demaill&"}" & vbCrLf
demaill=demaill&"</SCRIPT>" & vbCrLf
demaill=demaill&"</form>" & vbCrLf
end function
'订单查询
function ddcx(cx,m)
ddcx=ddcx&"<div id=""dd"">" & vbCrLf
ddcx=ddcx&"<form action="""&sitepath&""&orderso&"/"" method=""get"">" & vbCrLf
ddcx=ddcx&"<input name=""ddkey"" type=""text"" class=""txt"" value="""&word(23)&"""/>" & vbCrLf
ddcx=ddcx&"<input name=""language"" type=""hidden"" class=""txt"" value='"&language&"'/>" & vbCrLf
ddcx=ddcx&"<input name=""m"" type=""hidden"" class=""txt"" value='"&m&"'/>" & vbCrLf
ddcx=ddcx&"<input name="""" type=""submit"" class=""sub"" value="""&cx&"""/>" & vbCrLf
ddcx=ddcx&"</form> "    & vbCrLf       
ddcx=ddcx&"</div>" & vbCrLf
end function
ddkey=request.QueryString("ddkey") 
'显示订单结果
function showddso(style,msql,D,D1,p)
call initCodecs 
if ddkey=""&word(23)&"" or ddkey="" then
writemsg1 1,""&word(23)&"","","wrong"
Response.End()
end if
if ddkey<>"" then
set rs=server.createobject("adodb.recordset") 
sql="select * from orderdd where orderddh='"&ddkey&"' and Lang='"&Language&"'   " 


if msql<>"no" then
sql=sql&" and "&msql&""
end if
sql=sql&" order by "&p&""

'----------------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'----------------------调用样式结束


rs.open sql,conn,1,1 
If not( rs.bof and rs.EOF) Then
rs.PageSize = sodd
iCount = rs.RecordCount 
iPageSize = rs.PageSize
maxpage = rs.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rs.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If	



         
			For i = 1 To x
			
			
			'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束
			  if rs("isdate")=0 then
			  zt=""&word(30)&"" 
			  end if
			  if rs("isdate")=1 then  
			  zt=""&word(31)&"" 
			  end if
			  if rs("isdate")=2 then  
			  zt=""&word(32)&"" 
			  end if
			  if rs("isdate")=3 then  
			  zt=""&word(33)&"" 
			  end if
			  if rs("isdate")=5 then  
			  zt=""&word(34)&"" 
			  end if	
			
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rs("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)

tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{性别}",rs("Sex"))
tempListStr = replace(tempListStr,"{单位}",rs("Company"))
tempListStr = replace(tempListStr,"{地址}",rs("Address"))
tempListStr = replace(tempListStr,"{邮编}",rs("ZipCode"))
tempListStr = replace(tempListStr,"{传真}",rs("Fax"))
tempListStr = replace(tempListStr,"{邮箱}",rs("Email"))
tempListStr = replace(tempListStr,"{订单号}",rs("orderddh"))
tempListStr = replace(tempListStr,"{订购人}",rs("Linkman"))
tempListStr = replace(tempListStr,"{联系电话}",rs("Linkman"))
tempListStr = replace(tempListStr,"{订购时间}",FormatTime(rs("Addtime"),D))
tempListStr = replace(tempListStr,"{订购时间1}",FormatTime(rs("Addtime"),D1))
tempListStr = replace(tempListStr,"{订购状态}",zt)


if rs("zffs")=1 then 
zs=""&word(269)&""
end if
if rs("zffs")=2 then 
zs=""&word(270)&""
end if
tempListStr = replace(tempListStr,"{支付方式}",zs)
if rs("ReplyTime")<>"" then	
tempListStr = replace(tempListStr,"{回复时间}",rs("ReplyTime"))	
else
tempListStr = replace(tempListStr,"{回复时间}",""&word(24)&"")	
end if

if rs("ReplyContent")<>"" then	
tempListStr = replace(tempListStr,"{回复内容}",rs("ReplyContent"))
else
tempListStr = replace(tempListStr,"{回复内容}",""&word(24)&"")	
end if
tempListStr = replace(tempListStr,"{配送方式}",rs("shipmentName"))
tempListStr = replace(tempListStr,"{备注}",rs("Remark"))
tempListStr = replace(tempListStr,"{配送费用}",formatNumber2(rs("shipmentMoney")))
tempListStr = replace(tempListStr,"{产品费用}",formatNumber2(rs("aumont")))
tempListStr = replace(tempListStr,"{支付总额}",formatNumber2(rs("shipmentMoney")+rs("aumont")))
tempListStr = replace(tempListStr,"{支付链接}",""&template&"member/pay.asp?orderddh="&base64Encode(rs("orderddh"))&"&m="&m1&"")	


returnStr = returnStr & tempListStr	
   		
rs.movenext
Next
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)

else
writemsg1 1 ,""&word(41)&"","","wrong"
end if
rs.close
set rs=nothing
end if
showddso = returnStr
end function








'单页内容调用开始
aboutid=request.QueryString("aboutid")
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
page=request.QueryString("page")
Set rsabout=Server.CreateObject("ADODB.RecordSet") 
sql="select * from dy where id="&aboutid
rsabout.Open sql,conn,1,1
content=UBBCode(FormatImg(rsabout("content")))
content1=UBBCode(rsabout("content"))
zdfy=rsabout("zdfy")
zmone=rsabout("zmone")
id=rsabout("id")
'自定义字段开始
function Z_A(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=1 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=1 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rsabout(""&zdyzd1&"")
Z_A=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束
function A(C)
Select Case C
case "单页ID"
A=rsabout("ID")
case "单页标题"
A=rsabout("title")
case "标题页数"
A=Page
case "单页关键字"
A=rsabout("gjz")
case "单页描述"
A=rsabout("ms")
case "单页图片组"
pic=rsabout("ImagePath")
pic1=Split(pic,"|")
A=pic1
case "阅读权限"
A=ViewNoRight(rsabout("GroupID"),rsabout("Exclusive"))
case "所属分类"
if pone<>"" then
set lm=server.createobject("adodb.recordset") 
exec="select * from [mnclass] where id="&pone&" and Lang='"&Language&"'"
lm.open exec,conn,1,1
A=lm("ClassName")
lm.close
set lm=nothing
end if
'更新于20150507
case "单页内容"
A=ManualPagination(id,content,1)
case "wap单页内容"
A=mobileManualPagination(id,content) 
 '更新于20150507
case "发布时间"
A=rsabout("adddate")
end  Select
end  function
'新闻内容调用开始
newsid=request.QueryString("newsid")
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
sql="select * from news where id="&newsid&""
rsnews.Open sql,conn,1,1
content=UBBCode(FormatImg(rsnews("content")))
content1=UBBCode(rsnews("content"))
zdfy=rsnews("zdfy")
zmone=rsnews("zmone")
zmtwo=rsnews("zmtwo")
zmthree=rsnews("zmthree")
id=rsnews("ID")




'自定义字段开始
function Z_N(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=2 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=2 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rsnews(""&zdyzd1&"")
Z_N=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束
function N(C)
Select Case C
case "新闻标题"
N=rsnews("title")
case "标题颜色"
N=rsnews("btcolor")
case "发布时间"
N=rsnews("adddate")
case "新闻来源"
N=rsnews("ly")
case "标题页数"
N=Page
case "新闻图片组"
pic=rsnews("ImagePath")
pic1=Split(pic,"|")
N=pic1
case "阅读权限"
N=ViewNoRight(rsnews("GroupID"),rsnews("Exclusive"))
case "新闻作者"
N=rsnews("user")
case "新闻关键词"
N=rsnews("newkey")
case "新闻描述"
N=rsnews("newms")
case "新闻简介"
N=rsnews("bfnr")
case "新闻内容"
N=ManualPagination(id,content,2)
case "所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from newsclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
if yy=1 then
N="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
N="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from newsclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
N="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
N="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc1.close
set rsc1=nothing
end if
if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from newsclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from newsclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
N="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
N="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
case "WAP所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from newsclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
N="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if

if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from newsclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
N="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc1.close
set rsc1=nothing
end if

if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from newsclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from newsclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
N="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
case "wap新闻内容"
N=mobileManualPagination(id,content)
case "WAP下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id>" & request.QueryString("newsid") &" and adddate<=now() and Lang='"&Language&"'  order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id>" & request.QueryString("newsid") &" and adddate<=getdate() and Lang='"&Language&"'   order by id asc ",conn,1,1
end if
if not rstmp.eof then




if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if











if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if

else
nexttitle = ""&word(42)&""
end if
N=nexttitle
rstmp.close
set rstmp=nothing

case "WAP上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id<" & request.QueryString("newsid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id<" & request.QueryString("newsid") & " and adddate<=getdate() and Lang='"&Language&"' order by id desc ",conn,1,1
end if
if not rstmp.eof then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
N=prevtitle
rstmp.close
set rstmp=nothing

case "下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id>" & request.QueryString("newsid") &" and adddate<=now() and Lang='"&Language&"'   order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id>" & request.QueryString("newsid") &" and adddate<=getdate() and Lang='"&Language&"'   order by id asc ",conn,1,1
end if
if not rstmp.eof then



if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if









if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle="<a href="&NewName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>"&rstmp("title")&"</a>"
end if
else
nexttitle = ""&word(42)&""
end if
N=nexttitle
rstmp.close
set rstmp=nothing

case "上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id<" & request.QueryString("newsid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from news where id<" & request.QueryString("newsid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from newsclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if



if yy=1 then

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&newsid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if

else
prevtitle="<a href="&NewName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
N=prevtitle
rstmp.close
set rstmp=nothing
case "点击次数"

N="<script language=""javascript"" src="""&SitePath&"include/HitCount.asp?id="&newsid&"&LX=news""></script><script language=""javascript"" src="""&SitePath&"include/HitCount.asp?tj=yes&LX=news&id="&newsid&"""></script>"

end  Select
end  function

function Ntags(TQZ,lx,l)
If rsnews("KeyWord")<>"" then
Ntags=Ntags&"<div class=""tags"">"
Ntags=Ntags&"<ul>"
Ntags=Ntags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsnews("KeyWord")), "|")
For i=0 to Ubound(aa)
Ntags=Ntags&"<li>"
Ntags=Ntags&"<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(aa(i))&"&tt=1&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a>"
if lx=1 then
if i<> Ubound(aa) then 
Ntags=Ntags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ntags=Ntags&"</li>"
Next
Ntags=Ntags&"</ul></div>"
End if
end function




function wapNtags(TQZ,lx,l)
If rsnews("KeyWord")<>"" then
Ntags=Ntags&"<div class=""waptags"">"
Ntags=Ntags&"<ul>"
Ntags=Ntags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsnews("KeyWord")), "|")
For i=0 to Ubound(aa)
Ntags=Ntags&"<li>"
Ntags=Ntags&"<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(aa(i))&"&tt=1&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a>"
if lx=1 then
if i<> Ubound(aa) then 
Ntags=Ntags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ntags=Ntags&"</li>"
Next
Ntags=Ntags&"</ul></div>"
End if
end function









'产品内容调用开始
proid=request.QueryString("proid")
Set rspro=Server.CreateObject("ADODB.RecordSet") 
sql="select * from product where id="&proid
rspro.Open sql,conn,1,1
content=UBBCode(FormatImg(rspro("content")))
content1=UBBCode(rspro("content"))
zdfy=rspro("zdfy")
zmone=rspro("zmone")
zmtwo=rspro("zmtwo")
zmthree=rspro("zmthree")
groupid=rspro("groupid")
id=rspro("ID")
'自定义字段开始
function Z_P(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=3 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=3 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rspro(""&zdyzd1&"")
Z_P=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束
function P(C)
Select Case C
case "产品标题"
P=rspro("title")
case "标题颜色"
P=rspro("btcolor")
case "发布时间"
P=rspro("adddate")
case "产品来源"
P=rspro("ly")
case "产品库存"
P=rspro("kucun1")
case "产品销量"
P=rspro("sellbuys")
case "剩余库存"
P=rspro("kucun1")-rspro("sellbuys")
case "标题页数"

P=Page

case "所属分类"

set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from cpclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
if yy=1 then
P="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
P="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from cpclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
P="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
P="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc1.close
set rsc1=nothing
end if

if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from cpclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from cpclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
P="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
P="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if


rsc4.close
set rsc4=nothing

rsc3.close
set rsc3=nothing

end if

rsc.movenext
wend
rsc.close
set rsc=nothing

case "WAP所属分类"

set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from cpclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
response.write("<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&" >" & rsc("classname") & "</a>")
end if

if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from cpclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
P="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc1.close
set rsc1=nothing
end if
if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from cpclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from cpclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
P="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing



case "市场价格"
If Trim(RsPro("P_SpecificationName")) <> "" And Not isnull(Trim(RsPro("P_SpecificationName"))) Then 
								Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price asc")
								If Not rsprice.eof Then 
										Lowprice = formatnumber2(rspro("scPrice") + rsprice(0))
								End If 
								rsprice.close : Set rs = Nothing
								Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price desc")
								If Not rsprice.eof Then 
										Highprice = formatnumber2(rspro("scPrice") + rsprice(0))
								End If 
								rsprice.close : Set rs = Nothing
								
								If CDbl(Highprice) = CDbl(Lowprice) Then 
							
										p=lowprice
								Else
										P=""&lowprice&"-"&Highprice&""
								End If 
								else
								P=formatnumber2(rspro("scPrice"))

end if





case "会员价格"
If Trim(RsPro("P_SpecificationName")) <> "" And Not isnull(Trim(RsPro("P_SpecificationName"))) Then 
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price asc")
If Not rsprice.eof Then 
Lowprice = formatnumber2(rspro("hyPrice") + rsprice(0))
End If 
rsprice.close : Set rs = Nothing
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price desc")
If Not rsprice.eof Then 
Highprice =formatnumber2( rspro("hyPrice") + rsprice(0))
End If 
rsprice.close : Set rs = Nothing
If CDbl(Highprice) = CDbl(Lowprice) Then 
p=lowprice
Else
P=""&lowprice&"-"&Highprice&""
End If 															
else
set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&proid&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price asc")
If Not rsprice.eof Then 
Lowprice = formatnumber2(rsjg1("Price") + rsprice(0))
End If 
rsprice.close : Set rs = Nothing
Set rsprice = conn.execute("select top 1 p_price from [Product_Stock] where p_id = "& proid &" and Lang='"&Language&"' order by p_price desc")
If Not rsprice.eof Then 
Highprice =formatnumber2( rsjg1("Price") + rsprice(0) )
End If 
rsprice.close : Set rs = Nothing
If CDbl(Highprice) = CDbl(Lowprice) Then 
p=lowprice
Else
P=""&lowprice&"-"&Highprice&""
End If 

end if						
else
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed"   then
P=formatnumber2(rspro("hyPrice"))
else
set rsjg=server.CreateObject("adodb.recordset")
sqljg="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg.open sqljg,conn,1,1
set rsjg1=server.CreateObject("adodb.recordset")
sqljg1="select * from product_price where grade_id="&rsjg("GroupLevel")&" and product_id="&proid&" and Lang='"&Language&"'"
rsjg1.open sqljg1,conn,1,1
P=formatnumber2(rsjg1("Price"))
end if
end if
case "产品图片组"
pic=rspro("ImagePath")
pic1=Split(pic,"|")
P=pic1
case "阅读权限"
P=ViewNoRight(rspro("GroupID"),rspro("Exclusive"))
case "产品编号"
P=rspro("Product_Id")

case "产品作者"
P=rspro("user")
case "产品关键词"
P=rspro("productkey")
case "产品描述"
P=rspro("productms")

case "权限值"
if rspro("Exclusive") = ">=" then
P="隶属于"
end if
if rspro("Exclusive")= "=" then
p="专属于"
end if
case "产品内容"
P=ManualPagination(id,content,3)
case "wap产品内容"
P=mobileManualPagination(id,nohtml(content))
case "WAP下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id>" & request.QueryString("proid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id>" & request.QueryString("proid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then


if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if


if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&proid=" & rstmp("id") & "&m="&guojuan&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&proid=" & rstmp("id") & "&m="&guojuan&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&proid=" & rstmp("id") & "&m="&guojuan&""">" & rstmp("title") & "</a>"
end if
else
nexttitle = ""&word(42)&""
end if
P=nexttitle
rstmp.close
set rstmp=nothing
%>
<%
case "WAP上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id<" & request.QueryString("proid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id<" & request.QueryString("proid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'<p></p>",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
P=prevtitle
rstmp.close
set rstmp=nothing

case "下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id>" & request.QueryString("proid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id>" & request.QueryString("proid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle="<a href="&productName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>"&rstmp("title")&"</a>"
end if
else
nexttitle = ""&word(42)&""
end if
P=nexttitle
rstmp.close
set rstmp=nothing
case "上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id<" & request.QueryString("proid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from product where id<" & request.QueryString("proid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from cpclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&proid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle="<a href="&productName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
P=prevtitle
rstmp.close
set rstmp=nothing
case "点击次数"
P="<script language=""javascript"" src="""&SitePath&"include/HitCount.asp?id="&proid&"&LX=product""></script><script language=""javascript"" src="""&SitePath&"include/HitCount.asp?tj=yes&LX=product&id="&proid&"""></script>"
end  Select
end  function

function Ptags(TQZ,lx,l)
If rspro("KeyWord")<>"" then

Ptags=Ptags&"<div class=""tags"">"
Ptags=Ptags&"<ul>"
Ptags=Ptags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rspro("KeyWord")), "|")
For i=0 to Ubound(aa)
Ptags=Ptags&"<li>"
Ptags=Ptags&"<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(aa(i))&"&tt=2&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a>"
if lx=1 then
if i<> Ubound(aa) then
Ptags=Ptags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ptags=Ptags&"</li>"
Next
Ptags=Ptags&"</ul></div>"
End if

end function


function wapPtags(TQZ,lx,l)
If rspro("KeyWord")<>"" then

Ptags=Ptags&"<div class=""waptags"">"
Ptags=Ptags&"<ul>"
Ptags=Ptags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rspro("KeyWord")), "|")
For i=0 to Ubound(aa)
Ptags=Ptags&"<li>"
Ptags=Ptags&"<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(aa(i))&"&tt=2&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a>"
if lx=1 then
if i<> Ubound(aa) then
Ptags=Ptags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ptags=Ptags&"</li>"
Next
Ptags=Ptags&"</ul></div>"
End if

end function
dim Contentsc(50)
cont=UBBCode(rspro("content"))
content1=split(cont,"{选项卡标记}")
i=-1
zd=ubound(content)
For Each Strs In content1
i=i+1
Contentsc(i)=content1(i)
Contentsc_cont=Contentsc_cont&"<div>"&Contentsc(i)&"</div>"
content_title=content_title&"<li><a href='#con"&i+1&"'><span>"&FilterHTML(Retitle(Contentsc(i)))&"</span></a></li>"
next
cont0="<div id=""con1"">"&Contentsc(0)&"</div>"
if Contentsc(1)<>""then cont1="<div id=""con2"">"&Deltitle(Contentsc(1))&"</div>" 
if Contentsc(2)<>""then cont2="<div id=""con3"">"&Deltitle(Contentsc(2))&"</div>"
if Contentsc(3)<>""then cont3="<div id=""con4"">"&Deltitle(Contentsc(3))&"</div>"
if Contentsc(4)<>""then cont4="<div id=""con5"">"&Deltitle(Contentsc(4))&"</div>"
if Contentsc(5)<>""then cont5="<div id=""con6"">"&Deltitle(Contentsc(5))&"</div>"
if Contentsc(6)<>""then cont6="<div id=""con7"">"&Deltitle(Contentsc(6))&"</div>"
if Contentsc(7)<>""then cont7="<div id=""con8"">"&Deltitle(Contentsc(7))&"</div>"
if Contentsc(8)<>""then cont8="<div id=""con9"">"&Deltitle(Contentsc(8))&"</div>"
if Contentsc(9)<>""then cont9="<div id=""con10"">"&Deltitle(Contentsc(9))&"</div>"
if Contentsc(10)<>""then cont10="<div id=""con11"">"&Deltitle(Contentsc(10))&"</div>"
content2=content2&cont0&cont1&cont2&cont3&cont4&cont5&cont6&cont7&cont8&cont9&cont10
content2=nohtml1(content2)

'图片内容调用开始
phid=request.QueryString("phid")
Set rsph=Server.CreateObject("ADODB.RecordSet") 
sql="select * from al where id="&phid
rsph.Open sql,conn,1,1
content=UBBCode(FormatImg(rsph("content")))
content1=UBBCode(rsph("content"))
zdfy=rsph("zdfy")
zmone=rsph("zmone")
zmtwo=rsph("zmtwo")
zmthree=rsph("zmthree")
id=rsph("ID")
'自定义字段开始
function Z_C(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=4 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=4 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rsph(""&zdyzd1&"")
Z_C=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束
function C(D)
Select Case D
case "图片标题"
C=rsph("title")
case "标题颜色"
C=rsph("btcolor")
case "发布时间"
C=rsph("adddate")
case "图片来源"
C=rsph("ly")
case "标题页数"

C=Page

case "图片组"
pic=rsph("ImagePath")
pic1=Split(pic,"|")
C=pic1
case "阅读权限"
C=ViewNoRight(rsph("GroupID"),rsph("Exclusive"))
case "图片作者"
C=rsph("user")
case "图片关键词"
C=rsph("alkey")
case "图片描述"
C=rsph("alms")
case "图片内容"
C=ManualPagination(id,content,4)
case "wap图片内容"
C=mobileManualPagination(id,content)
case "WAP下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id>" & request.QueryString("phid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id>" & request.QueryString("phid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then


if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if









if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle = ""&word(42)&""
end if
C=nexttitle
rstmp.close
set rstmp=nothing
case "WAP上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id<" & request.QueryString("phid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id<" & request.QueryString("phid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then


if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if





if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
C=prevtitle
rstmp.close
set rstmp=nothing
case "下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id>" & request.QueryString("phid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id>" & request.QueryString("phid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then


if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if











if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle="<a href="&alName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>"&rstmp("title")&"</a>"
end if
else
nexttitle = ""&word(42)&""
end if
C=nexttitle
rstmp.close
set rstmp=nothing
case "上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id<" & request.QueryString("phid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from al where id<" & request.QueryString("phid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from alclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if




if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&phid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle="<a href="&alName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
C=prevtitle
rstmp.close
set rstmp=nothing
case "所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from alclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
if yy=1 then
C="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
C="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if

if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from alclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
C="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
C="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc1.close
set rsc1=nothing
end if

if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from alclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from alclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
C="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
C="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
case "WAP所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from alclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
C="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from alclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
C="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc1.close
set rsc1=nothing
end if
if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from alclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from alclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
C="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
case "点击次数"
C="<script language=""javascript"" src="""&SitePath&"include/HitCount.asp?id="&phid&"&LX=al""></script><script language=""javascript"" src="""&SitePath&"include/HitCount.asp?tj=yes&LX=al&id="&phid&"""></script>"
end  Select
end  function
function ctags(TQZ,lx,l)
If rsph("KeyWord")<>"" then

ctags=ctags&"<div class=""tags"">"
ctags=ctags&"<ul>"
ctags=ctags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsph("KeyWord")), "|")
For i=0 to Ubound(aa)
ctags=ctags&"<li>"
ctags=ctags&"<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(aa(i))&"&tt=3&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a>" 
if lx=1 then
if i<> Ubound(aa) then
ctags=ctags&"<span class=""fgx"">"&l&"</span>"
end if
end if
ctags=ctags&"</li>"
Next
ctags=ctags&"</ul></div>"
End if

end function

function wapctags(TQZ,lx,l)
If rsph("KeyWord")<>"" then

ctags=ctags&"<div class=""waptags"">"
ctags=ctags&"<ul>"
ctags=ctags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsph("KeyWord")), "|")
For i=0 to Ubound(aa)
ctags=ctags&"<li>"
ctags=ctags&"<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(aa(i))&"&tt=3&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a>" 
if lx=1 then
if i<> Ubound(aa) then
ctags=ctags&"<span class=""fgx"">"&l&"</span>"
end if
end if
ctags=ctags&"</li>"
Next
ctags=ctags&"</ul></div>"
End if

end function
'视频内容调用开始
videoid=request.QueryString("videoid")
Set rsvideo=Server.CreateObject("ADODB.RecordSet") 
sql="select * from zxsp where id="&videoid
rsvideo.Open sql,conn,1,1
content=UBBCode(FormatImg(rsvideo("content")))
content1=UBBCode(rsvideo("content"))
zdfy=rsvideo("zdfy")
zmone=rsvideo("zmone")
zmtwo=rsvideo("zmtwo")
zmthree=rsvideo("zmthree")
id=rsvideo("ID")
'自定义字段开始
function Z_T(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=5 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=5 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rsvideo(""&zdyzd1&"")
Z_T=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束
function T(D)
Select Case D
case "视频标题"
T=rsvideo("title")
case "标题颜色"
T=rsvideo("btcolor")
case "发布时间"
T=rsvideo("adddate")
case "视频来源"
T=rsvideo("ly")
case "标题页数"
T=Page
case "视频图片组"
pic=rsvideo("ImagePath")
pic1=Split(pic,"|")
T=pic1
case "阅读权限"
T=ViewNoRight(rsvideo("GroupID"),rsvideo("Exclusive"))
case "视频作者"
T=rsvideo("user")
case "视频关键词"
T=rsvideo("tvkey")
case "视频描述"
T=rsvideo("tvms")
case "单视频地址"
T=rsvideo("spurl")
case "多视频地址"
T=rsvideo("arrspurl")
case "WAP视频地址"
T=rsvideo("wapspurl")
case "视频内容"
T=ManualPagination(id,content,5)
case "wap视频内容"
T=mobileManualPagination(id,content)
case "所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from spclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
if yy=1 then
T="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
T="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from spclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
T="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
T="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc1.close
set rsc1=nothing
end if
if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from spclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from spclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
T="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
T="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
case "WAP所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from spclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
T="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if
if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from spclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
T="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc1.close
set rsc1=nothing
end if
if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from spclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from spclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
T="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing

case "WAP下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id>" & request.QueryString("videoid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id>" & request.QueryString("videoid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle = ""&word(42)&""
end if
T=nexttitle
rstmp.close
set rstmp=nothing
case "WAP上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id<" & request.QueryString("videoid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id<" & request.QueryString("videoid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if




if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
T=prevtitle
rstmp.close
set rstmp=nothing
case "下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id>" & request.QueryString("videoid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id>" & request.QueryString("videoid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if

if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle="<a href="&tvName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>"&rstmp("title")&"</a>"
end if
else
nexttitle = ""&word(42)&""
end if
T=nexttitle
rstmp.close
set rstmp=nothing
case "上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id<" & request.QueryString("videoid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from zxsp where id<" & request.QueryString("videoid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from spclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&videoid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle="<a href="&tvName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
T=prevtitle
rstmp.close
set rstmp=nothing
case "点击次数"
T="<script language=""javascript"" src="""&SitePath&"include/HitCount.asp?id="&videoid&"&LX=zxsp""></script><script language=""javascript"" src="""&SitePath&"include/HitCount.asp?tj=yes&LX=zxsp&id="&videoid&"""></script>"
end  Select
end  function
function Ttags(TQZ,lx,l)
If rsvideo("KeyWord")<>"" then

Ttags=Ttags&"<div class=""tags"">"
Ttags=Ttags&"<ul>"
Ttags=Ttags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsvideo("KeyWord")), "|")
For i=0 to Ubound(aa)
Ttags=Ttags&"<li>"
Ttags=Ttags&"<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(aa(i))&"&tt=4&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a> "
if lx=1 then
if i<> Ubound(aa) then
Ttags=Ttags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ttags=Ttags&"</li>"
Next
Ttags=Ttags&"</ul></div>"
End if

end function



function wapTtags(TQZ,lx,l)
If rsvideo("KeyWord")<>"" then

Ttags=Ttags&"<div class=""waptags"">"
Ttags=Ttags&"<ul>"
Ttags=Ttags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsvideo("KeyWord")), "|")
For i=0 to Ubound(aa)
Ttags=Ttags&"<li>"
Ttags=Ttags&"<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(aa(i))&"&tt=4&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a> "
if lx=1 then
if i<> Ubound(aa) then
Ttags=Ttags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ttags=Ttags&"</li>"
Next
Ttags=Ttags&"</ul></div>"
End if

end function
'下载内容调用开始
downid=request.QueryString("downid")
Set rsdown=Server.CreateObject("ADODB.RecordSet") 
sql="select * from download where id="&downid
rsdown.Open sql,conn,1,1
content=UBBCode(FormatImg(rsdown("content")))
content1=UBBCode(rsdown("content"))
zdfy=rsdown("zdfy")
zmone=rsdown("zmone")
zmtwo=rsdown("zmtwo")
zmthree=rsdown("zmthree")
id=rsdown("ID")
'自定义字段开始
function Z_Z(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=6 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=6 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rsdown(""&zdyzd1&"")
Z_Z=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束
function Z(D)
Select Case D
case "下载标题"
Z=rsdown("title")
case "标题颜色"
Z=rsdown("btcolor")
case "发布时间"
Z=rsdown("adddate")
case "下载来源"
Z=rsdown("ly")
if page<>1 then
end if
case "下载图片组"
pic=rsdown("ImagePath")
pic1=Split(pic,"|")
Z=pic1
case "阅读权限"
Z=ViewNoRight(rsdown("GroupID"),rsdown("Exclusive"))
case "发布作者"
Z=rsdown("user")
case "下载关键词"
Z=rsdown("downkey")
case "下载描述"
Z=rsdown("downms")
case "适用系统"
Z=rsdown("system")
case "软件语言"
Z=rsdown("Language")
case "软件类型"
Z=rsdown("Softclass")
case "下载地址一"
Z=rsdown("DownloadUrl")
case "下载地址二"
Z=rsdown("DownloadUrl1")
case "下载地址三"
Z=rsdown("DownloadUrl2")
case "软件大小"
Z=rsdown("zdy1")


case "下载内容"
Z=ManualPagination(id,content,6)
case "wap下载内容"
Z=mobileManualPagination(id,content)

case "所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from xzclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
if yy=1 then
Z="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
Z="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from xzclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
Z="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
Z="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc1.close
set rsc1=nothing
end if
if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from xzclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from xzclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
if yy=1 then
Z="<a href="&sitepath&""&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
Z="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
case "WAP所属分类"
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from xzclass where Lang='"&Language&"'",conn,1,1
while not rsc.eof
if zmtwo=0  and rsc("id") =zmone and zmthree=0 then
mqh=rsc("m")
Z="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
end if

if zmone<>0  and  zmtwo= rsc("id") and zmthree=0 then
mqh=rsc("m")
set rsc1=server.CreateObject("adodb.recordset")
rsc1.open "select * from xzclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
Z="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc1.close
set rsc1=nothing
end if
if zmone<>0  and  zmtwo<>0  and zmthree=rsc("id") then
mqh=rsc("m")
set rsc3=server.CreateObject("adodb.recordset")
rsc3.open "select * from xzclass where id="&zmone&" and Lang='"&Language&"'",conn,1,1
set rsc4=server.CreateObject("adodb.recordset")
rsc4.open "select * from xzclass where id="&zmtwo&" and Lang='"&Language&"'",conn,1,1
Z="<a href="&sitepath&"wap/"&iinfo&"/?pone="&zmone&"&ptwo="&zmtwo&"&pthree="&zmthree&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
rsc4.close
set rsc4=nothing
rsc3.close
set rsc3=nothing
end if
rsc.movenext
wend
rsc.close
set rsc=nothing

case "WAP下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id>" & request.QueryString("downid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id>" & request.QueryString("downid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then


if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if














if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle = ""&word(42)&""
end if
Z=nexttitle
rstmp.close
set rstmp=nothing
case "WAP上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id<" & request.QueryString("downid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id<" & request.QueryString("downid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if

if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
Z=prevtitle
rstmp.close
set rstmp=nothing
case "下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id>" & request.QueryString("downid") &" and adddate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id>" & request.QueryString("downid") &" and adddate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if

if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
nexttitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
nexttitle="<a href="&tvName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>"&rstmp("title")&"</a>"
end if
else
nexttitle = ""&word(42)&""
end if
Z=nexttitle
rstmp.close
set rstmp=nothing
case "上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id<" & request.QueryString("downid") & " and adddate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, title,zmone,zmtwo,zmthree from download where id<" & request.QueryString("downid") & " and adddate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp(2)&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")=0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmtwo")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if
if rstmp("zmone")<>0 and rstmp("zmtwo")<>0 and rstmp("zmthree")<>0 then

set rsc2=server.CreateObject("adodb.recordset")
rsc2.open "select * from xzclass where id="&rstmp("zmthree")&" and Lang='"&Language&"'",conn,1,1
guojuan=rsc2("m")
rsc2.close
set rec2=nothing
end if

if yy=1 then
if rstmp(2)<>0 and rstmp(3)=0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)=0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
if rstmp(2)<>0 and rstmp(3)<>0 and rstmp(4)<>0 then
prevtitle="<a href=""?pone="&rstmp("zmone")&"&ptwo="&rstmp("zmtwo")&"&pthree="&rstmp("zmthree")&"&downid=" & rstmp("id") & "&m="&guojuan&"&Language="&Language&""">" & rstmp("title") & "</a>"
end if
else
prevtitle="<a href="&tvName&""&Separated&"" & rstmp("zmone") & ""&Separated&"" & rstmp("id") & ""&Separated&"1.html>" & rstmp("title") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
Z=prevtitle
rstmp.close
set rstmp=nothing
case "点击次数"
Z="<script language=""javascript"" src="""&SitePath&"include/HitCount.asp?id="&downid&"&LX=download""></script><script language=""javascript"" src="""&SitePath&"include/HitCount.asp?tj=yes&LX=download&id="&downid&"""></script>"
end  Select
end  function
function Ztags(TQZ,lx,l)
If rsdown("KeyWord")<>"" then

Ztags=Ztags&"<div class=""tags"">"
Ztags=Ztags&"<ul>"
Ztags=Ztags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsdown("KeyWord")), "|")
For i=0 to Ubound(aa)
Ztags=Ztags&"<li>"
Ztags=Ztags&"<a href="""&SitePath&""&ssearch&"/?Key="&Server.UrlEncode(aa(i))&"&tt=5&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a> "
if lx=1 then
if i<> Ubound(aa) then 
Ztags=Ztags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ztags=Ztags&"</li>"
Next
Ztags=Ztags&"</ul></div>"
End if

end function

function wapZtags(TQZ,lx,l)
If rsdown("KeyWord")<>"" then
Ztags=Ztags&"<div class=""waptags"">"
Ztags=Ztags&"<ul>"
Ztags=Ztags&"<li><span>"&TQZ&"</span></li>"
aa = Split(ucase(rsdown("KeyWord")), "|")
For i=0 to Ubound(aa)
Ztags=Ztags&"<li>"
Ztags=Ztags&"<a href="""&SitePath&"wap/search/?Key="&Server.UrlEncode(aa(i))&"&tt=5&m="&request.QueryString("m")&"&Language="&Language&""">"&aa(i)&"</a> "
if lx=1 then
if i<> Ubound(aa) then 
Ztags=Ztags&"<span class=""fgx"">"&l&"</span>"
end if
end if
Ztags=Ztags&"</li>"
Next
Ztags=Ztags&"</ul></div>"
End if
end function




'招聘内容调用开始
jobid=request.QueryString("jobid")
Set rsjob=Server.CreateObject("ADODB.RecordSet") 
sql="select * from job where id="&jobid
rsjob.Open sql,conn,1,1
content=UBBCode(FormatImg(rsjob("HrDetail")))
content1=UBBCode(rsjob("HrDetail"))
zdfy=rsjob("zdfy")
id=rsjob("ID")
'自定义字段开始
function Z_J(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=7 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=7 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rsjob(""&zdyzd1&"")
Z_J=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束
function J(D)
Select Case D
case "职位名称"
J=rsjob("HrName")
case "标题颜色"
J=rsjob("btcolor")
case "需求人数"
J=rsjob("HrRequireNum")
case "工作地点"
J=rsjob("HrAddress")
case "工资"
J=rsjob("HrSalary")
case "开始时间"
J=rsjob("ksDate")
case "结束时间"
J=rsjob("jsdate")
case "用人单位"
J=rsjob("zpdw")
case "联系电话"
J=rsjob("dh")
case "详细说明"
J=ManualPagination(id,content,7)

case "wap详细说明"

J=mobileManualPagination(id,content)

case "单位地址"
J=rsjob("dz")
case "投递邮箱"
J=rsjob("emaill")
case "招聘状态"
if rsjob("jsDate")>now() then
J=""&word(1)&""
else
J=""&word(2)&""
end if
case "应聘地址"
J=""&sitepath&""&showjob&"/"&ttjjl&"/?Quarters=" & server.urlencode(rsjob("HrName"))&"&m="&request.QueryString("m")&"&Language="&Language&""
case "wap应聘地址"
J=""&sitepath&"wap/tjjl/?Quarters=" & server.urlencode(rsjob("HrName"))&"&m="&request.QueryString("m")&"&Language="&Language&""
case "WAP下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, HrName from job where id>" & request.QueryString("jobid") &" and ksdate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, HrName from job where id>" & request.QueryString("jobid") &" and ksdate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then
nexttitle="<a href="" ?jobid="&rstmp("id")&"&m="&request.QueryString("m")&"&Language="&Language&" "">"&rstmp("HrName")&"</a>"
else
nexttitle = ""&word(42)&""
end if
J=nexttitle
rstmp.close
set rstmp=nothing

case "WAP上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, HrName from job where id<" & request.QueryString("jobid") & " and ksdate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, HrName from job where id<" & request.QueryString("jobid") & " and ksdate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then
prevtitle="<a href=""?jobid=" & rstmp("id") & "&m="&request.QueryString("m")&"&Language="&Language&""">" & rstmp("HrName") & "</a>"
else
prevtitle = ""&word(42)&""
end if
J=prevtitle
rstmp.close
set rstmp=nothing

case "下条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, HrName from job where id>" & request.QueryString("jobid") &" and ksdate<=now() and Lang='"&Language&"' order by id asc ",conn,1,1
else
rstmp.open "select top 1 id, HrName from job where id>" & request.QueryString("jobid") &" and ksdate<=getdate() and Lang='"&Language&"' order by id asc ",conn,1,1
end if
if not rstmp.eof then
if yy=1 then
nexttitle="<a href="" ?jobid="&rstmp("id")&"&m="&request.QueryString("m")&"&Language="&Language&" "">"&rstmp("HrName")&"</a>"
else
nexttitle="<a href="&zpName&""&Separated&"" & rstmp("id") & ".html>"&rstmp("HrName")&"</a>"
end if
else
nexttitle = ""&word(42)&""
end if
J=nexttitle
rstmp.close
set rstmp=nothing
case "上条信息"
set rstmp=server.CreateObject("adodb.recordset")
if sqlno=1 then
rstmp.open "select top 1 id, HrName from job where id<" & request.QueryString("jobid") & " and ksdate<=now() and Lang='"&Language&"'  order by id desc ",conn,1,1
else
rstmp.open "select top 1 id, HrName from job where id<" & request.QueryString("jobid") & " and ksdate<=getdate() and Lang='"&Language&"'  order by id desc ",conn,1,1
end if
if not rstmp.eof then
if yy=1 then
prevtitle="<a href=""?jobid=" & rstmp("id") & "&m="&request.QueryString("m")&"&Language="&Language&""">" & rstmp("HrName") & "</a>"
else
prevtitle="<a href="&zpName&""&Separated&"" & rstmp("id") & ".html>" & rstmp("HrName") & "</a>"
end if
else
prevtitle = ""&word(42)&""
end if
J=prevtitle
rstmp.close
set rstmp=nothing
case "点击次数"
J="<script language=""javascript"" src="""&SitePath&"include/HitCount.asp?id="&jobid&"&LX=job""></script><script language=""javascript"" src="""&SitePath&"include/HitCount.asp?tj=yes&LX=job&id="&jobid&"""></script>"
end  Select
end  function
'营销网点内容调用开始
yxwlid=request.QueryString("yxwlid")
Set rsyxwl=Server.CreateObject("ADODB.RecordSet") 
sql="select * from network where id="&yxwlid
rsyxwl.Open sql,conn,1,1
id=rsyxwl("ID")
content=UBBCode(FormatImg(rsyxwl("content")))
content1=UBBCode(rsyxwl("content"))
zdfy=rsyxwl("zdfy")
'自定义字段开始
function Z_Y(title)
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=8 and ColumnInfo='"&title&"' and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=8 and ColumnInfo='"&title&"'  order by Sequence"
rsObj.open sqlp1,conn,1,1
zdyzd1=rsObj("ColumnName")
zdyzd2= rsjob(""&zdyzd1&"")
Z_Y=zdyzd2
rsObj.close
set rsObj=nothing
END FUNCTION
'自定义字段结束

function Y(D)
Select Case D
case "所属省份"
Y=rsyxwl("Province")
case "公司名称"
Y=rsyxwl("Company")
case "联系地址"
Y=rsyxwl("Address")
case "联系人"
Y=rsyxwl("User")
case "联系电话"
Y=rsyxwl("Tel")
case "传真号码"
Y=rsyxwl("Fax")
case "邮政编码"
Y=rsyxwl("Zip")
case "发布时间"
Y=rsyxwl("Adddate")
case "公司图片组"
pic=rsyxwl("ImagePath")
pic1=Split(pic,"|")
Y=pic1
case "详细说明"
Y=ManualPagination(id,content,8)
case "wap详细说明"
Y=mobileManualPagination(id,content)
case "联系电话"
Y=rsyxwl("dh")
case "公司网址"
Y=rsyxwl("gsurl")
case "下条信息"
set rstmp=server.CreateObject("adodb.recordset")
rstmp.open "select top 1 id, Company from NetWork where id>" & request.QueryString("yxwlid") &" and Lang='"&Language&"'  order by id asc ",conn,1,1

if not rstmp.eof then
nexttitle="<a href="" ?yxwlid="&rstmp(0)&"&m="&request.QueryString("m")&"&Language="&Language&" "">"&rstmp(1)&"</a>"
else
nexttitle = ""&word(42)&""
end if
Y=nexttitle
rstmp.close
set rstmp=nothing
case "上条信息"
set rstmp=server.CreateObject("adodb.recordset")

rstmp.open "select top 1 id, Company from NetWork where id<" & request.QueryString("yxwlid") & " and Lang='"&Language&"'  order by id desc ",conn,1,1

if not rstmp.eof then
prevtitle="<a href=""?yxwlid=" & rstmp(0) & "&m="&request.QueryString("m")&"&Language="&Language&""">" & rstmp(1) & "</a>"

else
prevtitle = ""&word(42)&""
end if
Y=prevtitle
rstmp.close
set rstmp=nothing
end  Select
end  function
function aboutinfo(LX,w,style,msql,gl,N,TZ,D,D1,CZ,P)
set rszm=server.createobject("adodb.recordset")
if lx=1  then
zm = "select top "&n&" * from news where id <> "&rsnews("id")&"  and   ("+sqlKeyWord(rsnews("keyword"),"keyword")+") and Lang='"&Language&"' "

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""

end if
if lx=2  then
zm = "select top "&n&" * from product where id <> "&rspro("id")&"  and   ("+sqlKeyWord(rspro("keyword"),"keyword")+") and Lang='"&Language&"' "

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=3  then
zm = "select top "&n&" * from al where id <> "&rsph("id")&"  and   ("+sqlKeyWord(rsph("keyword"),"keyword")+") and Lang='"&Language&"' "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=4  then
zm = "select top "&n&" * from zxsp where id <> "&rsvideo("id")&"  and   ("+sqlKeyWord(rsvideo("keyword"),"keyword")+") and Lang='"&Language&"' "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
if lx=5  then
zm = "select top "&n&" * from Download where id <> "&rsdown("id")&"  and   ("+sqlKeyWord(rsdown("keyword"),"keyword")+") and Lang='"&Language&"' "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""
end if
'----------------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'----------------------调用样式结束
rszm.open zm,conn,1,1
If not( rszm.bof and rszm.EOF) Then

For i = 1 To rszm.RecordCount
'LX 模块 所属分类带链接结束
dim rsc
set rsc=server.CreateObject("adodb.recordset")
if lx=1 then
rsc.open "select * from newsclass where Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc.open "select * from cpclass where Lang='"&Language&"' ",conn,1,1
end if
if lx=3 then
rsc.open "select * from alclass where Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc.open "select * from spclass where Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc.open "select * from xzclass where Lang='"&Language&"'",conn,1,1
end if
while not rsc.eof
if rszm("zmtwo")=0  and rsc("id") =rszm("zmone") and rszm("zmthree")=0 then
if lx=1 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/info/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if lx=2 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/product/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if lx=3 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/photo/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if


if lx=4 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/video/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if lx=5 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/down/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

end if


if rszm("zmone")<>0  and  rszm("zmtwo")= rsc("id") and rszm("zmthree")=0 then
set rsc1=server.CreateObject("adodb.recordset")
if lx=1 then
rsc1.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

if lx=2 then
rsc1.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc1.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc1.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc1.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if









if lx=1 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if lx=2 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if lx=3 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if lx=4 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if lx=5 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc1("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

end if


if rszm("zmone")<>0  and  rszm("zmtwo")<>0  and rszm("zmthree")=rsc("id") then
set rsc3=server.CreateObject("adodb.recordset")
if lx=1 then
rsc3.open "select * from newsclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc3.open "select * from cpclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc3.open "select * from alclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc3.open "select * from spclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc3.open "select * from xzclass where id="&rszm("zmone")&" and Lang='"&Language&"'",conn,1,1
end if

set rsc4=server.CreateObject("adodb.recordset")
if lx=1 then
rsc4.open "select * from newsclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=2 then
rsc4.open "select * from cpclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=3 then
rsc4.open "select * from alclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=4 then
rsc4.open "select * from spclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if
if lx=5 then
rsc4.open "select * from xzclass where id="&rszm("zmtwo")&" and Lang='"&Language&"'",conn,1,1
end if








if lx=1 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/info/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&iinfo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&newClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if
if lx=2 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/product/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&product&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&cpClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if lx=3 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/photo/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&photo&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&alClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if lx=4 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/video/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else

if yy=1 then
ssfl="<a href="&sitepath&""&video&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&tvClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

if lx=5 then
mqh=rsc("m")
if w<>1 then
ssfl="<a href="&sitepath&"wap/down/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
if yy=1 then
ssfl="<a href="&sitepath&""&down&"/?pone="&rszm("zmone")&"&ptwo="&rszm("zmtwo")&"&pthree="&rsc("id")&"&m="&mqh&"&Language="&Language&">" & rsc("classname") & "</a>"
else
ssfl="<a href="&SitePath&""&rsc3("mulu")&"/"&rsc4("mulu")&"/"&rsc("mulu")&"/"&downClass&""&Separated&""&rsc("id")&""&Separated&"1.html>" & rsc("classname") & "</a>"
end if
end if
end if

end if



rsc.movenext
wend
rsc.close
set rsc=nothing

'LX 模块 所属分类带链接结束
'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束

if lx=1 then
if w=1 then
tempPage=""&sitepath&""&showinfo&"/"
else
tempPage=""&sitepath&"wap/showinfo/"
end if
end if
if lx=2 then
if w=1 then
tempPage=""&sitepath&""&showproduct&"/"
else
tempPage=""&sitepath&"wap/showproduct/"
end if
end if
if lx=3 then
if w=1 then
tempPage=""&sitepath&""&showphoto&"/"
else
tempPage=""&sitepath&"wap/showphoto/"
end if
end if
if lx=4 then
if w=1 then
tempPage=""&sitepath&""&showvideo&"/"
else
tempPage=""&sitepath&"wap/showvideo/"
end if
end if
if lx=5 then
if w=1 then
tempPage=""&sitepath&""&showdown&"/"
else
tempPage=""&sitepath&"wap/showdown/"
end if
end if


'----------------------新闻模块替换开始
if lx=1 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{新闻ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{是否头条}",rsZM("gg"))
tempListStr = replace(tempListStr,"{是否幻灯}",rsZM("tj"))
tempListStr = replace(tempListStr,"{是否置顶}",rsZM("tt"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("sytj"))
'tempListStr = replace(tempListStr,"{是否头条}",rsZM("ft"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{新闻缩略图地址}",RSZM("SmallImage")) '仅仅图片地址
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage")) '仅仅幻灯图片
IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{新闻标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{新闻标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=2 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=2  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("btcolor"))
tempListStr = replace(tempListStr,"{信息来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{新闻作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{新闻缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if








if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if


if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")=0 then
if rszm("wlurl")="" then

if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if

if rszm("zmone")<>0  and rszm("zmtwo")<>0 and rszm("zmthree")<>0 then
if rszm("wlurl")="" then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&newname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&newsID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
else
tempListStr = replace(tempListStr,"{链接地址}",RSZM("wlurl"))
end if
end if




if RSZM("BFNR")="" THEN
tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
ELSE
tempListStr = replace(tempListStr,"{内容简介}",rszm("bfnr"))
END IF
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
END IF

'----------------------新闻模块替换结束

'--------产品模块替换开始
if lx=2 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{产品ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{产品编号}",RSZM("Product_Id"))
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))

IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{产品标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{产品标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
tempListStr = replace(tempListStr,"{是否新品}",RSZM("Newproduct"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage")) '仅仅地址
tempListStr = replace(tempListStr,"{幻灯图地址}",RSZM("IcoImage")) '仅仅地址
tempListStr = replace(tempListStr,"{产品来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{产品作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{标准价格}",formatNumber2(RSZM("scprice")))
tempListStr = replace(tempListStr,"{优惠价格}",formatNumber2(RSZM("hyprice")))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if


if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if


if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )

else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&productname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&proID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=3 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=3  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{订购地址}",""&sitepath&"member/ProductBuy/?step=1&id="&RSZM("id")&"&m="&gl&"")
end if

'--------产品模块替换结束

'--------图片模块替换开始
if lx=3 then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{图片ID}",rsZM("ID"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{标题颜色}",rsZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{缩略图地址}",rsZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",rsZM("IcoImage"))

IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{图片标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{图片标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{图片来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{图片作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{图片缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if



if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&alname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&phid=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=4 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=4  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
end if
'--------图片模块替换结束
'--------视频模块替换开始
if lx=4  then
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{单视频地址}",RSZM("spurl"))
tempListStr = replace(tempListStr,"{多视频地址}",RSZM("arrspurl"))
tempListStr = replace(tempListStr,"{wap视频地址}",RSZM("wapspurl"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{视频ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",RSZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",RSZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",RSZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",RSZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))

IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{视频标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{视频标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
tempListStr = replace(tempListStr,"{视频来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{视频作者}",RSZM("user"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=5 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=5  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if

if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if

if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{视频缩略图}",""&sitepath&"images/demo.jpg")
end if
if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if



if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id") &"&m="&mqh&"&Language="&Language&"")
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if



if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if

if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id") &"&m="&mqh&"&Language="&Language&"")
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&tvname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&videoID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if







tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

end if

'-----------------视频模块替换结束
'-----------------下载模块替换开始
if lx=5 then

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{多图地址}",rszm("ImagePath"))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{标题颜色}",RSZM("BTCOLOR"))
tempListStr = replace(tempListStr,"{发布时间}",rsZM("adddate"))
tempListStr = replace(tempListStr,"{下载ID}",rsZM("id"))
tempListStr = replace(tempListStr,"{一级栏目ID}",rsZM("zmone"))
tempListStr = replace(tempListStr,"{二级栏目ID}",rsZM("zmtwo"))
tempListStr = replace(tempListStr,"{三级栏目ID}",rsZM("zmthree"))
tempListStr = replace(tempListStr,"{是否推荐}",rsZM("tj"))
tempListStr = replace(tempListStr,"{缩略图地址}",RSZM("SmallImage"))
tempListStr = replace(tempListStr,"{幻灯图片地址}",RSZM("IcoImage"))


IF RSZM("BTCOLOR")<>"" THEN
tempListStr = replace(tempListStr,"{下载标题}","<font color="&RSZM("btcolor")&">"&cutstrx(rsZM("title"),TZ)&"</font>")
ELSE
tempListStr = replace(tempListStr,"{下载标题}",cutstrx(rsZM("title"),TZ))
END IF
tempListStr = replace(tempListStr,"{提示标题}",rsZM("title"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=6 and Lang='"&Language&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=6  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd3=rsObj("ColumnInfo")
zdyzd2=RSZM(""&rsObj("ColumnName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
tempListStr = replace(tempListStr,"{下载来源}",RSZM("ly"))
tempListStr = replace(tempListStr,"{下载作者}",RSZM("user"))
tempListStr = replace(tempListStr,"{点击次数}",RSZM("hits"))
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if


if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
tempListStr = replace(tempListStr,"{适用系统}",RSZM("System"))
tempListStr = replace(tempListStr,"{软件语言}",RSZM("Language"))
tempListStr = replace(tempListStr,"{软件类型}",RSZM("Softclass"))
tempListStr = replace(tempListStr,"{下载地址一}",RSZM("DownloadUrl"))
tempListStr = replace(tempListStr,"{下载地址二}",RSZM("DownloadUrl1"))
tempListStr = replace(tempListStr,"{下载地址三}",RSZM("DownloadUrl2"))
tempListStr = replace(tempListStr,"{文件大小}",RSZM("zdy1"))
if rszm("SmallImage")<>"" then
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&""&RSZM("SmallImage")&"")
else
tempListStr = replace(tempListStr,"{下载缩略图}",""&sitepath&"images/demo.jpg")
end if

if rszm("IcoImage")<>"" then
sltdz=right(""&rszm("IcoImage")&"",(len(""&rszm("IcoImage")&"")-1))	
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&""&sltdz&"")
else
tempListStr = replace(tempListStr,"{幻灯图片}",""&sitepath&"images/demo.jpg")
end if

if rszm("zmtwo")=0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")=0 then
if w=1 then

if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


if rszm("zmtwo")<>0  and rszm("zmone")<>0 and rszm("zmthree")<>0 then
if w=1 then
if yy=1 then
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"html/"&downname&""&Separated&""&rszm("zmone")&""&Separated&""&rszm("id")&""& Separated&"1.html")
end if
else
tempListStr = replace(tempListStr,"{链接地址}",tempPage & "?pone=" & rszm("zmone") & "&ptwo="&rszm("zmtwo")&"&pthree="&rszm("zmthree")&"&downID=" & rszm("id")&"&m="&mqh&"&Language="&Language&"" )
end if
end if


tempListStr = replace(tempListStr,"{内容简介}",cutstrx(rsZM("CONTENT"),CZ))
tempListStr = replace(tempListStr,"{所属分类}",ssfl)

end if
'-----------------下载模块替换结束


returnStr = returnStr & tempListStr
RSZM.movenext
next
RSZM.close 
set RSZM=nothing
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)
aboutINFO = returnStr
else
aboutINFO = word(0)
end if
end function

'----------------在线客服
function SHOWZXKF
if  kf=1 then
Set rs1 = server.CreateObject("adodb.recordset")
sql1 = "select * from kfcs"
rs1.Open sql1, conn, 1, 1
if rs1("kfstyle")="gray"  or rs1("kfstyle")="green" or rs1("kfstyle")="blue" or rs1("kfstyle")="yellow" or rs1("kfstyle")="white" or rs1("kfstyle")="orange"   then 
SHOWZXKF=SHOWZXKF&"<script type=""text/javascript"">" & vbCrLf 
SHOWZXKF=SHOWZXKF&"$(document).ready(function() {" & vbCrLf 
SHOWZXKF=SHOWZXKF&"$(""#scrollsidebar"").fix({" & vbCrLf 
SHOWZXKF=SHOWZXKF&"float : '"&rs1("wz")&"'," & vbCrLf 
SHOWZXKF=SHOWZXKF&"minStatue : "&rs1("zk")&"," & vbCrLf 
SHOWZXKF=SHOWZXKF&"skin : '"&rs1("kfstyle")&"'," & vbCrLf 	
SHOWZXKF=SHOWZXKF&"durationTime : 600" & vbCrLf 
SHOWZXKF=SHOWZXKF&"});" & vbCrLf 
SHOWZXKF=SHOWZXKF&"});" & vbCrLf 
SHOWZXKF=SHOWZXKF&"</script>" & vbCrLf 
SHOWZXKF=SHOWZXKF&"<div class=""scrollsidebar"" id=""scrollsidebar"" style=""top:"&rs1("top")&"px"">" & vbCrLf
SHOWZXKF=SHOWZXKF&"<div class=""side_content"" style="" float:"&rs1("wz")&""">" & vbCrLf
SHOWZXKF=SHOWZXKF&"<div class=""side_list"">" & vbCrLf
SHOWZXKF=SHOWZXKF&"<div class=""side_title""><a title=""隐藏"" class=""close_btn""><span>"&word(55)&"</span></a></div>" & vbCrLf
SHOWZXKF=SHOWZXKF&"<div class=""side_center""> " & vbCrLf           	
SHOWZXKF=SHOWZXKF&"<div class=""qqserver"">" & vbCrLf
set rs_qq=server.createobject("adodb.recordset")
qq1="select * from kf where Lang='"&Language&"' order by px  asc"
rs_qq.open qq1,conn,1,1	
if rs_qq.recordcount<>0 then			 	

do while not rs_qq.eof

SHOWZXKF=SHOWZXKF&"<p>" & vbCrLf
SHOWZXKF=SHOWZXKF&""&rs_qq("qqsm")&"：<a title=""点击这里给我发消息"" href=""http://wpa.qq.com/msgrd?v=3&amp;uin="&rs_qq("qq")&"&amp;site=qq&amp;menu=yes"" target=""_blank"">" & vbCrLf
SHOWZXKF=SHOWZXKF&"<img src=""http://wpa.qq.com/pa?p=2:"&rs_qq("qq")&":"&rs1("qqstyle")&"""> " & vbCrLf
SHOWZXKF=SHOWZXKF&"</a>" & vbCrLf
SHOWZXKF=SHOWZXKF&"</p>" & vbCrLf
rs_qq.movenext
loop
if rs1("fgx")="true" then
SHOWZXKF=SHOWZXKF&"<div class=""x""></div>" & vbCrLf
end if
end if
rs_qq.close
set rs_qq=nothing 

set rs_ww=server.createobject("adodb.recordset")
ww1="select * from ww where Lang='ch' order by px asc"
rs_ww.open ww1,conn,1,1	
if rs_ww.recordcount<>0 then		 	
do while not rs_ww.eof
SHOWZXKF=SHOWZXKF&"<p>" & vbCrLf
SHOWZXKF=SHOWZXKF&""&rs_ww("wwsm")&"： <a target=""_blank"" href=""http://www.taobao.com/webww/ww.php?ver=3&touid="&UrlEncode_GBToUtf8(rs_ww("ww"))&"&siteid=cntaobao&status="&rs1("wwys")&"&charset=utf-8""><img border=""0"" src=""http://amos.alicdn.com/realonline.aw?v=2&uid="&UrlEncode_GBToUtf8(rs_ww("ww"))&"&site=cntaobao&s="&rs1("wwys")&"&charset=utf-8"" alt=""点击这里给我发消息"" /></a>" & vbCrLf
SHOWZXKF=SHOWZXKF&"</p>" & vbCrLf
rs_ww.movenext
loop
if rs1("fgx")="true" then
SHOWZXKF=SHOWZXKF&"<div class=""x""></div>" & vbCrLf
end if
end if
rs_ww.close
set rs_ww=nothing 
SHOWZXKF=SHOWZXKF&"</div>"& vbCrLf
if rs1("ifewm")="true" then
SHOWZXKF=SHOWZXKF&"<div class=""phoneserver"">" & vbCrLf
SHOWZXKF=SHOWZXKF&"<p>" & vbCrLf
if rs1("xzewm")="true" then
SHOWZXKF=SHOWZXKF&"<img src="""&sitepath&""&ewm&""" width="&ewmwidth&"px height="&ewmheight&"px>" & vbCrLf
else
SHOWZXKF=SHOWZXKF&"<img src="""&sitepath&""&wxewm&"""  width="&wxewmwidth&"px height="&wxewmheight&"px>" & vbCrLf
end if
SHOWZXKF=SHOWZXKF&"</p>"& vbCrLf
SHOWZXKF=SHOWZXKF&"<p>"& vbCrLf
SHOWZXKF=SHOWZXKF&""&rs1("ewmsm")&""& vbCrLf
SHOWZXKF=SHOWZXKF&"</p>	"& vbCrLf
SHOWZXKF=SHOWZXKF&"</div> "& vbCrLf
end if             
SHOWZXKF=SHOWZXKF&"</div>" & vbCrLf
SHOWZXKF=SHOWZXKF&"<div class=""side_bottom""></div>" & vbCrLf
SHOWZXKF=SHOWZXKF&"</div>" & vbCrLf
SHOWZXKF=SHOWZXKF&"</div>" & vbCrLf
SHOWZXKF=SHOWZXKF&"<div class=""show_btn"" style="" float:"&rs1("wz")&"""><span>在线客服</span></div>" & vbCrLf
SHOWZXKF=SHOWZXKF&"</div>" & vbCrLf
end if
if rs1("kfstyle")="syellow"  or rs1("kfstyle")="dblue" or rs1("kfstyle")="sblue" or rs1("kfstyle")="black" or rs1("kfstyle")="red" or rs1("kfstyle")="tblue"   then
set rs_ww=server.createobject("adodb.recordset"):ww1="select * from ww where Lang='ch' order by px asc":rs_ww.open ww1,conn,1,1
set rs_qq=server.createobject("adodb.recordset"):qq1="select * from kf where Lang='ch' order by px  asc":rs_qq.open qq1,conn,1,1

SHOWZXKF=SHOWZXKF&"<script type=""text/javascript"">"
SHOWZXKF=SHOWZXKF&"$(document).ready(function () {onlineService.init({style: '"&rs1("kfstyle")&"',onlineside: '"&rs1("wz")&"',onlineside_width: "&rs1("jl")&",top: """&rs1("top")&""",title: '在线客服',"
SHOWZXKF=SHOWZXKF&"content: ""<div class='chat'><div class='chat_t'>"&word(466)&"</div><ul>"
do while not rs_qq.eof
SHOWZXKF=SHOWZXKF&"<li>"&rs_qq("qqsm")&"：<a target='_self'  href='tencent://message/?uin="&rs_qq("qq")&"&Site=追梦工作室&Menu=yes'><img border=0 SRC='http://wpa.qq.com/pa?p=2:"&rs_qq("qq")&":"&rs1("qqstyle")&"' alt=Q我 align='absmiddle'></a></li>"
rs_qq.movenext:loop:rs_qq.close:set rs_qq=nothing 
SHOWZXKF=SHOWZXKF&"</ul></div>"
if rs_ww.recordcount<>0 then
SHOWZXKF=SHOWZXKF&"<div class='chat'><div class='chat_t'>"&word(467)&"</div><ul>"
do while not rs_ww.eof
SHOWZXKF=SHOWZXKF&"<li>"&rs_ww("wwsm")&"：<a target='_blank' href='http://www.taobao.com/webww/ww.php?ver=3&touid="&UrlEncode_GBToUtf8(rs_ww("ww"))&"&siteid=cntaobao&status="&rs1("wwys")&"&charset=utf-8'><img border=0 src='http://amos.alicdn.com/realonline.aw?v=2&uid="&UrlEncode_GBToUtf8(rs_ww("ww"))&"&site=cntaobao&s="&rs1("wwys")&"&charset=utf-8' alt='点击这里给我发消息' /></a></li>"
rs_ww.movenext:loop:rs_ww.close:set rs_ww=nothing
SHOWZXKF=SHOWZXKF&"</ul></div>"
end if 
if rs1("ifewm")="true" then
SHOWZXKF=SHOWZXKF&"<div class='ewm'><ul><li>"
if rs1("xzewm")="true" then
SHOWZXKF=SHOWZXKF&"<img src='"&sitepath&""&ewm&"' width="&ewmwidth&"px height="&ewmheight&"px>"
else
SHOWZXKF=SHOWZXKF&"<img src='"&sitepath&""&wxewm&"'  width="&wxewmwidth&"px height="&wxewmheight&"px>"
end if
SHOWZXKF=SHOWZXKF&"</li></ul><div class='ewmsm'>"&rs1("ewmsm")&"</div></div>"
END IF
SHOWZXKF=SHOWZXKF&""",oday_from: """&rs1("gzks")&""",oday_to: """&rs1("gzjs")&""",otime_from: """&rs1("zxsjks")&""",otime_to: """&rs1("zxsjjs")&""",hotline: """&rs1("fwdh")&""",isexpand: """&rs1("sfzk")&"""});setTimeout(function () {$("".smc_loading li,.brands li"").css(""background"", ""none"");}, 5000);});"

SHOWZXKF=SHOWZXKF&"</script>"
end if
end if
end function

'在线调查
function ShowVote(nlt,w,style,u,m,D,D1)
'打开数据库
if nlt=0 then
nlt= dcid
else
end if

set rsVote=server.CreateObject("adodb.recordset")
sqlVote="select top 1 id,title,vote,result,stype,StartTime,EndTime,lang from vote Where yn = 1 and Lang='"&Language&"'"
if clng(nlt)<>0 then sqlVote=sqlVote&" And id="&clng(nlt)&""
sqlVote=sqlVote&" order by id desc"

'循环标签开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr

'循环标签结束
rsVote.open sqlVote,conn,1,1	

if rsVote.eof then
else



if rsVote(4)=1 then
v_type="radio"
else
v_type="checkbox"
end if
result=split(rsVote(3),"|")
for i=0 to ubound(result)
next
vote=split(rsVote(2),"|")




for i=0 to ubound(vote)-1


linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rsVote(7))
tempListStr = replace(tempListStr,"{地址语言}",language)


tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{type类型}",""&v_type&"")
tempListStr = replace(tempListStr,"{type名称}","vote")
tempListStr = replace(tempListStr,"{调查选性名称}",""&vote(i)&"")
returnStr = returnStr & tempListStr
next

returnStr = returnStr & endStr
returnStr = replace(returnStr,"{地址栏语言}",language)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)
returnStr = replace(returnStr,"{开始时间}",FormatTime(rsVote(5),D))
returnStr = replace(returnStr,"{结束时间}",FormatTime(rsVote(6),D))
returnStr = replace(returnStr,"{开始时间1}",FormatTime(rsVote(5),D1))
returnStr = replace(returnStr,"{结束时间1}",FormatTime(rsVote(6),D1))
returnStr = replace(returnStr,"{调查项目名称}",rsVote(1))
if w=1 then
returnStr = replace(returnStr,"{表单提交地址}",""&SitePath&""&Vote1&"/?action=add&id="&rsVote(0)&"&u="&u&"&Language="&Language&"")
returnStr = replace(returnStr,"{查看地址}",""&SitePath&""&Vote1&"/?action=see&id="&rsVote(0)&"&m="&m&"&Language="&Language&"")
else
returnStr = replace(returnStr,"{表单提交地址}",""&SitePath&"wap/vote/?action=add&id="&rsVote(0)&"&u="&u&"&Language="&Language&"")
returnStr = replace(returnStr,"{查看地址}",""&SitePath&"wap/vote/?action=see&id="&rsVote(0)&"&m="&m&"&Language="&Language&"")
end if
ShowVote = returnStr


end if
RsVote.Close
Set RsVote = Nothing	
End function
'查看调查结果
function seevote(style,D,D1)
uu=request("u")
ID=request("id")
if request("action") = "add" then 
set rs=server.CreateObject("adodb.recordset")
sql="select result,StartTime,EndTime from Vote where id="&id&""
Rs.Open sql,Conn,1,3 
If  rs.eof then
writemsg1 2,""&word(58)&"",""&uu&"","wrong"
Else
If Request.Form("vote")="" then
writemsg1 2,""&word(59)&"",""&uu&"","wrong"
End if
If Datediff("s",Rs(1),Now())<0 Then writemsg1 2,""&word(60)&"",""&uu&"","wrong"
If Datediff("s",Rs(2),Now())>0 Then writemsg1 2,""&word(61)&"",""&uu&"","wrong"
If Request.Cookies("zmcmsvote_"&id&"")="" then
for each ho in Request.Form("vote")
pollresult=split(Rs(0),"|")
for i = 0 to (ubound(pollresult)-1)
if not pollresult(i)="" then
if int(ho)=i then
pollresult(i)=pollresult(i)+1
end if
allpollresult=""&allpollresult&""&pollresult(i)&"|"
end if
next
Rs(0)=allpollresult
Rs.update
allpollresult=""
next
Response.Cookies("zmcmsvote_"&id&"")=id
Response.Cookies("zmcmsvote_"&id&"").Expires=Date+30
writemsg1 2,""&word(62)&"",""&uu&"","right"
Else
writemsg1 2,""&word(63)&"",""&uu&"","wrong"
end if 
End if
rs.close
set rs=nothing

'投票已结束

'查看投票开始

elseif request("action")="see" then
set rs=server.CreateObject("adodb.recordset")
sql="select title,vote,result,stype,StartTime,EndTime,lang from Vote where id="&ID&""
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
rs.open sql,conn,1,1

if rs.eof then
writemsg1 2,""&word(64)&"",""&uu&"","wrong"
else
vote=rs(1)
result=rs(2)
total_vote=0
vote=split(vote,"|")
result=split(result,"|")
for i=0 to ubound(result)
if not result(i)="" then total_vote=result(i)+total_vote
next
for i=0 to (ubound(vote)-1)
if result(i)=0 then
b2="0%"
b3="0"
b4="0%"
b5="0"
else
b2=Formatpercent(result(i)/total_vote,0)
b3=FormatNumber((result(i)/total_vote)*100,0)
b4=Formatpercent(result(i)/total_vote,-1)
b5=FormatNumber((result(i)/total_vote)*100,2)
end if

linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rs(6))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{调查选性}",""&vote(i)&"")
tempListStr = replace(tempListStr,"{票数}",result(i))
tempListStr = replace(tempListStr,"{百分比1}",b2)
tempListStr = replace(tempListStr,"{百分比2}",b3)
tempListStr = replace(tempListStr,"{百分比3}",b4)
tempListStr = replace(tempListStr,"{百分比4}",b5)

returnStr = returnStr & tempListStr
next
returnStr = returnStr & endStr
tempListStr = replace(tempListStr,"{地址栏语言}",language)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)
returnStr = replace(returnStr,"{开始时间}",FormatTime(rs(4),D))
returnStr = replace(returnStr,"{结束时间}",FormatTime(rs(5),D))
returnStr = replace(returnStr,"{开始时间1}",FormatTime(rs(4),D1))
returnStr = replace(returnStr,"{结束时间1}",FormatTime(rs(5),D1))
returnStr = replace(returnStr,"{调查项目名称}",rs(0))
returnStr = replace(returnStr,"{总票数}",total_vote)


seevote = returnStr
end if
rs.close
set rs=nothing
end if
end function

'热搜关键词
function hotkey(style,w,msql,n,m,p)



set rs=server.CreateObject("adodb.recordset")
zm = "select top "&n&" * from hotkey where Lang='"&Language&"' " 

if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""


tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
RS.open ZM,conn,1,1





do while not rs.eof

linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rs("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{搜索关键词}",""&rs("keywords")&"")
if w=1 then
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&""&ssearch&"/?key="&rs("keywords")&"&tt="&rs("ssfl")&"&m="&m&"&Language="&Language&"")
else
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&"wap/search/?key="&rs("keywords")&"&tt="&rs("ssfl")&"&m="&m&"&Language="&Language&"")
end if
returnStr = returnStr & tempListStr

rs.movenext
loop

returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)
hotkey = returnStr
rs.close
set rs=nothing
end function








function MemGroup(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from hyGroup where GroupID='"&GroupID&"' and Lang='"&Language&"'"
  rs.open sql,conn,1,1
  MemGroup=rs("GroupName")
  rs.close
  set rs=nothing
end function
function GroupName(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from hyGroup where GroupID='"&GroupID&"' and Lang='"&Language&"'"
  rs.open sql,conn,1,1
  GroupName=rs("GroupName")
  rs.close
  set rs=nothing
end function

Function sqlKeyWord(obj,field) 
Dim temp:temp = split(obj,"|") 
For i = 0 To ubound(temp) 
sqlKeyWord = sqlKeyWord & field&" like '%"&temp(i)&"%' or " 
Next 
sqlKeyWord = left(sqlKeyWord,len(sqlKeyWord)-3) 
End Function 

if (session("MemName")<>"" and session("MemLogin")="Succeed") then 
set rsly = server.createobject("adodb.recordset")
sqlly="select * from Members where MemName='"&session("MemName")&"' and Lang='"&Language&"'"
rsly.open sqlly,conn,1,1
mMemID=rsly("ID")
if not rsly("RealName")="" then
mRealName=rsly("RealName")
else 
mRealName=rsly("MemName")
end if
mSex=rsly("Sex")
mCompany=rsly("Company")
mAddress=rsly("Address")
mZipCode=rsly("ZipCode")
mFax=rsly("Fax")
mMobile=rsly("Mobile")
mEmail=rsly("Email")
qqh="<input type=""checkbox"" name=""qqh"" value=""1"" style=""width:10px;height:10px"">"
rsly.close
set rsly=nothing
else
mSex=""&word(69)&""
mMemID=0
  qqh="<font color="""&word(71)&""">"&word(70)&"</font>"
  end if

'默认图片和产品多图展示样式
function SHOWTPnr(lx,ys,slt,slth,w,h)
DIM WH,HW,slt1,slth1
WH=w
hw=h
slt1=slt
slth1=slth
SHOWTPnr=SHOWTPnr&"<link rel=""stylesheet"" type=""text/css"" href="""&sitepath&"images/jquery.lightbox-0.5.css"" media=""screen"" />"&VbCrLf
SHOWTPnr=SHOWTPnr&"<script type=""text/javascript"">"&VbCrLf
SHOWTPnr=SHOWTPnr&"myFocus.set({"&VbCrLf
if lx=1 then
SHOWTPnr=SHOWTPnr&"id:'cpzs',"&VbCrLf
end if
if lx=2 then
SHOWTPnr=SHOWTPnr&"id:'alzs',"&VbCrLf
end if
if ys=1 then
SHOWTPnr=SHOWTPnr&"pattern:'mF_games_tb',"&VbCrLf
end if
if ys=2 then
SHOWTPnr=SHOWTPnr&"pattern:'mF_fscreen_tb',"&VbCrLf
end if
SHOWTPnr=SHOWTPnr&"width:"&wh&","&VbCrLf
SHOWTPnr=SHOWTPnr&"height:"&hw&","&VbCrLf
SHOWTPnr=SHOWTPnr&"txtHeight:0,"&VbCrLf
SHOWTPnr=SHOWTPnr&"thumbShowNum:"&slt1&","&VbCrLf
SHOWTPnr=SHOWTPnr&"thumbHeight:"&slth1&","&VbCrLf
SHOWTPnr=SHOWTPnr&"wrap:false"&VbCrLf
SHOWTPnr=SHOWTPnr&"});"&VbCrLf
SHOWTPnr=SHOWTPnr&"</script>"&VbCrLf
if lx=1 then
SHOWTPnr=SHOWTPnr&"<div id=""cpzs"">"&VbCrLf
end if
if lx=2 then
SHOWTPnr=SHOWTPnr&"<div id=""alzs"">"&VbCrLf
end if
SHOWTPnr=SHOWTPnr&"<div class=""loading""></div>"&VbCrLf
SHOWTPnr=SHOWTPnr&"<div class=""pic"">"&VbCrLf
SHOWTPnr=SHOWTPnr&"<div id=""gallery""  class=""tpjl"">"&VbCrLf
SHOWTPnr=SHOWTPnr&"<ul>"&VbCrLf
if lx=1 then
for i=0 to ubound(P("产品图片组"))
sltdz=right(""&P("产品图片组")(i)&"",(len(""&P("产品图片组")(i)&"")-1))	
sltdz1=""&sitepath&""&sltdz&""
SHOWTPnr=SHOWTPnr& "<li><a href="""&sltdz1&""" title="""&P("产品标题")&"""  target=""_blank""><img src="""&sltdz1&""" width="""&wh&""" height="""&hw&""" alt="""&title&""" ></a></li>"&VbCrLf
next
end if
if lx=2 then
for i=0 to ubound(C("图片组"))
sltdz=right(""&c("图片组")(i)&"",(len(""&c("图片组")(i)&"")-1))	
sltdz1=""&sitepath&""&sltdz&""
SHOWTPnr=SHOWTPnr& "<li><a href="""&sltdz1&""" title="""&c("图片标题")&"""  target=""_blank""><img src="""&sltdz1&""" width="""&wh&""" height="""&hw&""" alt="""&title&""" ></a></li>"&VbCrLf
next
end if
SHOWTPnr=SHOWTPnr&"</ul></div></div></div>"&VbCrLf
end function
'**************************************************
'函 数 名：showtv
'作    用：显示视频
'*************************************************
function SHOWTV()
Dim Exts,isExt,wh,hw,url1

url1=rsvideo("spurl")
If url1 <> "" Then
   isExt = LCase(Mid(url1,InStrRev(url1, ".")+1))
Else
   isExt = ""
End If
Exts = "swf,flv,mp4"

If Instr(Exts,isExt) = 0 Then

SHOWTV=SHOWTV&"<embed src=""http://static.youku.com/v/swf/qplayer.swf?VideoIDS="&rsvideo("spurl")&"&winType=adshow&isAutoPlay=false"" quality=""high"" width="""&rsvideo("spwidth")&""" height="""&rsvideo("spheight")&""" allowfullscreen=""true"" type=""application/x-shockwave-flash""></embed>"& vbCrLf 

Else
 Select Case isExt
  Case "flv","mp4"

SHOWTV=SHOWTV&"<div id=""container"">Loading.....</div>  "& vbCrLf 
SHOWTV=SHOWTV&"<script type=""text/javascript"">  "& vbCrLf 
SHOWTV=SHOWTV&"var thePlayer;"  & vbCrLf 
SHOWTV=SHOWTV&"$(function() {  "& vbCrLf 
SHOWTV=SHOWTV&"thePlayer = jwplayer('container').setup({  "& vbCrLf 
SHOWTV=SHOWTV&"flashplayer: '"&sitepath&"inc/player5.7.swf',  	"		& vbCrLf 
if rsvideo("arrspurl")<>"" then
pic=rsvideo("arrspurl")
pic1=Split(pic,"|")
SHOWTV=SHOWTV&"playlist:["& vbCrLf 
for i=0 to ubound(pic1)
SHOWTV=SHOWTV&"{file: '"&UrlEncode_GBToUtf8(pic1(i))&"', image: '"&sitepath&""&rsvideo("SmallImage")&"'},"& vbCrLf 
next
SHOWTV=SHOWTV&"],  "& vbCrLf 
else
SHOWTV=SHOWTV&"playlist: [{file: '"&UrlEncode_GBToUtf8(rsvideo("spurl"))&"', image: '"&sitepath&""&rsvideo("SmallImage")&"',title: '"&rsvideo("title")&"'}],"& vbCrLf 
end if
       SHOWTV=SHOWTV&"width: "&rsvideo("spwidth")&", "& vbCrLf 
			SHOWTV=SHOWTV&"screencolor:"""&rsvideo("screencolor")&""", "& vbCrLf 
            SHOWTV=SHOWTV&"height: "&rsvideo("spheight")&","& vbCrLf 
			SHOWTV=SHOWTV&"stretching:"""&rsvideo("stretching")&""",  "& vbCrLf 
			SHOWTV=SHOWTV&"controlbar: """&rsvideo("controlbar")&""", "& vbCrLf 
			SHOWTV=SHOWTV&"skin:"""&sitepath&"inc/skin\/"&rsvideo("skin")&".zip"", "& vbCrLf 
			SHOWTV=SHOWTV&"repeat:""list"""& vbCrLf 
SHOWTV=SHOWTV&"}); " & vbCrLf 
SHOWTV=SHOWTV&" });  "& vbCrLf 
SHOWTV=SHOWTV&"</script>"  & vbCrLf 
Case "swf"
SHOWTV=SHOWTV&"<object codeBase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width="""&rsvideo("spwidth")&""" height="""&rsvideo("spheight")&""" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""_cx"" value=""19050"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""_cy"" value=""3863"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""FlashVars"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Movie"" value="""&rsvideo("spurl")&""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Src"" value="""&rsvideo("spurl")&""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Play"" value=""-1"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Loop"" value=""-1"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Quality"" value=""High"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""SAlign"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Menu"" value=""0"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Base"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""AllowScriptAccess"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Scale"" value=""ShowAll"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""DeviceFont"" value=""0"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""EmbedMovie"" value=""0"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""BGColor"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""SWRemote"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""allowFullScreen"" value=""true"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""MovieData"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""SeamlessTabbing"" value=""1"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""Profile"" value=""0"">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""ProfileAddress"" value="""">"& vbCrLf 
SHOWTV=SHOWTV&"<param name=""ProfilePort"" value=""0"">"& vbCrLf 
SHOWTV=SHOWTV&"<embed src="""&rsvideo("spurl")&""" width="""&rsvideo("spwidth")&""" height="""&rsvideo("spheight")&""" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" menu=""false"" allowFullScreen=""true"" ></embed>"& vbCrLf 
SHOWTV=SHOWTV&"</object>"& vbCrLf 
End Select
End If
end function
function FormatNum(num,n)
if num<1 then
num="0"&cstr(FormatNumber(num,n))
else
num=cstr(FormatNumber(num,n))
end if
FormatNum=replace(num,",","")
end function

  set rsmemberinfo = server.createobject("adodb.recordset")
  sqlmemberinfo="select * from Members where MemName='"&session("MemName")&"' and Lang='"&Language&"'"
  rsmemberinfo.open sqlmemberinfo,conn,1,1
  if rsmemberinfo.eof then
  
  else
    MemIDmemberinfo=rsmemberinfo("ID")
	mMemNamememberinfo=rsmemberinfo("MemName")
	mGroupIdNamememberinfo=GroupName(rsmemberinfo("GroupID"))
	mRealNamememberinfo=rsmemberinfo("RealName")
	mSexmemberinfo=rsmemberinfo("Sex")
	mQusetionmemberinfo=rsmemberinfo("Question")
	mCompanymemberinfo=rsmemberinfo("Company")
	mAddressmemberinfo=rsmemberinfo("Address")
	mZipCodememberinfo=rsmemberinfo("ZipCode")
	mTelephonememberinfo=rsmemberinfo("Telephone")
	mFaxmemberinfo=rsmemberinfo("Fax")
	mMobilememberinfo=rsmemberinfo("Mobile")
	mEmailmemberinfo=rsmemberinfo("Email")
	mHomePagememberinfo=rsmemberinfo("HomePage")
	mAddTimememberinfo=rsmemberinfo("AddTime")
	mLoginTimesmemberinfo=rsmemberinfo("LoginTimes")
	mLastLoginTimememberinfo=rsmemberinfo("LastLoginTime")
	mLastLoginIPmemberinfo=rsmemberinfo("LastLoginIP")
	qqopenapi=rsmemberinfo("qqopenapi")
  end if
  rsmemberinfo.close
  set rsmemberinfo=nothing


'------当前位置的调用
function ddwz(LX,F)
pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
if pone=""  then
zmcmsTitle=""&word(297)&""
end if
if pone<>"" then
set lm=server.createobject("adodb.recordset") 
if lx=1 then
exec="select * from [newsclass] where id="&pone&" and Lang='"&Language&"'"
end if
if lx=2 then
exec="select * from [cpclass] where id="&pone&" and Lang='"&Language&"'"
end if
if lx=3 then
exec="select * from [alclass] where id="&pone&" and Lang='"&Language&"'"
end if
if lx=4 then
exec="select * from [spclass] where id="&pone&" and Lang='"&Language&"'"
end if
if lx=5 then
exec="select * from [xzclass] where id="&pone&" and Lang='"&Language&"'"
end if
lm.open exec,conn,1,1
if lx=1 then
zmcmsTitle="<a href="""&sitepath&"info/?pone="&pone&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=2 then
zmcmsTitle="<a href="""&sitepath&"product/?pone="&pone&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=3 then
zmcmsTitle="<a href="""&sitepath&"photo/?pone="&pone&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=4 then
zmcmsTitle="<a href="""&sitepath&"video/?pone="&pone&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=5 then
zmcmsTitle="<a href="""&sitepath&"down/?pone="&pone&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
lm.close
set lm=nothing
end if
if ptwo<>"" then
set lm=server.createobject("adodb.recordset") 
if lx=1 then
exec="select * from [newsclass] where id="&ptwo&" and Lang='"&Language&"'"
end if
if lx=2 then
exec="select * from [cpclass] where id="&ptwo&" and Lang='"&Language&"'"
end if
if lx=3 then
exec="select * from [alclass] where id="&ptwo&" and Lang='"&Language&"'"
end if
if lx=4 then
exec="select * from [spclass] where id="&ptwo&" and Lang='"&Language&"'"
end if
if lx=5 then
exec="select * from [xzclass] where id="&ptwo&" and Lang='"&Language&"'"
end if
lm.open exec,conn,1,1

if lx=1 then
zmcmsTitlew=""&F&"<a href="""&sitepath&"info/?pone="&pone&"&ptwo="&ptwo&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=2 then
zmcmsTitlew=""&F&"<a href="""&sitepath&"product/?pone="&pone&"&ptwo="&ptwo&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=3 then
zmcmsTitlew=""&F&"<a href="""&sitepath&"photo/?pone="&pone&"&ptwo="&ptwo&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=4 then
zmcmsTitlew=""&F&"<a href="""&sitepath&"video/?pone="&pone&"&ptwo="&ptwo&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=5 then
zmcmsTitlew=""&F&"<a href="""&sitepath&"down/?pone="&pone&"&ptwo="&ptwo&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
lm.close
set lm=nothing
end if

if pthree<>"" then
set lm=server.createobject("adodb.recordset") 
if lx=1 then
exec="select * from [newsclass] where id="&pthree&" and Lang='"&Language&"'"
end if
if lx=2 then
exec="select * from [cpclass] where id="&pthree&" and Lang='"&Language&"'"
end if
if lx=3 then
exec="select * from [alclass] where id="&pthree&" and Lang='"&Language&"'"
end if
if lx=4 then
exec="select * from [spclass] where id="&pthree&" and Lang='"&Language&"'"
end if
if lx=5 then
exec="select * from [xzclass] where id="&pthree&" and Lang='"&Language&"'"
end if
lm.open exec,conn,1,1

if lx=1 then
zmcmsTitleh=""&F&"<a href="""&sitepath&"info/?pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=2 then
zmcmsTitleh=""&F&"<a href="""&sitepath&"product/?pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=3 then
zmcmsTitleh=""&F&"<a href="""&sitepath&"photo/?pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=4 then
zmcmsTitleh=""&F&"<a href="""&sitepath&"video/?pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
if lx=5 then
zmcmsTitleh=""&F&"<a href="""&sitepath&"down/?pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&m="&lm("m")&"&Language="&Language&""">"&lm("ClassName")&"</a>"
end if
lm.close
set lm=nothing
end if
ddwz=ddwz&"<a href="""&sitepath&""">"&word(300)&"</a>"&F&""&zmcmstitle&""&zmcmstitlew&""&zmcmstitleh&""
end function



'----------------首页调用视频
function SelPlay(id,ys,w,h)
if id=0 then
id= tvid
else
end if
Set rssp=Server.CreateObject("ADODB.RecordSet") 
sql="select * from zxsp where id="&id&" "
rssp.Open sql,conn,1,1
Dim Exts,isExt,wh,hw,url
wh=w 
hw=h
url=rssp("spurl")
If url <> "" Then
   isExt = LCase(Mid(url,InStrRev(url, ".")+1))
Else
   isExt = ""
End If
Exts = "swf,flv,mp4"
If Instr(Exts,isExt) = 0 Then

SelPlay=SelPlay&"<embed src=""http://static.youku.com/v/swf/qplayer.swf?VideoIDS="&rssp("spurl")&"&winType=adshow&isAutoPlay=false"" quality=""high"" width="""&wh&""" height="""&hw&""" allowfullscreen=""true"" type=""application/x-shockwave-flash""></embed>"
else
Select Case isExt
  Case "flv","mp4"
SelPlay=SelPlay&"<div id=""container"">Loading.....</div>  "
SelPlay=SelPlay&"<script type=""text/javascript"">"  
SelPlay=SelPlay&"var thePlayer;"  
SelPlay=SelPlay&"$(function() {"  
SelPlay=SelPlay&"thePlayer = jwplayer('container').setup({  "
SelPlay=SelPlay&"flashplayer: '"&sitepath&"inc/player5.7.swf',  "

if rssp("arrspurl")<>"" then
pic=rssp("arrspurl")
pic1=Split(pic,"|")
SelPlay=SelPlay&"playlist:["
for i=0 to ubound(pic1)
SelPlay=SelPlay&"{file: '"&UrlEncode_GBToUtf8(pic1(i))&"', image: '"&sitepath&""&rssp("SmallImage")&"'},"
next
SelPlay=SelPlay&" ], " 
else
SelPlay=SelPlay&"playlist: [{file: '"&UrlEncode_GBToUtf8(rssp("spurl"))&"', image: '"&sitepath&""&rssp("SmallImage")&"',title: '第一个视频'}],"
end if
            SelPlay=SelPlay&"width: "&wh&", "
			SelPlay=SelPlay&"screencolor:"""&rssp("screencolor")&"""," 
            SelPlay=SelPlay&"height: "&hw&","
			SelPlay=SelPlay&"stretching:"""&rssp("stretching")&""",  "
			SelPlay=SelPlay&"controlbar: """&rssp("controlbar")&""","
			SelPlay=SelPlay&"skin:"""&sitepath&"inc/skin\/"&ys&".zip"", "
			SelPlay=SelPlay&"repeat:""list"""
			  
        SelPlay=SelPlay&"});  "
          

          
  
    SelPlay=SelPlay&"});  "
SelPlay=SelPlay&"</script>  "
Case "swf"
SelPlay=SelPlay&"<object codeBase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width="""&wh&""" height="""&hw&""" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"">"
SelPlay=SelPlay&"<param name=""_cx"" value=""19050"">"
SelPlay=SelPlay&"<param name=""_cy"" value=""3863"">"
SelPlay=SelPlay&"<param name=""FlashVars"" value="""">"
SelPlay=SelPlay&"<param name=""Movie"" value="""&rssp("spurl")&""">"
SelPlay=SelPlay&"<param name=""Src"" value="""&rssp("spurl")&""">"
SelPlay=SelPlay&"<param name=""Play"" value=""-1"">"
SelPlay=SelPlay&"<param name=""Loop"" value=""-1"">"
SelPlay=SelPlay&"<param name=""Quality"" value=""High"">"
SelPlay=SelPlay&"<param name=""SAlign"" value="""">"
SelPlay=SelPlay&"<param name=""Menu"" value=""0"">"
SelPlay=SelPlay&"<param name=""Base"" value="""">"
SelPlay=SelPlay&"<param name=""AllowScriptAccess"" value="""">"
SelPlay=SelPlay&"<param name=""Scale"" value=""ShowAll"">"
SelPlay=SelPlay&"<param name=""DeviceFont"" value=""0"">"
SelPlay=SelPlay&"<param name=""EmbedMovie"" value=""0"">"
SelPlay=SelPlay&"<param name=""BGColor"" value="""">"
SelPlay=SelPlay&"<param name=""SWRemote"" value="""">"
SelPlay=SelPlay&"<param name=""allowFullScreen"" value=""true"">"
SelPlay=SelPlay&"<param name=""MovieData"" value="""">"
SelPlay=SelPlay&"<param name=""SeamlessTabbing"" value=""1"">"
SelPlay=SelPlay&"<param name=""Profile"" value=""0"">"
SelPlay=SelPlay&"<param name=""ProfileAddress"" value="""">"
SelPlay=SelPlay&"<param name=""ProfilePort"" value=""0"">"
SelPlay=SelPlay&"<embed src="""&rssp("spurl")&""" width="""&wh&""" height="""&hw&""" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" menu=""false"" allowFullScreen=""true""></embed>"
SelPlay=SelPlay&"</object>"
End Select
End If
End function

'-----及时咨询
function messageskin
If JMailDisplay="1" Then 
if zxys="1" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=default&SkinB=Black"" charset=""utf-8""></script>"
elseif zxys="2" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=default&SkinB=Blue"" charset=""utf-8""></script>"
elseif zxys="3" then 
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=default&SkinB=Gray"" charset=""utf-8""></script>"
elseif zxys="4" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=default&SkinB=red"" charset=""utf-8""></script>"
elseif zxys="5" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=honey&SkinB=Green"" charset=""utf-8""></script>"
elseif zxys="6" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=friendly&SkinB=Red"" charset=""utf-8""></script>"
elseif zxys="7" then 
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=honey&SkinB=Gray"" charset=""utf-8""></script>"
elseif zxys="8" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=honey&SkinB=Red"" charset=""utf-8""></script>"
elseif zxys="9" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=special&SkinB=QQ"" charset=""utf-8""></script>"
elseif zxys="10" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=cidilife&SkinB=Blue"" charset=""utf-8""></script>"
elseif zxys="11" then 
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=cidilife&SkinB=Gray"" charset=""utf-8""></script>"
elseif zxys="12" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=cidilife&SkinB=Red"" charset=""utf-8""></script>"

elseif zxys="13" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=classic&SkinB=Black"" charset=""utf-8""></script>"

elseif zxys="14" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=classic&SkinB=Blue"" charset=""utf-8""></script>"

elseif zxys="15" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=classic&SkinB=Gray"" charset=""utf-8""></script>"

elseif zxys="16" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=classic&SkinB=Red"" charset=""utf-8""></script>"
elseif zxys="17" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=friendly&SkinB=Black"" charset=""utf-8""></script>"

elseif zxys="18" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=friendly&SkinB=Blue"" charset=""utf-8""></script>"

elseif zxys="19" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=friendly&SkinB=Gray"" charset=""utf-8""></script>"

elseif zxys="20" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=friendly&SkinB=Red"" charset=""utf-8""></script>"

elseif zxys="21" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=luxury&SkinB=Black"" charset=""utf-8""></script>"

elseif zxys="22" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=luxury&SkinB=Blue"" charset=""utf-8""></script>"

elseif zxys="23" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=luxury&SkinB=Gray"" charset=""utf-8""></script>"

elseif zxys="24" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=luxury&SkinB=Red"" charset=""utf-8""></script>"

elseif zxys="25" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=special&SkinB=Black"" charset=""utf-8""></script>"

elseif zxys="26" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=special&SkinB=Blue"" charset=""utf-8""></script>"

elseif zxys="27" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=special&SkinB=Gray"" charset=""utf-8""></script>"
elseif zxys="28" then
messageskin=messageskin&"<script type=""text/javascript"" src=""scripts/Message.js?SkinA=special&SkinB=Red"" charset=""utf-8""></script>"
End If
End If
end function
function Map(MapW,MapH)
Map=Map&"<iframe margin=""0"" src="""&sitepath&"Map/"" width="""&MapW&""" height="""&MapH&""" scrolling=""no"" frameborder=""0""></iframe>"
end function
function nymap(MapW,MapH)
nymap=nymap&"<iframe margin=""0"" src="""&sitepath&"Map/ny.asp?id="&yxwlid&""" width="""&MapW&""" height="""&MapH&""" scrolling=""no"" frameborder=""0""></iframe>"
end function
function yxwlmap
 
	yxwlmap=yxwlmap& "              <object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0"" width=""650"" height=""320"">"&VbCrLf
    yxwlmap=yxwlmap& "                <param name=""movie"" value="""&sitepath&"Images/map.swf"" />"&VbCrLf
    yxwlmap=yxwlmap& "                <param name=""quality"" value=""high"" />"&VbCrLf
    yxwlmap=yxwlmap& "                <param name=""wmode"" value=""opaque"" />"&VbCrLf
    yxwlmap=yxwlmap& "                <embed src="""&sitepath&"Images/map.swf"" quality=""high"" wmode=""opaque"" pluginspage=""http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""650"" height=""320""></embed>"&VbCrLf
    yxwlmap=yxwlmap& "              </object>"&VbCrLf
end function

'建立购物车
Sub CreateCart()
    If IsArray(Session("Cart")) = False Then
        Dim aryCart()
        Redim aryCart(6,0)     
		Session("Cart") = aryCart
    End If
End Sub

'判断是否为数组,即购物车是否是空的
Function CheckCart()
    If IsArray(Session("Cart")) Then
        CheckCart = True
    Else
        CheckCart = False
    End If
End Function

'判断该产品ID是否已经在购物车里
Function CheckItem(ID)
    Dim aryCart,findFlag,i
    If CheckCart = True Then 'True为购物车里有东西
        aryCart = Session("Cart")
        findFlag = False
        For i = LBound(aryCart,2) To UBound(aryCart,2) 
            If CLng(aryCart(0,i)) = ID Then
                findFlag = True
            Exit For
            End If
        Next
    CheckItem = findFlag
    End If
End Function

'删除产品
Sub RemoveItem(ID)
    ID = Clng(ID)
    If not CheckItem(ID) Then '如果无该产品ID,则退出函数;
        Exit Sub
    End If
    Dim i,intPos,aryCart,newSize
    aryCart = Session("Cart") '先提取购物车中现有的的产品
    For i = LBound(aryCart,2) To UBound(aryCart,2)
        If Clng(aryCart(0,i)) = ID Then
            intPos = i
            Exit For
        End If
    Next
    
    For i = intPos To UBound(aryCart,2) - 1
        If Not aryCart(0,i) = "" Then
            aryCart(0,i) = aryCart(0,i+1)
            aryCart(1,i) = aryCart(1,i+1)
            aryCart(2,i) = aryCart(2,i+1)
            aryCart(3,i) = aryCart(3,i+1)
            aryCart(4,i) = aryCart(4,i+1)
			aryCart(5,i) = aryCart(5,i+1)
			aryCart(6,i) = aryCart(6,i+1)
        End If
    Next
    newSize = UBound(aryCart,2) - 1   '减小数组的2维的值,用preserve参数删除产品
    ReDim preserve aryCart(6,newSize) '//////////////////////////////////////////////////////////////
    Session("Cart") = aryCart  '重新赋值给数组;
End Sub

'更新[数量]和[说明信息]
Sub UpdateItem(ID,Num,name,retailprice,price,img,sf)
    ID = Clng(ID)
    Num = Clng(Num)
    Dim aryUpdateCart,i
    aryUpdateCart = Session("Cart") '先提取购物车中现有的的产品
    For i = LBound(aryUpdateCart,2) To UBound(aryUpdateCart,2)
        If CLng(aryUpdateCart(0,i)) = ID Then
            aryUpdateCart(1,i) = Num
            aryUpdateCart(2,i) = name
            aryUpdateCart(3,i) = retailprice
            aryUpdateCart(4,i) = price
			aryUpdateCart(5,i) = img
			aryUpdateCart(6,i) = sf
            Session("Cart") = aryUpdateCart  '重新赋值给数组;
        Exit For
        End If
    Next
End Sub

'添加产品
Sub AddItem(ID,Num,name,retailprice,price,img,sf)
    ID = Clng(ID)
    Num = Clng(Num)
    Dim btnCartStatus,aryAddCart,newSize,i
    btnCartStatus = CheckCart
    If btnCartStatus = False Then '购物车不存在,为空,即第一次添加产品;
        CreateCart '建立购物车,等价于Call CreateCart()
        aryAddCart = Session("Cart")   'Session存储信息
        aryAddCart (0,0) = ID          '将[ID]存储到1维的第一个元素;
        aryAddCart (1,0) = Num         '将[数量]存储到1维的第2个元素;
        aryAddCart (2,0) = name       '将[名称]存储到1维的第3个元素;
        aryAddCart (3,0) = retailprice '将[价格]存储到1维的第4个元素;
        aryAddCart (4,0) = price      '将[会员价格]存储到1维的第4个元素;
		aryAddCart (5,0) = img
		aryAddCart (6,0) = sf
        Session ("Cart") = aryAddCart  '重新赋值给数组;
        Exit Sub
    ElseIf btnCartStatus = True Then '购物车不为空,有以下2种情况;
        If CheckItem(ID) = True Then  '此产品已经在购物车中,在现有的数量上+num
            aryAddCart = Session("Cart") '先提取购物车中现有的的产品
            For i = LBound(aryAddCart,2) To UBound(aryAddCart,2)
                If Clng(aryAddCart(0,i)) = ID Then
                    aryAddCart(1,i) = aryAddCart(1,i) + Num '现有数量上+ num
                    If aryAddCart(2,i) = "" Then  
                        aryAddCart(2,i) = name
                    End If

					If aryAddCart(3,i) = "" Then  
                        aryAddCart(3,i) = retailprice
                    End If

					If aryAddCart(4,i) = "" Then  
                        aryAddCart(4,i) = price
                    End If
					If aryAddCart(5,i) = "" Then  
                        aryAddCart(5,i) = img
                    End If
					If aryAddCart(6,i) = "" Then  
                        aryAddCart(6,i) = sf
                    End If

					Session("Cart") = aryAddCart '重新赋值给数组;
                    Exit For
                End If
            Next
        ElseIf CheckItem(ID) = False Then  '此产品不在购物车中,往里添加;
            aryAddCart = Session("Cart") '先提取购物车中现有的的产品
            newSize = UBound(aryAddCart,2)+1 '增加数组的2维的值,用preserve参数增加产品
            ReDim preserve aryAddCart(6,newSize) '////////////////////////////////////////////////////////// preserve:当更改现有数组最后一维的大小时保留数据,可查阅REDIM语句
            
            '将新添的产品信息(ID,NUM,name,retailprice)赋值给数组里2维的值为newsize元素中去 
            aryAddCart(0,newSize) = ID
            aryAddCart(1,newSize) = Num
            aryAddCart(2,newSize) = name
            aryAddCart(3,newSize) = retailprice
            aryAddCart(4,newSize) = price
			aryAddCart(5,newSize) = img
			aryAddCart(6,newSize) = sf
            Session("Cart") = aryAddCart '重新赋值给数组;
            Exit Sub
        End If
    End If
End Sub
			
function mobilemap
mobilemap=mobilemap&"<style type=""text/css"">"
mobilemap=mobilemap&"#container {width: 100%;height: 500px;overflow: hidden;margin:0;}"
mobilemap=mobilemap&"</style>"
mobilemap=mobilemap&"<div id=""container""></div>"
end function

Function ProductInfoItemValue()
		On Error Resume Next 
		Dim Arr_name,Arr_value
		Arr_name = Split(rspro("pro_name"),"$")
		Arr_value = Split(rspro("pro_info"),"$$$")
		For i = 0 To UBound(Arr_name)
		If Arr_value(i) <> "" Then 
		ProductInfoItemValue = ProductInfoItemValue&"<li><b>"& Arr_name(i) &"：</b>"& Arr_value(i) &"</li>"
		End If 
		Next 
End Function 

Function ProductInfoItemSpecification()
If rspro("P_SpecificationName") = "" Or IsNull(rspro("P_SpecificationName")) Then 
Exit Function 
End If 
if session("MemName")="" or session("GroupID")="" or session("GroupID")=999177 or session("MemLogin")<>"Succeed" then
If InStr(rspro("P_SpecificationName"),",") = 0 Then 
q = 0
P_SpecificationValue = Replace(rspro("P_SpecificationValue"),"|",",")
If Right(P_SpecificationValue,1) = "|" Then P_SpecificationValue = cut_Right(P_SpecificationValue)
If Right(P_SpecificationValue,1) = "," Then P_SpecificationValue = cut_Right(P_SpecificationValue)
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &":</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and Lang='"&Language&"' and id in("& P_SpecificationValue &")")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 
P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);""><img src="""&sitepath&"/"&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);"">"& rs(0) &"</a></div>"
End If 
rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& id &""" id=""hidv_0001_"& id &"_0""></dl>"
Else
For q = 0 To UBound(Split(rspro("P_SpecificationName"),","))
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &"：</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"' and id in("& Split(rspro("P_SpecificationValue"),"|")(q) &")")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 

P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);""><img src="""&sitepath&""&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rspro("hyprice") &"','0001',0);"">"& rs(0) &"</a></div>"
End If 

rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& proid &""" id=""hidv_0001_"& proid &"_"& q &"""></dl>"
Next 
End If 



else

set rsjg2=server.CreateObject("adodb.recordset")
sqljg2="select * from hygroup where groupid='"&session("GroupID")&"' and Lang='"&Language&"'"
rsjg2.open sqljg2,conn,1,1

set rsjg12=server.CreateObject("adodb.recordset")
sqljg12="select * from product_price where grade_id="&rsjg2("GroupLevel")&" and product_id="&proid&" and Lang='"&Language&"'"
rsjg12.open sqljg12,conn,1,1

If InStr(rspro("P_SpecificationName"),",") = 0 Then 
q = 0
P_SpecificationValue = Replace(rspro("P_SpecificationValue"),"|",",")
If Right(P_SpecificationValue,1) = "|" Then P_SpecificationValue = cut_Right(P_SpecificationValue)
If Right(P_SpecificationValue,1) = "," Then P_SpecificationValue = cut_Right(P_SpecificationValue)
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &"：</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(rspro("P_SpecificationName"),0) &" and Lang='"&Language&"' and id in("& P_SpecificationValue &")")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 
P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price") &"','0001',0);""><img src="""&sitepath&"/"&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price")  &"','0001',0);"">"& rs(0) &"</a></div>"
End If 
rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& id &""" id=""hidv_0001_"& id &"_0""></dl>"
Else
For q = 0 To UBound(Split(rspro("P_SpecificationName"),","))
Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"'")
If Not rs.eof Then 
SName = rs(0)
rsprotemp = rsprotemp &"<dl><dt>"& SName &"：</dt>"
End If 
rs.close : Set rs = Nothing
Set rs = conn.execute("select s_value,id from [SelectValue] where s_id = "& checknum(Split(rspro("P_SpecificationName"),",")(q),0) &" and Lang='"&Language&"' and id in("& Split(rspro("P_SpecificationValue"),"|")(q) &")")
If Not rs.eof Then 
jj = 0
rsprotemp = rsprotemp &"<dd>"
Do While Not rs.eof 

P_SpecificationPic = rspro("P_SpecificationPic")
If P_SpecificationPic = "" Or IsNull(P_SpecificationPic) Then P_SpecificationPic = "|||||||||||||||||||||||||||||||||||"
k = getN(Replace(rspro("P_SpecificationValue"),"|",","),rs(1),",")
if Split(P_SpecificationPic,"|")(k) <> "" Then
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva-img"" title="""& rs(0) &""" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price")  &"','0001',0);""><img src="""&sitepath&""&Split(P_SpecificationPic,"|")(k)&""" /><span></span></a></div>"
Else 
rsprotemp = rsprotemp &"<div><a id=""sv_"& id &"_"& q &"_"& jj &""" name=""sv_n_"& id &"_"& q &""" class=""spva"" onclick=""select_sp('"& id &"','"& q &"','"& jj &"','"& SName &"|"& rs(0) &"','"& rsjg12("price")  &"','0001',0);"">"& rs(0) &"</a></div>"
End If 

rs.movenext
jj = jj + 1
Loop
rsprotemp = rsprotemp &"</dd>"
End If 
rs.close : Set rs = Nothing
rsprotemp = rsprotemp &"<input type=""hidden"" name=""hidv_0001_"& proid &""" id=""hidv_0001_"& proid &"_"& q &"""></dl>"
Next 
End If 
end if
ProductInfoItemSpecification = rsprotemp
End Function 



Function SelectValue()

pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")



if pone="" and ptwo="" and pthree="" then
else
		IsSelectType = 0
		IsSelectValue = 0
		SelectValue = "<ul class=""selectvalue"">"


if pone<>"" then
		sql = "select s_id from [SelectType] where zmone="&pone&" and Lang='"&Language&"'  order by s_order desc "
end if		
if pone<>"" and ptwo<>"" then
		sql = "select s_id from [SelectType] where zmone="&pone&" and zmtwo="&ptwo&" and Lang='"&Language&"' order by s_order desc "
end if			
if pone<>"" and ptwo<>"" and pthree<>"" then
		sql = "select s_id from [SelectType] where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and Lang='"&Language&"' order by s_order desc "
end if		
		
		
		Set rs = conn.execute(sql)
		If Not rs.eof Then 
			IsSelectType = 1
			DataSelectType = rs.getrows()
		End If 
		rs.close : Set rs = Nothing 
		
		
		If IsSelectType = 1 Then 
			For i = 0 To UBound(DataSelectType,2)
				SelectTypeId = SelectTypeId & DataSelectType(0,i)
				If i < UBound(DataSelectType,2) Then SelectTypeId = SelectTypeId & ","
			Next 
			sql = "select s_name,s_id,s_name from [SelectName] where s_mode = 0 and typeid in("& SelectTypeId &") and Lang='"&Language&"' order by s_order asc"
			Set rs = conn.execute(sql)
			If Not rs.eof Then 
				IsSelectName = 1
				DataSelectName = rs.getrows()
			End If 
			rs.close : Set rs = Nothing 
			If IsSelectName = 1 Then 
			For nn = 0 To UBound(DataSelectName,2)
			SelectValue = SelectValue &"<li><label class=""select_name""><span>"& DataSelectName(0,nn) &"：</span></label><label class=""select_value""><a href="""& paramurl(UnionUrl(DataSelectName(2,nn) &"_0","")) &""""
			
		
			
			If value<>"" and InStr(value,DataSelectName(2,nn) &"_0") > 0 then
		
			SelectValue = SelectValue &" class=""selectall"""
			SelectValue = SelectValue &">"&word(517)&"</a>"
			end if
			If value<>"" and InStr(value,DataSelectName(2,nn) &"_0") = 0 Then 
			if InStr(value,DataSelectName(2,nn) &"_") < 1 then
			
			SelectValue = SelectValue &" class=""selectall"""
			end if
			SelectValue = SelectValue &">"&word(517)&"</a>"
            end if
		
	
			
			
			if value="" then
			SelectValue = SelectValue &" class=""selectall"""
			if InStr(value,DataSelectName(2,nn) &"_") = 0 then
			
			SelectValue = SelectValue &" class=""selectall"""
			end if
			SelectValue = SelectValue &">"&word(517)&"</a>"
			end if
			
			
	
			
			
			
			
		
		   
		
			
         
			
			
			
				
			
			
			
			sql = "select s_value,id from [SelectValue] where s_id = "& DataSelectName(1,nn) &" and Lang='"&Language&"' order by s_order asc"
			Set rs = conn.execute(sql)
			If Not rs.eof Then 
				IsSelectValue = 1
				DataSelectValue = rs.getrows()
			End If 
			rs.close : Set rs = Nothing 
			If IsSelectValue = 1 Then 
			For vv = 0 To UBound(DataSelectValue,2)
				If Not IsNull(Trim(DataSelectValue(0,vv))) And Trim(DataSelectValue(0,vv)) <> "" Then
					SelectValue = SelectValue &"<a href="""& paramurl(UnionUrl(DataSelectName(2,nn)&"_"& DataSelectValue(1,vv),"")) &""""
					If InStr(value,DataSelectName(2,nn) &"_"& DataSelectValue(1,vv)) > 0 Then SelectValue = SelectValue &" class=""selectone"""
					SelectValue = SelectValue &">"& DataSelectValue(0,vv) &"</a>"
				End If 
			Next 
			End If 
			SelectValue = SelectValue &"</label></li>"
			Next 
			End If 
		End If 
		SelectValue = SelectValue &"</ul>"
		end if
End Function 

Function UnionUrl(param,param2)
		If param = "" Then Exit Function 
	pone=request.QueryString("pone")
ptwo=request.QueryString("ptwo")
pthree=request.QueryString("pthree")
m=request.QueryString("m")
if pone<>"" then
        url = ""&sitepath&"product/?pone="&pone&"&m="&m&"&Language="&Language&""
		end if
		
		
		if pone<>"" and ptwo<>"" then
        url = ""&sitepath&"product/?pone="&pone&"&ptwo="&ptwo&"&m="&m&"&Language="&Language&""
		end if
		if pone<>"" and ptwo<>"" and pthree<>"" then
        url = ""&sitepath&"product/?pone="&pone&"&ptwo="&ptwo&"&pthree="&pthree&"&m="&m&"&Language="&Language&""
		end if
		
		
		
		
		
		If Split(param,"_")(1) = "0" Then 
				If InStr(value,"|") > 0 Then 
						arr_value = Split(value,"|")
						paramorder = GetArrayOrder(param)
						If paramorder > -1 Then
								For i = 0 To paramorder-1
										new_param = new_param & arr_value(i) &"|"
								Next 
								new_param = new_param & param
								For i = paramorder+1 To UBound(arr_value)
										new_param = new_param &"|"& arr_value(i)
								Next 
								param = new_param
						Else
								If value <> "" Then 
										param = value &"|"& param
								Else
										param = param
								End If 
						End If 
				Else 
						param = ""
				End If 
		Else 
				If InStr(value,"|") > 0 Then 
						arr_value = Split(value,"|")
						paramorder = GetArrayOrder(param)
						If paramorder > -1 Then 
								For i = 0 To paramorder-1
										new_param = new_param & arr_value(i) &"|"
								Next 
								new_param = new_param & param
								For i = paramorder+1 To UBound(arr_value)
										new_param = new_param &"|"& arr_value(i)
								Next 
								param = new_param
						Else
								If value <> "" Then 
										param = value &"|"& param
								Else
										param = param
								End If 
						End If 
				Else 
						paramorder = GetArrayOrder(param)
						If paramorder > -1 Then 
								If value <> "" Then 
										If CStr(Split(value,"_")(0)) = CStr(Split(param,"_")(0)) Then 
												param = param
										Else 
												param = value &"|"& param
										End If 
								Else
										param = param
								End If 
						Else 
								param = value &"|"& param
						End If 
'						If value <> "" Then param = value &"|"& param
				End If 
		End If 

		If InStr(url,param2) = 0 Then 
				If InStr(url,"?") > 0 Then 
						url = url &"&"& param2
				Else
						url = url &"?"& param2
				End If 
		End If 
		If InStr(url,"?") > 0 Then 
				UnionUrl = url &"&value="& server.urlencode(param)
		Else
				UnionUrl = url &"?value="& server.urlencode(param)
		End If 
'		If param2 = "brand=" Then 
'				arr_param2 = Split(UnionUrl,param2)
'				UnionUrl = arr_param2(0)&"&brand="
'				If InStr(arr_param2(1),"&") > 0 Then UnionUrl = UnionUrl &"&"& Split(arr_param2(1),"&")(1)
'		End If 
End Function 

Function GetArrayOrder(param)
		GetArrayOrder = -1
		If InStr(value,"|") > 0 Then 
				arr_value = Split(value,"|")
				For i = 0 To UBound(arr_value)
						If InStr(arr_value(i),Split(param,"_")(0) &"_") > 0 Then 
								GetArrayOrder = i
								Exit Function 
						End If 
				Next 
		Else
				GetArrayOrder = 0
		End If 
End Function 


Function paramurl(url)
	If request("perpage") <> "" And InStr(url,"perpage=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&perpage="& request("perpage")
		Else
			url = url &"?perpage="& request("perpage")
		End If 
	End If 
	If request("color") <> "" And InStr(url,"color=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&color="& request("color")
		Else
			url = url &"?color="& request("color")
		End If 
	End If 
	If request("listprice") <> "" And InStr(url,"listprice=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&listprice="& request("listprice")
		Else
			url = url &"?listprice="& request("listprice")
		End If 
	End If 
	If request("listtype") <> "" And InStr(url,"listtype=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&listtype="& request("listtype")
		Else
			url = url &"?listtype="& request("listtype")
		End If 
	End If 
	If request("listtime") <> "" And InStr(url,"listtime=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&listtime="& request("listtime")
		Else
			url = url &"?listtime="& request("listtime")
		End If 
	End If 
	If request("listview") <> "" And InStr(url,"listview=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&listview="& request("listview")
		Else
			url = url &"?listview="& request("listview")
		End If 
	End If 
	If request("listbuy") <> "" And InStr(url,"listbuy=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&listbuy="& request("listbuy")
		Else
			url = url &"?listbuy="& request("listbuy")
		End If 
	End If 
	If request("pricetype") <> "" And InStr(url,"pricetype=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&pricetype="& request("pricetype")
		Else
			url = url &"?pricetype="& request("pricetype")
		End If 
	End If 
	If request("price1") <> "" And InStr(url,"price1=") =< 0 And InStr(url,"price2=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&price1="& request("price1") &"&price2="& request("price2")
		Else
			url = url &"?price1="& request("price1") &"&price2="& request("price2")
		End If 
	End If 
	If request("material") <> "" And InStr(url,"material=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&material="& request("material")
		Else
			url = url &"?material="& request("material")
		End If 
	End If 
	If request("design") <> "" And InStr(url,"design=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&design="& request("design")
		Else
			url = url &"?design="& request("design")
		End If 
	End If 
	If request("brand") <> "" And InStr(url,"brand=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&brand="& server.urlencode(request("brand"))
		Else
			url = url &"?brand="& server.urlencode(request("brand"))
		End If 
	End If 
	If request("sle_type") <> "" And InStr(url,"sle_type=") =< 0 Then
		If InStr(url,"?") > 0 Then 
			url = url &"&sle_type="& request("sle_type")
		Else
			url = url &"?sle_type="& request("sle_type")
		End If 
	End If 
	paramurl = url
End Function 

'--------产品内容页会员各个组价格展示-------------
function group(style)
set RSZM=server.createobject("adodb.recordset")
ZM="select * from product_price where product_id="&proid&" and Lang='"&Language&"' order by grade_id asc"
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
RSZM.open ZM,conn,1,1		
For i = 1 To RSZM.RecordCount
set rsc=server.CreateObject("adodb.recordset")
rsc.open "select * from  hygroup where grouplevel<>0 and Lang='"&Language&"'",conn,1,1
while not rsc.eof
if rszm("grade_id")=rsc("GroupLevel") then
hyz=rsc("groupname")
end if
rsc.movenext
wend
rsc.close
set rsc=nothing
linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)

tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{会员组价格}",RSZM("price"))
tempListStr = replace(tempListStr,"{会员组名称}",hyz)
returnStr = returnStr & tempListStr
RSZM.movenext
next
RSZM.close 
set RSZM=nothing
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)
group = returnStr
end function

function diyform(id,ys,tys,Dt,url)
		If tys="" Then
		tys=1 
	End if
set rsObj1=server.CreateObject("adodb.recordset")
Sqlp ="select * from FreeForm Where id="&id&""     
rsObj1.open sqlp,conn,1,1 
If rsObj1.EOF Then
WriteMsg1 1,""&word(43)&"","","wrong"
End If
if rsObj1("IsStatus")=0 then 
WriteMsg1 1,""&word(44)&"","","wrong"
else
mystr=mystr&"<form method=""post"" action="""&sitepath&"inc/diyform/?id="&id&"&language="&language&"&url="&url&"""  class=""checkform"">"
mystr=mystr&"<table class="""&ys&""">"

set rsObj=server.CreateObject("adodb.recordset")
Sqlp1 ="select * from FreeField where ChannelId="&id&" order by px asc"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
select case rsObj("FieldType")
case 1
mystr=mystr&" <tr >"
if rsObj1("zxs")=1 then
mystr=mystr&" <td class=""zkd"">"&rsObj("Alias")&"：</td>"
end if
mystr=mystr&"<td class=""ckd""><input type='text' name='"&rsObj("FieldName")&"' class=""inputxt""  "

if rsObj1("txs")=0 then
mystr=mystr&"TIP='"&rsObj("FieldNull")&"' value='"&rsObj("FieldNull")&"'"
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=0 then
mystr=mystr&"datatype=""*"" nullmsg="""&rsObj("Alias")&""&word(45)&"""" 	
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=1 then
mystr=mystr&"datatype=""e"" nullmsg="""&rsObj("Alias")&""&word(45)&""" errormsg="""&word(46)&""" " 	
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=2 then
mystr=mystr&"datatype=""*15-18"" nullmsg="""&rsObj("Alias")&""&word(45)&""" errormsg="""&word(47)&""" " 	
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=3 then
mystr=mystr&"datatype=""n"" nullmsg="""&rsObj("Alias")&""&word(45)&""" errormsg="""&word(48)&""" " 	
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=4 then
mystr=mystr&"datatype=""m"" nullmsg="""&rsObj("Alias")&""&word(45)&""" errormsg="""&word(49)&""" " 	
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=5 then
mystr=mystr&"datatype=""p"" nullmsg="""&rsObj("Alias")&""&word(45)&""" errormsg="""&word(50)&""" " 	
end if
if rsObj("IsNotNull")=0 and rsObj("yanzheng")=1 then
mystr=mystr&"datatype=""e""  errormsg="""&word(46)&""" ignore=""ignore"" " 	
end if
if rsObj("IsNotNull")=0 and rsObj("yanzheng")=2 then
mystr=mystr&"datatype=""*15-18""  errormsg="""&word(47)&""" ignore=""ignore"" " 	
end if
if rsObj("IsNotNull")=0 and rsObj("yanzheng")=3 then
mystr=mystr&"datatype=""n""  errormsg="""&word(48)&""" ignore=""ignore"" " 	
end if
if rsObj("IsNotNull")=0 and rsObj("yanzheng")=4 then
mystr=mystr&"datatype=""m""  errormsg="""&word(49)&""" ignore=""ignore"" " 	
end if
if rsObj("IsNotNull")=0 and rsObj("yanzheng")=5 then
mystr=mystr&"datatype=""p""  errormsg="""&word(50)&""" ignore=""ignore"" " 	
end if
mystr=mystr&"/>"
if rsObj("IsNotNull")=1 then
mystr=mystr&"<span class=""fh""> "
mystr=mystr&"*"
mystr=mystr&"</span> "
end if
mystr=mystr&"</td>" 





mystr=mystr&" <td class=""ykd""><div class=""Validform_checktip"">"
if rsObj1("txs")=1 then
mystr=mystr&""&rsObj("FieldNull")&""
end if
mystr=mystr&"</div></td>"
mystr=mystr&" </tr >"

case 2

mystr=mystr&" <tr >"
if rsObj1("zxs")=1 then
mystr=mystr&" <td class=""zkd"">"&rsObj("Alias")&"：</td>"
end if
mystr=mystr&" <td class=""ckd"">"
rdx=rsObj("ListVal")
rdx1=Split(rdx,"|")
for i=0 to ubound(rdx1)
mystr=mystr&"<input type='radio' value='"&rdx1(i)&"' name='"&rsObj("FieldName")&"' class='radio' "

if rsObj("IsNotNull")=1 and rsObj("yanzheng")=0 then
mystr=mystr&"datatype=""*"" nullmsg="""&rsObj("Alias")&""&word(45)&"""" 	
end if

mystr=mystr&"/> "&rdx1(i)&" " 
next
mystr=mystr&" </td><td class=""ykd""><div class=""Validform_checktip"">"
if rsObj1("txs")=1 then
mystr=mystr&""&rsObj("FieldNull")&""
end if
mystr=mystr&"</div></td>"
mystr=mystr&" </tr >"

	
case 3

mystr=mystr&" <tr >"
if rsObj1("zxs")=1 then
mystr=mystr&" <td class=""zkd"">"&rsObj("Alias")&"：</td>"
end if
mystr=mystr&" <td class=""ckd"">"
cdx=rsObj("ListVal")
cdx1=Split(cdx,"|")
for i=0 to ubound(cdx1)
mystr=mystr&"<input type='checkbox' name='"&rsObj("FieldName")&"' value='"&cdx1(i)&"' style=""width:13px;height:15px;border:none"" class=""checkbox"" " 
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=0 then
mystr=mystr&"datatype=""*"" nullmsg="""&rsObj("Alias")&""&word(45)&"""" 	
end if
mystr=mystr&"/> "&cdx1(i)&" " 
next
mystr=mystr&" </td><td class=""ykd""><div class=""Validform_checktip"">"
if rsObj1("txs")=1 then
mystr=mystr&""&rsObj("FieldNull")&""
end if
mystr=mystr&"</div></td>"
mystr=mystr&" </tr >"



case 4

mystr=mystr&" <tr >"
if rsObj1("zxs")=1 then
mystr=mystr&" <td class=""zkd"">"&rsObj("Alias")&"：</td>"
end if
mystr=mystr&" <td class=""ckd"">"
mystr=mystr&"<select size=""1"" name='"&rsObj("FieldName")&"' class=""select"" "
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=0 then
mystr=mystr&"datatype=""*"" nullmsg="""&rsObj("Alias")&""&word(45)&"""" 	
end if
mystr=mystr&" > "
ldx=rsObj("ListVal")
ldx1=Split(ldx,",")
for i=0 to ubound(ldx1)
ldx2=Split(ldx1(i),"|")

mystr=mystr&"<option   value='"&ldx2(1)&"' />"&ldx2(0)&"</option>" 
next
mystr=mystr&" </select >"
	
mystr=mystr&" </td><td class=""ykd""><div class=""Validform_checktip"">"
if rsObj1("txs")=1 then
mystr=mystr&""&rsObj("FieldNull")&""
end if
mystr=mystr&"</div></td>"
mystr=mystr&" </tr >"


case 5
mystr=mystr&" <tr >"
if rsObj1("zxs")=1 then
mystr=mystr&" <td class=""zkd"">"&rsObj("Alias")&"：</td>"
end if
mystr=mystr&"<td class=""ckd""><textarea class=""textarea""  name='"&rsObj("FieldName")&"'" 
if rsObj1("txs")=0 then
mystr=mystr&"tip='"&rsObj("FieldNull")&"'"
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=0 then
mystr=mystr&"datatype=""*"" nullmsg="""&rsObj("Alias")&""&word(45)&"""" 	
end if
mystr=mystr&" />"
if rsObj1("txs")=0 then
mystr=mystr&"Value='"&rsObj("FieldNull")&"'"
end if

mystr=mystr&"</textarea>"
mystr=mystr&" </td><td class=""ykd""><div class=""Validform_checktip"">"
if rsObj1("txs")=1 then
mystr=mystr&""&rsObj("FieldNull")&""
end if
mystr=mystr&"</div></td>"
mystr=mystr&" </tr >"

case 6
mystr=mystr&" <tr >"
if rsObj1("zxs")=1 then
mystr=mystr&" <td class=""zkd"">"&rsObj("Alias")&"：</td>"
end if
mystr=mystr&"<td class=""ckd""><input type='text' name='"&rsObj("FieldName")&"' onclick=""laydate({istime: true, format: '"&dt&"',festival: true})"" class=""inputxt laydate-icon"" " 
if rsObj1("txs")=0 then
mystr=mystr&"TIP='"&rsObj("FieldNull")&"' value='"&rsObj("FieldNull")&"'"
end if
if rsObj("IsNotNull")=1 and rsObj("yanzheng")=0 then
mystr=mystr&"datatype=""*"" nullmsg="""&rsObj("Alias")&""&word(45)&"""" 	
end if
mystr=mystr&" />"
mystr=mystr&" </td><td class=""ykd""><div class=""Validform_checktip"">"
if rsObj1("txs")=1 then
mystr=mystr&""&rsObj("FieldNull")&""
end if
mystr=mystr&"</div></td>"
mystr=mystr&" </tr >"

end select
rsObj.movenext
loop
end if
if rsObj1("IsCode")=1 then
mystr=mystr&" <tr>"
if rsObj1("zxs")=1 then
mystr=mystr&" <td class=""zkd"">"&word(51)&"</td>"
end if
mystr=mystr&"<td class=""ckd"">"
mystr=mystr&"<input name=""captchacode"" id=""captchacode"" type=""text"" style=""width:65px""  datatype=""*""  nullmsg="""&word(52)&""" class=""inputxt"" ajaxurl="""&sitepath&"ajax/yzm.asp"""

if rsObj1("txs")=0 then
mystr=mystr&"TIP='"&word(493)&"' value='"&word(493)&"'"
end if

mystr=mystr&"/> <img id=""imgCaptcha"" src="""&sitepath&"inc/captcha.asp"" /> <a href=""###"" onClick=""RefreshImage('imgCaptcha')""><img src="""&sitepath&"images/reload.png"" width=""20"" height=""20"" /></a>"

mystr=mystr&" </td><td class=""ykd""><div class=""Validform_checktip"">"
if rsObj1("txs")=1 then
mystr=mystr&""&word(494)&""
end if
mystr=mystr&"</div></td>"
mystr=mystr&"</tr>"
end if
mystr=mystr&"<tr>"
mystr=mystr&"<td colspan=""3"" class=""last"">"
mystr=mystr&"<input name=""pid"" type=""hidden"" id=""pid"" value="""" />"   
mystr=mystr&"<input type=""submit"" class=""tjan"" value="""&word(53)&""" name=""B1"" />" 
mystr=mystr&"<input name="""" type=""reset"" value="""&word(54)&""" class=""czan"" />"
mystr=mystr&"</td>"
mystr=mystr&"</tr>"
mystr=mystr&"</table>"

mystr=mystr&"</form>"
end if
rsObj.close
set rsObj=nothing
rsObj1.close
set rsObj1=Nothing
mystr=mystr&"<script>"
mystr=mystr&";!function(){"
If tys=1 then
mystr=mystr&"laydate.skin('default');"
End If
If tys=2 then
mystr=mystr&"laydate.skin('dahong');"
End If
If tys=3 then
mystr=mystr&"laydate.skin('molv');"
End If
If tys=4 then
mystr=mystr&"laydate.skin('danlan');"
End If
If tys=5 then
mystr=mystr&"laydate.skin('huanglv');"
End If
If tys=6 then
mystr=mystr&"laydate.skin('qianhuang');"
End If
If tys=7 then
mystr=mystr&"laydate.skin('yahui');"
End If
If tys=8 then
mystr=mystr&"laydate.skin('yalan');"
End if
mystr=mystr&"}();"
mystr=mystr&"</script>"
diyform=mystr
end function

function formlist(id,n,w,style,msql,D,D1,P)
set RSform=server.createobject("adodb.recordset")
Sqlp1 ="select * from FreeForm Where Id="&id&""     
RSform.open sqlp1,conn,1,1 
FormTableName=rsform("FormTableName")
fsh=rsform("fsh")
set RSZM=server.createobject("adodb.recordset")
if fsh=2 then
zm = "select * from "&FormTableName&"  "
if msql<>"no" then
zm=zm&" where "&msql&""
end if
else
zm = "select * from "&FormTableName&" where sh=1 "
if msql<>"no" then
zm=zm&" and "&msql&""
end if
end if	
zm=zm&" order by "&p&""
'--------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""&word(649)&""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'--------------调用样式结束
RSZM.open ZM,conn,1,1
If not( rszm.bof and rszm.EOF) Then
rszm.PageSize = n
iCount = rszm.RecordCount 
iPageSize = rszm.PageSize
maxpage = rszm.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rszm.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If 		
For i = 1 To x
'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)

tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{提交IP}",RSZM("IP"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("adddate"))
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
Sqlp1 ="select * from FreeField where ChannelId="&id&" order by px asc"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("FieldName")
zdyzd3=rsObj("Alias")
zdyzd2=RSZM(""&rsObj("FieldName")&"")
tempListStr = replace(tempListStr,"{"&zdyzd3&"}",""&zdyzd2&"")
rsObj.movenext
loop
end if
rsObj.close
set rsObj=nothing
'自定义字段结束
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间}",FormatTime(RSZM("adddate"),D))
end if
if DateDiff("d",""&rszm("adddate")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{更新时间1}","<font color="&sjys&">"&FormatTime(RSZM("adddate"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{更新时间1}",FormatTime(RSZM("adddate"),D1))
end if
if trim(rszm("replay")&"")="" then
replayContent=""&word(495)&""
else
replayContent=rszm("replay")
end if


tempListStr = replace(tempListStr,"{回复内容}",replayContent)
if rszm("replaytime")<>"" then
tempListStr = replace(tempListStr,"{回复时间}",rszm("replaytime"))
else
tempListStr = replace(tempListStr,"{回复时间}",""&word(495)&"")
end if
returnStr = returnStr & tempListStr
RSZM.movenext
next
rszm.Close
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)
if instr(returnStr,"{分页}") > 0 then
if w=1 then
returnStr = replace(returnStr,"{分页}","<div id=""pagepc"" class=""page""></div>")
else
returnStr = replace(returnStr,"{分页}","<div id=""pagewap"" class=""page""></div>")
end if
end if 

if w=1 then
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagepc',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(475)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(476)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(477)&","& vbCrLf
returnStr=returnStr&  "next: "&word(478)&","& vbCrLf
if word(479)<>"" then
returnStr=returnStr&  "last: "&word(479)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(480)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(481)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
else
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pagewap',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(482)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(483)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(484)&","& vbCrLf
returnStr=returnStr&  "next: "&word(485)&","& vbCrLf
if word(486)<>"" then
returnStr=returnStr&  "last: "&word(486)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(496)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(497)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
end if
formlist = returnStr
else
formlist = word(0)
end if
end function

	
function memberguest(style,D,D1,P)
if not (session("MemName")<>"" and session("MemLogin")="Succeed") then 
writemsg1 2,""&word(95)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","wrong"
response.end
end if

Set rs1 = server.CreateObject("adodb.recordset")
sql1="select * from members where MemName='"&session("MemName")&"' and Lang='"&Language&"'"
rs1.Open sql1, conn, 1, 1
id2=rs1("id")


set RSZM=server.createobject("adodb.recordset")
if ly=2 then
zm = "select * from ly where MemID="&id2&" and Lang='"&Language&"'  order by "&p&"  "
else
zm = "select * from ly where sh=1 and MemID="&id2&" and Lang='"&Language&"' order by "&p&"  "
end if

'--------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""&word(649)&""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'--------------调用样式结束
RSZM.open ZM,conn,1,1
If not( rszm.bof and rszm.EOF) Then
rszm.PageSize = hylyinfo
iCount = rszm.RecordCount 
iPageSize = rszm.PageSize
maxpage = rszm.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rszm.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If 		
For i = 1 To x
'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束

hqlydz= gethttp("http://whois.pconline.com.cn/ip.jsp?ip="&RSZM("lyip")&"","")

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)

tempListStr = replace(tempListStr,"{留言ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{发布时间}",RSZM("sj"))
tempListStr = replace(tempListStr,"{留言用户}",RSZM("title"))
tempListStr = replace(tempListStr,"{电子邮件}",RSZM("emaill"))
tempListStr = replace(tempListStr,"{联系电话}",RSZM("phone"))
tempListStr = replace(tempListStr,"{联系QQ}",RSZM("qq"))
tempListStr = replace(tempListStr,"{留言详细地址}",hqlydz)
if DateDiff("d",""&rszm("sj")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{留言时间}","<font color="&sjys&">"&FormatTime(RSZM("sj"),D)&"</font>")
else
tempListStr = replace(tempListStr,"{留言时间}",FormatTime(RSZM("sj"),D))
end if
if DateDiff("d",""&rszm("sj")&"",Date())<=zmsj then
tempListStr = replace(tempListStr,"{留言时间1}","<font color="&sjys&">"&FormatTime(RSZM("sj"),D1)&"</font>")
else
tempListStr = replace(tempListStr,"{留言时间1}",FormatTime(RSZM("sj"),D1))
end if

 if rszm("qqh")=1 then
    set rs10 = server.createobject("adodb.recordset")
    sql10="select * from Members where ID="&rszm("MemID")&""
    rs10.open sql10,conn,1,1
    if rs10("MemName")=session("MemName") and session("MemLogin")="Succeed" then
	  MessageContent=rszm("Content")
	else
	  MessageContent=""&word(3)&""
	end if
    rs10.close
    set rs10=nothing
  else
    MessageContent=rszm("Content")
  end if
tempListStr = replace(tempListStr,"{留言内容}",MessageContent)

if rszm("qqh")=1 then

set rs11 = server.createobject("adodb.recordset")
sql11="select * from Members where ID="&rszm("MemID")&" and Lang='"&Language&"'"
rs11.open sql11,conn,1,1
if rs11("MemName")=session("MemName") and session("MemLogin")="Succeed" then
if IsNull(rszm("hf")) or trim(rszm("hf")&"")="" then
replayContent=""&word(4)&""
else
replayContent=rszm("hf")
end if
else
replayContent=""&word(3)&""
end if
rs11.close
set rs11=nothing
else
   
if trim(rszm("hf")&"")="" then
replayContent=""&word(4)&""
else
replayContent=rszm("hf")
end if
end if

tempListStr = replace(tempListStr,"{回复内容}",replayContent)
if trim(rszm("ReplyTime")&"")="" then
replaytime=""&word(4)&""
else
replaytime=rszm("ReplyTime")
end if
tempListStr = replace(tempListStr,"{回复时间}",replaytime)

returnStr = returnStr & tempListStr
RSZM.movenext
next
rszm.Close
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)
if instr(returnStr,"{分页}") > 0 then
returnStr = replace(returnStr,"{分页}","<div id=""pageguest"" class=""page""></div>")

end if 

returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pageguest',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(475)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(476)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(477)&","& vbCrLf
returnStr=returnStr&  "next: "&word(478)&","& vbCrLf
if word(479)<>"" then
returnStr=returnStr&  "last: "&word(479)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(480)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(481)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf



memberguest = returnStr
else
memberguest = word(0)
end if
end function
function memberyp(style,p)
if not (session("MemName")<>"" and session("MemLogin")="Succeed") then 
writemsg1 2,""&word(95)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","wrong"
response.end
end if
Set rs1 = server.CreateObject("adodb.recordset")
sql1="select * from members where MemName='"&session("MemName")&"' and Lang='"&Language&"'"
rs1.Open sql1, conn, 1, 1
id2=rs1("id")

set RSZM=server.createobject("adodb.recordset")
zm = "select * from HrDemandAccept where MemID="&id2&" and Lang='"&Language&"' order by "&p&"  "
'--------------调用样式开始
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""&word(649)&""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
'--------------调用样式结束
RSZM.open ZM,conn,1,1
If not( rszm.bof and rszm.EOF) Then
rszm.PageSize = hyypinfo
iCount = rszm.RecordCount 
iPageSize = rszm.PageSize
maxpage = rszm.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rszm.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If 		
For i = 1 To x
'标签{FH:2:twostyle}的用法
linenums = linenums + 1
tempListStr = listStr
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if
'标签{FH:2:twostyle}的用法结束

	tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{ID}",RSZM("id"))
tempListStr = replace(tempListStr,"{应聘职位}",RSZM("Quarters"))
tempListStr = replace(tempListStr,"{姓名}",RSZM("Name"))
tempListStr = replace(tempListStr,"{性别}",RSZM("Sex"))
tempListStr = replace(tempListStr,"{出生日期}",RSZM("Birthday"))
tempListStr = replace(tempListStr,"{身高}",RSZM("Stature"))
tempListStr = replace(tempListStr,"{户籍地}",RSZM("Residence"))
tempListStr = replace(tempListStr,"{婚姻状况}",RSZM("Marry"))
tempListStr = replace(tempListStr,"{毕业院校}",RSZM("School"))
tempListStr = replace(tempListStr,"{学历}",RSZM("Studydegree"))
tempListStr = replace(tempListStr,"{专业}",RSZM("Specialty"))
tempListStr = replace(tempListStr,"{毕业时间}",RSZM("Gradyear"))
tempListStr = replace(tempListStr,"{emaill}",RSZM("Email"))
tempListStr = replace(tempListStr,"{联系地址}",RSZM("Add"))
tempListStr = replace(tempListStr,"{邮编}",RSZM("Postcode"))
tempListStr = replace(tempListStr,"{教育经历}",RSZM("Edulevel"))
tempListStr = replace(tempListStr,"{个人简历}",RSZM("Experience"))
tempListStr = replace(tempListStr,"{应聘日期}",RSZM("Adddate"))
tempListStr = replace(tempListStr,"{联系电话}",RSZM("phone"))
if request.QueryString("m")<>"" then
tempListStr = replace(tempListStr,"{详细地址}",""&sitepath&"member/memberjobxx/?ypid="&rszm("id")&"&m="&request.QueryString("m")&"&Language="&Language&"")
else
tempListStr = replace(tempListStr,"{详细地址}",""&sitepath&"member/memberjobxx/?ypid="&rszm("id")&"&Language="&Language&"")
end if

if rszm("ReplyContent")<>"" then
tempListStr = replace(tempListStr,"{回复状态}",""&word(211)&"")
	else
tempListStr = replace(tempListStr,"{回复状态}",""&word(212)&"")
end if

if rszm("ReplyContent")<>"" then
tempListStr = replace(tempListStr,"{回复内容}",RSZM("ReplyContent"))
	else
tempListStr = replace(tempListStr,"{回复内容}",""&word(495)&"")
end if


returnStr = returnStr & tempListStr
RSZM.movenext
next
rszm.Close
returnStr = returnStr & endStr
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = replaceGlobal(returnStr)
returnStr = unrepreg(returnStr)
if instr(returnStr,"{分页}") > 0 then
returnStr = replace(returnStr,"{分页}","<div id=""pageyp"" class=""page""></div>")

end if 

returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pageyp',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(475)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(476)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(477)&","& vbCrLf
returnStr=returnStr&  "next: "&word(478)&","& vbCrLf
if word(479)<>"" then
returnStr=returnStr&  "last: "&word(479)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(480)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(481)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
memberyp = returnStr
else
memberyp = word(0)
end if
end function


'会员中心我的订单列表页开始

function membershowdd(style,msql,w,D,D1,p)
'是否是会员
if not (session("MemName")<>"" and session("MemLogin")="Succeed") then 
writemsg1 2,""&word(95)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","wrong"
end if

call initCodecs 
orderddh=request.QueryString("orderddh")
orderdate=request.QueryString("isdate")
if request.QueryString("action")="gb" then
sql2 = "update orderdd set isdate=5 where orderddh='"&orderddh&"'"
set rs2 = conn.Execute(sql2)
response.Redirect(""&sitepath&"member/memberorder/?Language="&Language&"&m="&request.QueryString("m")&"")
end if

Set rs1 = server.CreateObject("adodb.recordset")
sql1="select * from members where MemName='"&session("MemName")&"' and Lang='"&Language&"'"
rs1.Open sql1, conn, 1, 1
id2=rs1("id")


Set rs = server.CreateObject("adodb.recordset")
if orderdate="" then
sql = "select * from orderdd where MemID="&id2&" and Lang='"&Language&"'  "

if msql<>"no" then
sql=sql&" and "&msql&""
end if
sql=sql&" order by "&p&""


end if
if orderdate=1 then
sql = "select * from orderdd where MemID="&id2&" and Lang='"&Language&"' and isdate=0  "
if msql<>"no" then
sql=sql&" and "&msql&""
end if
sql=sql&" order by "&p&""
end if
if orderdate=2 then
sql = "select * from orderdd where MemID="&id2&" and Lang='"&Language&"' and isdate=1  "
if msql<>"no" then
sql=sql&" and "&msql&""
end if
sql=sql&" order by "&p&""
end if
if orderdate=3 then

sql = "select * from orderdd where MemID="&id2&" and isdate=2 and Lang='"&Language&"'   "
if msql<>"no" then
sql=sql&" and "&msql&""
end if
sql=sql&" order by "&p&""
end if
if orderdate=4 then
sql = "select * from orderdd where MemID="&id2&" and isdate=3 and Lang='"&Language&"'   "
if msql<>"no" then
sql=sql&" and "&msql&""
end if
sql=sql&" order by "&p&""
end if
if orderdate=5 then
sql = "select * from orderdd where MemID="&id2&" and isdate=5 and Lang='"&Language&"'   "
if msql<>"no" then
sql=sql&" and "&msql&""
end if
sql=sql&" order by "&p&""
end if
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
rs.Open sql, conn, 1, 1
If not( rs.bof and rs.EOF) Then
rs.PageSize = hyorderinfo
iCount = rs.RecordCount 
iPageSize = rs.PageSize
maxpage = rs.PageCount
page = request("page")
If Not IsNumeric(page) Or page = "" Then
page = 1
Else
page = CInt(page)
End If
If page<1 Then
page = 1
ElseIf page>maxpage Then
page = maxpage
End If
rs.AbsolutePage = Page
If page = maxpage Then
 x = iCount - (maxpage -1) * iPageSize
Else
 x = iPageSize
End If

For i = 1 To x

linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if



 if rs("isdate")=0 then
			  zt=""&word(239)&"" 
			  end if
			  if rs("isdate")=1 then  
			  zt=""&word(240)&"" 
			  end if
			  if rs("isdate")=2 then  
			  zt=""&word(241)&"" 
			  end if
			  if rs("isdate")=3 then  
			  zt=""&word(242)&"" 
			  end if
			  if rs("isdate")=5 then  
			  zt=""&word(243)&"" 
			  end if	
			
tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{所属语言}",rs("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{性别}",rs("Sex"))
tempListStr = replace(tempListStr,"{单位}",rs("Company"))
tempListStr = replace(tempListStr,"{地址}",rs("Address"))
tempListStr = replace(tempListStr,"{邮编}",rs("ZipCode"))
tempListStr = replace(tempListStr,"{传真}",rs("Fax"))
tempListStr = replace(tempListStr,"{邮箱}",rs("Email"))
tempListStr = replace(tempListStr,"{订单号}",rs("orderddh"))
tempListStr = replace(tempListStr,"{订购人}",rs("Linkman"))
tempListStr = replace(tempListStr,"{联系电话}",rs("Linkman"))
tempListStr = replace(tempListStr,"{订购时间}",FormatTime(rs("Addtime"),D))
tempListStr = replace(tempListStr,"{订购时间1}",FormatTime(rs("Addtime"),D1))
tempListStr = replace(tempListStr,"{订购状态}",zt)


if rs("zffs")=1 then 
zs=""&word(236)&""
end if
if rs("zffs")=2 then 
zs=""&word(237)&""
end if
tempListStr = replace(tempListStr,"{支付方式}",zs)
if rs("ReplyTime")<>"" then	
tempListStr = replace(tempListStr,"{回复时间}",rs("ReplyTime"))	
else
tempListStr = replace(tempListStr,"{回复时间}",""&word(234)&"")	
end if

if rs("ReplyContent")<>"" then	
tempListStr = replace(tempListStr,"{回复内容}",rs("ReplyContent"))
else
tempListStr = replace(tempListStr,"{回复内容}",""&word(234)&"")	
end if
tempListStr = replace(tempListStr,"{配送方式}",rs("shipmentName"))
tempListStr = replace(tempListStr,"{备注}",rs("Remark"))
tempListStr = replace(tempListStr,"{配送费用}",formatNumber2(rs("shipmentMoney")))
tempListStr = replace(tempListStr,"{产品费用}",formatNumber2(rs("aumont")))
tempListStr = replace(tempListStr,"{支付总额}",formatNumber2(rs("shipmentMoney")+rs("aumont")))
tempListStr = replace(tempListStr,"{支付链接}",""&template&"member/pay.asp?orderddh="&base64Encode(rs("orderddh"))&"&m="&m1&"&language="&language&"")	

if rs("isdate")=0 then
tempListStr = replace(tempListStr,"{支付}",""&word(249)&"")
else
tempListStr = replace(tempListStr,"{支付}","")
end if

if rs("isdate")=3 then
shouhuo="<a href=""?orderddh="&rs("orderddh")&"&action=gb&language="&language&"&m="&request.QueryString("m")&""">"&word(250)&"</a>"
end if
if rs("isdate")=5 then
shouhuo=""&word(251)&""
end if

tempListStr = replace(tempListStr,"{是否收货}",shouhuo)

qxurl="?ddh="&rs("orderddh")&"&action=Del&language="&language&"&m="&request.QueryString("m")&"" 

tempListStr = replace(tempListStr,"{取消订单url}",qxurl)

view=""&sitepath&"member/memberorderview/?orderid="&rs("id")&"&m="&request.QueryString("m")&"&language="&language&"&orderdd="&rs("orderddh")&""

tempListStr = replace(tempListStr,"{查看详细}",view)



returnStr = returnStr & tempListStr	
   		
rs.movenext
Next
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)



'分页改动采用js分页！更新于2015年05.07
'更新人:追梦工作室
if instr(returnStr,"{分页}") > 0 then
if w=1 then
returnStr = replace(returnStr,"{分页}","<div id=""pageorderpc"" class=""page""></div>")
else
returnStr = replace(returnStr,"{分页}","<div id=""pageorderwap"" class=""page""></div>")
end if
end if 



if w=1 then
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pageorderpc',"& vbCrLf
returnStr=returnStr&  "pages: "&maxpage&","& vbCrLf
returnStr=returnStr&  "skin: '"&word(475)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(476)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(477)&","& vbCrLf
returnStr=returnStr&  "next: "&word(478)&","& vbCrLf
if word(479)<>"" then
returnStr=returnStr&  "last: "&word(479)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(480)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(481)&","& vbCrLf
returnStr=returnStr&"curr: "&page&","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf

else
returnStr=returnStr&"<script type=""text/javascript"">"& vbCrLf  
returnStr=returnStr&"laypage({"& vbCrLf
returnStr=returnStr&  "cont: 'pageorderwap',"& vbCrLf
returnStr=returnStr&  "pages: """&maxpage&""","& vbCrLf
returnStr=returnStr&  "skin: '"&word(482)&"',"& vbCrLf
returnStr=returnStr&  "first: "&word(483)&","& vbCrLf
returnStr=returnStr&  "prev: "&word(484)&","& vbCrLf
returnStr=returnStr&  "next: "&word(485)&","& vbCrLf
if word(486)<>"" then
returnStr=returnStr&  "last: "&word(486)&","& vbCrLf
end if
returnStr=returnStr&  "groups: "&word(496)&","& vbCrLf
returnStr=returnStr&  "skip: "&word(497)&","& vbCrLf
returnStr=returnStr&"curr: """&page&""","& vbCrLf
returnStr=returnStr&"jump: function(e, first){"& vbCrLf
returnStr=returnStr&"if(!first){"& vbCrLf
returnStr=returnStr&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
returnStr=returnStr&"}"& vbCrLf
returnStr=returnStr&" }"& vbCrLf
returnStr=returnStr&"});"& vbCrLf
returnStr=returnStr&"</script>"& vbCrLf
end if
membershowdd = returnStr
else
membershowdd = word(253)

end if
rs.close
set rs=nothing


if request("action")="Del" then	
ddh=request.QueryString("ddh")
sql8="Delete from orderdd where orderddh='"&ddh&"'"
sql9="Delete from ordercp where orderddh='"&ddh&"'"
conn.execute(sql8)
conn.execute(sql9)
response.Write "<script language=javascript>window.location.href="""&sitepath&"member/Memberorder/?m="&request.QueryString("m")&"&Language="&Language&"""</script>"
end if

end function



call initCodecs 
orderid=request.QueryString("orderid")
orderdd=request.QueryString("orderdd")

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from orderdd where ID="&orderid&"  "
rs.Open sql, conn, 1, 1
isdate1=rs("isdate")
rs.close
set rs=nothing

function memberorderview(style,msql,D,D1,p)
set rs1=server.createobject("adodb.recordset")
sql1="select * from ordercp where orderddh='"&orderdd&"' and Lang='"&Language&"' "
if msql<>"no" then
sql1=sql1&" and "&msql&""
end if
sql1=sql1&" order by "&p&""
tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
rs1.open sql1,conn,1,1
totalprice=0
totalcount=0
For i = 1 To rs1.RecordCount
linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rs1("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
if rs1("productslt")<>"" then
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&""&rs1("productslt")&"")
else
tempListStr = replace(tempListStr,"{产品缩略图}",""&sitepath&"images/demo.jpg")
end if
tempListStr = replace(tempListStr,"{订单号}",""&rs1("orderddh")&"")
tempListStr = replace(tempListStr,"{产品名称}",""&rs1("productname")&"")
set rs3=server.CreateObject("adodb.recordset")
sql3="select * from product_stock where id="&rs1("p_id")&" and lang='"&language&"'"
rs3.open sql3,conn,1,1

if not rs3.eof then
guige=rs3("s_value")
kucun=rs3("p_stock")
else
guige=""&word(650)&""
kucun=""&word(651)&""
end if

rs3.close
set rs3=nothing
tempListStr = replace(tempListStr,"{产品规格}",guige)
tempListStr = replace(tempListStr,"{产品库存}",kucun)

tempListStr = replace(tempListStr,"{购买价}",formatNumber2(rs1("hyj")))
tempListStr = replace(tempListStr,"{购买数量}",rs1("bookcount"))
tempListStr = replace(tempListStr,"{商品价格}",formatNumber2(rs1("hyj")*rs1("bookcount")))

tempListStr = replace(tempListStr,"{订购时间}",FormatTime(RS1("ordersj"),D))
tempListStr = replace(tempListStr,"{订购时间}",FormatTime(RS1("ordersj"),D1))
totalprice=formatNumber2(totalprice+rs1("hyj")*rs1("bookcount"))
totalcount=totalcount+rs1("bookcount")

returnStr = returnStr & tempListStr
rs1.movenext
next

returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
returnStr = replace(returnStr,"{地址栏语言}",language)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)

returnStr = replace(returnStr,"{商品总计}",totalcount)
returnStr = replace(returnStr,"{商品总价}",totalprice)

memberorderview = returnStr



end function
'会员中心 我的订单详细页面
Set rsorder = server.CreateObject("adodb.recordset")
sqlorder = "select * from orderdd where ID="&orderid&"  "
rsorder.Open sqlorder, conn, 1, 1
 '订购时间
orderdgsj=rsorder("addtime")
'支付方式
if rsorder("zffs")=1 then
orderzffs=""&word(269)&""
end if
if rsorder("zffs")=2 then
orderzffs=""&word(270)&""
end if
'配送方式
orderpsfs=rsorder("shipmentName")
'配送费用
orderpsfy=formatNumber2(rsorder("shipmentmoney"))
'支付总额"
orderzfze=formatNumber2(rsorder("shipmentmoney")+rsorder("aumont"))
'支付链接
orderzflj=""&template&"member/pay.asp?orderddh="&base64Encode(rsorder("orderddh"))&"&m="&m1&"&language="&language&""
'订单状态"

if rsorder("isdate")=0 then
orderddzt=""&word(278)&" " 
end if
if rsorder("isdate")=1 then  
orderddzt=""&word(279)&"" 
end if
if rsorder("isdate")=2 then  
orderddzt=""&word(280)&"" 
end if
if rsorder("isdate")=3 then  
orderddzt=""&word(281)&"" 
end if
if rsorder("isdate")=5 then  
orderddzt=""&word(282)&"" 
end if
'订单数字状态
orderszzt=rsorder("isdate")
'处理时间"
orderclsj=rsorder("ReplyTime")	
'回复内容"
orderhfnr=rsorder("ReplyContent")	
'订购用户"
orderdgyh=rsorder("Linkman")
'用户性别"
orderyhxb=rsorder("sex")
'用户单位"
orderyhdw=rsorder("company")

'用户地址"
orderyhdz=rsorder("address")

'用户邮编"
orderyhyb=rsorder("ZipCode")
'电话号码"
orderdhhm=rsorder("Mobile")
'传真号码"
orderczhm=rsorder("fax")
'电子邮箱"
orderdzyx=rsorder("email")
'订单备注"
orderddbz=HtmlStrReplace(rsorder("Remark"))

rsorder.close
set rsorder=nothing

'会员中心我的应聘详细内容
ypid=request.QueryString("ypid")
Set rsyp= server.CreateObject("adodb.recordset")
sqlyp = "select * from HrDemandAccept where ID="&ypid&"  "
rsyp.Open sqlyp, conn, 1, 1


'"应聘职位"
ypzw=rsyp("Quarters")

'"姓名"
ypxm=rsyp("name")
'"性别"
ypxb=rsyp("sex")
'"身高"
ypsg=rsyp("Stature")
'"户籍"
yphj=rsyp("Residence")
'"联系电话"
yplxdh=rsyp("phone")
'"毕业院校"
ypbyyx=rsyp("School")
'"学历"
ypxl=rsyp("Studydegree")

'"专业"
ypzy=rsyp("Specialty")

'"毕业时间"
ypbysj=rsyp("Gradyear")

'"教育经历"
ypjyjl=rsyp("Edulevel")
'"个人简历"
ypgrjl=rsyp("Experience")
'"应聘日期"
ypyprq=rsyp("Adddate")

'"出生日期"
ypcsrq=rsyp("BirthDay")
'"婚姻状况"
yphyzk=rsyp("marry")
'"联系地址"
yplxdz=rsyp("add")

'"邮政编码"
ypyzbm=rsyp("postcode")

'"电子邮件"
ypdzyj=rsyp("Email")


'"回复时间"
yphfsj=rsyp("ReplyTime")
'"回复内容"
yphfnr=rsyp("ReplyContent")


rsyp.close
set rsyp=nothing

function province(style,N,msql,p)
set RSZM=server.createobject("adodb.recordset")
ZM="select top "&N&" * from province"
if msql<>"no" then
zm=zm&" and "&msql&""
end if
zm=zm&" order by "&p&""


tempTagStr=getTagStr(style)
if tempTagStr = "" or isnull(tempTagStr) then
tempTagStr = ""
end if
if instr(tempTagStr,"[LOOP]") > 0 and instr(tempTagStr,"[/LOOP]") then
startStr = mid(tempTagStr,1,instr(tempTagStr,"[LOOP]")-1)
listStr = mid(tempTagStr,instr(tempTagStr,"[LOOP]")+6,instr(tempTagStr,"[/LOOP]") - instr(tempTagStr,"[LOOP]") - 6)
endStr = mid(tempTagStr,instr(tempTagStr,"[/LOOP]")+7)
else
startStr = ""
endStr = ""
listStr = tempTagStr
end if
returnStr = startStr
RSZM.open ZM,conn,1,1	
 	
For i = 1 To RSZM.RecordCount


	
linenums = linenums + 1
tempListStr = listStr
'{FH:2:twostyle}
if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
else
'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
end if
end if

lqy=request.QueryString("key")	      
set RSZM1=server.createobject("adodb.recordset")
ZM1="select * from province where province='"&lqy&"' "		 
RSZM1.open ZM1,conn,1,1	
lid=RSZM1("id")
RSZM1.close 
Set RSZM1=Nothing 
MyComp = StrComp(lqy,rszm("province"),1)  

tempListStr = replace(tempListStr,"{i}",linenums)
tempListStr = replace(tempListStr,"{ii}",right("0" & linenums, 2))
tempListStr = replace(tempListStr,"{所属语言}",rszm("lang"))
tempListStr = replace(tempListStr,"{地址语言}",language)
tempListStr = replace(tempListStr,"{网站路径}",""&sitepath&"")
tempListStr = replace(tempListStr,"{模板路径}",""&template&"")
tempListStr = replace(tempListStr,"{省份ID}",RSZM("ID"))
tempListStr = replace(tempListStr,"{地址栏ID}",lid)
tempListStr = replace(tempListStr,"{省份名称}",RSZM("Province"))
tempListStr = replace(tempListStr,"{条件对比}",MyComp)
tempListStr = replace(tempListStr,"{链接地址}",""&sitepath&""&yyxwl&"/?key="&rszm("province")&"&m="&rszm("m")&"")

returnStr = returnStr & tempListStr
RSZM.movenext
next
RSZM.close 
set RSZM=nothing
returnStr = returnStr & endStr
returnStr = replaceGlobal(returnStr)
Set re = New RegExp
re.ignoreCase = True
re.Global = True 
re.Pattern = "{if:(.+?)}([\s\S]*?){else}([\s\S]*?){end if}"
set tmatch = re.Execute(returnStr )
For Each m in tmatch 
returnStr = replace(returnStr , m.Value, Replace_tpl_If(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2))) 
Next
returnStr = unrepreg(returnStr)
province = returnStr
end function
%>