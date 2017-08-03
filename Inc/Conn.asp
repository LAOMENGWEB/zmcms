<%
dim conn,db
dim connstr
Dim dataname,sitepath
%>
<!--#include file="sqlhost.asp"-->
<%
sitepath="/"
dataname="zmcmszhuimeng.mdb"
admindb="../"
db=admindb&"Databases/"&DataName
SiteDataPath = SitePath&"Databases/"&DataName   
SiteDataBakPath = SitePath&"Databases/bak_"&DataName 

	
	
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	if sqlno=1 then
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(SiteDataPath)
	end if
	if sqlno=2 then
	connstr="Provider=Sqloledb;User ID="&datauser&";Password="&datapass&";Initial CataLog="&datasqlname&";Data Source="&datahost&";"
	end if
	conn.Open connstr
	If Err Then
	err.Clear
	Set Conn = Nothing
	Response.End
	End If




'20150525 增加前台默认语言
set rslang=server.CreateObject("adodb.recordset")
rslang.open "select * from lang where qt=1 ",conn,1,1
langqt=rslang("langbs")
langid=rslang("id")
rslang.close
Set rslang=Nothing




If request.QueryString("Language")<>"" Then
Language=request.QueryString("Language")
Else
Language=langqt
End if

'获取language语言，防止获取不到的页面使用
set rs5=server.CreateObject("adodb.recordset")
rs5.open "select * from lanyy where id=1 ",conn,1,1
lanyybs=rs5("lanyybs")
rs5.close
set rs5=nothing
sub zmqylanyy
if request.QueryString("language")<>"" then
if lanyybs <> language  then
zmsql="update lanyy set lanyybs='"&language&"' where id =1"
conn.Execute ( zmsql )
end if
end if
end sub

'20150525 增加前台默认语言
set rslang=server.CreateObject("adodb.recordset")
rslang.open "select * from lang where langbs='"&language&"' ",conn,1,1

langnameqt=rslang("langname")
hbfh=rslang("huobi")
hbwz=rslang("huobi1")
rslang.close
Set rslang=Nothing

If language="ch" then
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
Newclass="list"&language&""        '新闻分类前缀
newname="New"&language&""        '新闻详细页前缀
ProductClass="Productlist"&language&""        '产品分类前缀
ProductName="Product"&language&""        '产品详细页前缀
alClass="photolist"&language&""        '案例分类前缀
alName="photo"&language&""        '案例详细页前缀
downClass="downlist"&language&""        '下载分类前缀
downName="down"&language&""        '下载详细页前缀
tvClass="videolist"&language&""        '视频分类前缀
tvName="video"&language&""        '视频详细页前缀
zpClass="joblist"&language&""        '招聘列表前缀
zpName="job"&language&""        '招聘详细页前缀
aboutname="about"&language&""        '企业详细页前缀
guestname="showguest"&language&""        '显示留言页前缀
Separated="_"        '分隔符

End if




set rswzset=server.CreateObject("adodb.recordset")
rswzset.open "select * from wzset where lang='"&language&"'",conn,1,1
if rswzset.bof and rswzset.eof then response.end
wzkg=rswzset("wzkg")
wh=rswzset("wh")
yy=rswzset("yy")
kq=rswzset("kq")
mc=rswzset("mc")
tjdm=rswzset("tjdm")

tvid=rswzset("tvid")
dcid=rswzset("dcid")
qyid=rswzset("qyid")
zid=rswzset("zid")
ljid=rswzset("ljid")
m1=rswzset("m1")
m2=rswzset("m2")
m3=rswzset("m3")
m4=rswzset("m4")
m6=rswzset("m6")

ewm=rswzset("ewm")
ewmwidth=rswzset("ewmwidth")
ewmheight=rswzset("ewmheight")

wxewm=rswzset("wxewm")
wxewmwidth=rswzset("wxewmwidth")
wxewmheight=rswzset("wxewmheight")


wapmap=rswzset("wapmap")
wapmap1=Split(wapmap,",")
for i=0 to ubound(wapmap1)
wapx=wapmap1(0)
wapy=wapmap1(1)
next
wapmc=rswzset("wapmc")
wapcontent=rswzset("wapcontent")
mapphone=rswzset("mapphone")
ms=rswzset("ms")
gjc=rswzset("gjc")
ym=rswzset("ym")
yx=rswzset("yx")
dw=rswzset("dw")
lxdz=rswzset("lxdz")
lxr=rswzset("lxr")
gddh=rswzset("gddh")
sj=rswzset("sj")
ba=rswzset("ba")
cz=rswzset("cz")
banquan=rswzset("bq")
jt=rswzset("jt")
kf=rswzset("kf")
tc=rswzset("tc")
qq=rswzset("qq")
pfad=rswzset("pfad")
dlad=rswzset("dlad")
ly=rswzset("ly")
killword=rswzset("killword")
notzc=rswzset("notzc")
sjys=rswzset("sjys")
zmsj=rswzset("zmsj")
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
hdsl=rswzset("hdsl")
piclx=rswzset("piclx")
picw=rswzset("PicW")
pich=rswzset("Pich")
hyzc=rswzset("hyzc")
hysh=rswzset("hysh")
hylyinfo=rswzset("hylyinfo")
hyypinfo=rswzset("hyypinfo")
yxwlinfo=rswzset("yxwlinfo")
hyorderinfo=rswzset("hyorderinfo")
ykgw=rswzset("ykgw")
maillock=rswzset("maillock") 
MailServer=rswzset("MailServer")
MailAdd=rswzset("MailAdd")
MailPassword=rswzset("MailPassword")
MailUserName=rswzset("MailUserName")
MailUserNameen=rswzset("MailUserNameen")
PostMail=rswzset("PostMail")
sendmail=rswzset("sendmail")
newguest=rswzset("newguest")
newyp=rswzset("newyp")
newdd=rswzset("newdd")
hfguest=rswzset("hfguest")
hfyp=rswzset("hfyp")
getpass1=rswzset("getpass")
zxyp=rswzset("zxyp")
hfdd=rswzset("hfdd")
zchy=rswzset("zchy")
map=rswzset("map")
JMailDisplay=rswzset("JMailDisplay")
zxys=rswzset("zxys")
sodd=rswzset("sodd")
alipayemaill=rswzset("alipayemaill")
alipayapi=rswzset("alipayapi")
alipaypid=rswzset("alipaypid")
alipaykey=rswzset("alipaykey")
blank=rswzset("blank")
skins=rswzset("skins")
mde=rswzset("mde")
minfo=rswzset("minfo")
mpro=rswzset("mpro")
mph=rswzset("mph")
mtv=rswzset("mtv")
mdo=rswzset("mdo")
mzp=rswzset("mzp")
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
template=""&sitepath&"template/"&skins&"/"



set rswzset=server.CreateObject("adodb.recordset")
rswzset.open "select * from logo where lang='"&language&"' and skins='"&skins&"'",conn,1,1
if not(rswzset.bof and rswzset.eof) then 

logo=rswzset("logourl")
logowidth=rswzset("logowidth")
logoheight=rswzset("logoheight")

end if
rswzset.close
set rswzset=nothing


set rswzset=server.CreateObject("adodb.recordset")
rswzset.open "select * from waplogo where lang='"&language&"' and skins='"&skins&"'",conn,1,1
if not(rswzset.bof and rswzset.eof) then 

waplogo=rswzset("waplogourl")
waplogowidth=rswzset("waplogowidth")
waplogoheight=rswzset("waplogoheight")
end if
rswzset.close
set rswzset=nothing


%>
<!--#include file="Config.asp"-->
<%

'********************************
'字符串加密类
'********************************
const BASE_64_MAP_INIT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/" 
dim newline 
dim Base64EncMap(63) 
dim Base64DecMap(127) 
'初始化函数 
PUBLIC SUB initCodecs() 
	' 初始化变量 
	newline = "<P>" & chr(13) & chr(10) 
	dim max, IDx 
	max = len(BASE_64_MAP_INIT) 
	for IDx = 0 to max - 1 
	Base64EncMap(IDx) = mID(BASE_64_MAP_INIT, IDx + 1, 1) 
	next 
	for IDx = 0 to max - 1 
	Base64DecMap(ASC(Base64EncMap(IDx))) = IDx 
	next 
END SUB

'Base64加密函数 
PUBLIC FUNCTION base64Encode(plain) 
	if len(plain) = 0 then 
	base64Encode = "" 
	exit function 
	end if 
	dim ret, ndx, by3, first, second, third 
	by3 = (len(plain) \ 3) * 3 
	ndx = 1 
	do while ndx <= by3 
	first = asc(mID(plain, ndx+0, 1)) 
	second = asc(mID(plain, ndx+1, 1)) 
	third = asc(mID(plain, ndx+2, 1)) 
	ret = ret & Base64EncMap( (first \ 4) AND 63 ) 
	ret = ret & Base64EncMap( ((first * 16) AND 48) + ((second \ 16) AND 15 ) ) 
	ret = ret & Base64EncMap( ((second * 4) AND 60) + ((third \ 64) AND 3 ) ) 
	ret = ret & Base64EncMap( third AND 63) 
	ndx = ndx + 3 
	loop 
	if by3 < len(plain) then 
	first = asc(mID(plain, ndx+0, 1)) 
	ret = ret & Base64EncMap( (first \ 4) AND 63 ) 
	if (len(plain) MOD 3 ) = 2 then 
	second = asc(mID(plain, ndx+1, 1)) 
	ret = ret & Base64EncMap( ((first * 16) AND 48) + ((second \ 16) AND 15 ) ) 
	ret = ret & Base64EncMap( ((second * 4) AND 60) ) 
	else 
	ret = ret & Base64EncMap( (first * 16) AND 48) 
	ret = ret '& "=" 
	end if 
	ret = ret '& "=" 
	end if 
	base64Encode = ret 
END FUNCTION 

'Base64解密函数 
PUBLIC FUNCTION base64Decode(scrambled) 
	if len(scrambled) = 0 then 
	base64Decode = "" 
	exit function 
	end if 
	dim realLen 
	realLen = len(scrambled) 
	do while mID(scrambled, realLen, 1) = "=" 
	realLen = realLen - 1 
	loop 
	dim ret, ndx, by4, first, second, third, fourth 
	ret = "" 
	by4 = (realLen \ 4) * 4 
	ndx = 1 
	do while ndx <= by4 
	first = Base64DecMap(asc(mID(scrambled, ndx+0, 1))) 
	second = Base64DecMap(asc(mID(scrambled, ndx+1, 1))) 
	third = Base64DecMap(asc(mID(scrambled, ndx+2, 1))) 
	fourth = Base64DecMap(asc(mID(scrambled, ndx+3, 1))) 
	ret = ret & chr( ((first * 4) AND 255) + ((second \ 16) AND 3)) 
	ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15)) 
	ret = ret & chr( ((third * 64) AND 255) + (fourth AND 63)) 
	ndx = ndx + 4 
	loop 
	if ndx < realLen then 
	first = Base64DecMap(asc(mID(scrambled, ndx+0, 1))) 
	second = Base64DecMap(asc(mID(scrambled, ndx+1, 1))) 
	ret = ret & chr( ((first * 4) AND 255) + ((second \ 16) AND 3)) 
	if realLen MOD 4 = 3 then 
	third = Base64DecMap(asc(mID(scrambled,ndx+2,1))) 
	ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15)) 
	end if 
	end if 
	base64Decode = ret 
END FUNCTION


'加载不同的语言包开始 更新于20150526
include ""&template&"lang/sys-"&Language&".inc"
include ""&template&"lang/temp-"&Language&".inc"
'加载不同的语言包结束

%>
