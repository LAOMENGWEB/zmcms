<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|1000002,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request.Form("submit") = "添加" Then
langname=Trim(request.form("langname"))
langbs=trim(request.form("langbs"))
huobi=trim(request.form("huobi"))
huobi1=trim(request.form("huobi1"))
px=trim(request.form("px"))

SmallImage=trim(request.form("SmallImage"))
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from lang"
	rs.open sql,conn,1,3

		rs.AddNew 
		rs("langname")=langname
		rs("langbs")=langbs	
	  rs("huobi")=huobi
	  rs("huobi1")=huobi
	  rs("ht")=0
	  rs("qt")=0
	  rs("xs")=1

	  rs("px")=px
	  rs("SmallImage")=SmallImage
		rs.update
		rs.close
	set rs=nothing
	call addlog("添加多语言")


'获取最新的语言标识
set rs=conn.execute("select top 1 langbs from [lang] order by id desc")
langbs1=rs("langbs")
rs.close:set rs=Nothing
'获取系统的基本设置
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from wzset where id=1",conn,1,1
wzkg=rs("wzkg")
yy=rs("yy")
jt=rs("jt")
wh=rs("wh")
mc=rs("mc")
tjdm=rs("tjdm")
logo=rs("logo")
logowidth=rs("logowidth")
logoheight=rs("logoheight")
ym=rs("ym")
yx=rs("yx")
dw=rs("dw")
lxdz=rs("lxdz")
lxr=rs("lxr")
gddh=rs("gddh")
sj=rs("sj")
ba=rs("ba")
cz=rs("cz")
ms=rs("ms")
gjc=rs("gjc")
killword=rs("killword")
sjys=rs("sjys")
zmsj=rs("zmsj")
newinfo=rs("newinfo")
ProductInfo=rs("ProductInfo")
alInfo=rs("alInfo")
downinfo=rs("downinfo")
tvinfo=rs("tvinfo")
zpinfo=rs("zpinfo")
lyinfo=rs("lyinfo")
yxwlinfo=rs("yxwlinfo")
sodd=rs("sodd")
ly=rs("ly")
kf=rs("kf")
tc=rs("tc")
pfad=rs("pfad")
dlad=rs("dlad")
hdsl=rs("hdsl")
PicW=rs("PicW")
PicH=rs("PicH")
Piclx=rs("Piclx")
PicBW=rs("PicBW")
PicBH=rs("PicBH")
Pictxt=rs("Pictxt")
PicSize=rs("PicSize")
PicColor=rs("PicColor")
PicLocation=rs("PicLocation")
hyzc=rs("hyzc")
hysh=rs("hysh")
zxyp=rs("zxyp")
getpass=rs("getpass")
hylyinfo=rs("hylyinfo")
hyypinfo=rs("hyypinfo")
hyorderinfo=rs("hyorderinfo")
maillock=rs("maillock")
MailServer=rs("MailServer")
MailAdd=rs("MailAdd")
MailPassword=rs("MailPassword")
MailUserName=rs("MailUserName")
ykgw=rs("ykgw")
MailUserNameen=rs("MailUserNameen")
PostMail=rs("PostMail")
sendmail=rs("sendmail")
newguest=rs("newguest")
newyp=rs("newyp")
newdd=rs("newdd")
hfguest=rs("hfguest")
hfyp=rs("hfyp")
hfdd=rs("hfdd")
zchy=rs("zchy")
sytp=rs("sytp")
notzc=rs("notzc")
map=rs("map")
JMailDisplay=rs("JMailDisplay")
zxys=rs("zxys")
bq=rs("bq")
qq=rs("qq")
tvid=rs("tvid")
dcid=rs("dcid")
alipayemaill=rs("alipayemaill")
alipaypid=rs("alipaypid")
alipaykey=rs("alipaykey")
alipayapi=rs("alipayapi")
blank=rs("blank")
waplogo=rs("waplogo")
waplogowidth=rs("waplogowidth")
waplogoheight=rs("waplogoheight")
wapnewinfo=rs("wapnewinfo")
wapProductInfo=rs("wapProductInfo")
wapalInfo=rs("wapalInfo")
wapdowninfo=rs("wapdowninfo")
waptvinfo=rs("waptvinfo")
wapzpinfo=rs("wapzpinfo")
waplyinfo=rs("waplyinfo")
wapmap=rs("wapmap")
wapmc=rs("wapmc")
wapcontent=rs("wapcontent")
ewm=rs("ewm")
ewmwidth=rs("ewmwidth")
ewmheight=rs("ewmheight")
wxewm=rs("wxewm")
wxewmwidth=rs("wxewmwidth")
wxewmheight=rs("wxewmheight")
mapphone=rs("mapphone")
kq=rs("kq")
skins=rs("skins")
mde=rs("mde")
minfo=rs("minfo")
mpro=rs("mpro")
mph=rs("mph")
mtv=rs("mtv")
mdo=rs("mdo")
mzp=rs("mzp")
tvid=rs("tvid")
dcid=rs("dcid")
qyid=rs("qyid")
m1=rs("m1")	
m2=rs("m2")
m3=rs("m3")
m4=rs("m4")
m6=rs("m6")
zid=rs("zid")
ljid=rs("ljid")
rs.close:set rs=Nothing


Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset  "
rs.Open sql, conn, 1, 3
rs.addnew

rs("wzkg")=wzkg
rs("yy")=yy
rs("jt")=jt
rs("wh")=wh
rs("mc")=mc
rs("tjdm")=tjdm
rs("logo")=logo
rs("logowidth")=logowidth
rs("logoheight")=logoheight
rs("ym")=ym
rs("yx")=yx
rs("dw")=dw
rs("lxdz")=lxdz
rs("lxr")=lxr
rs("gddh")=gddh
rs("sj")=sj
rs("ba")=ba
rs("cz")=cz
rs("ms")=ms
rs("gjc")=gjc
rs("killword")=killword
rs("sjys")=sjys
rs("zmsj")=zmsj
rs("newinfo")=newinfo
rs("ProductInfo")=ProductInfo
rs("alInfo")=alInfo
rs("downinfo")=downinfo
rs("tvinfo")=tvinfo
rs("zpinfo")=zpinfo
rs("lyinfo")=lyinfo
rs("yxwlinfo")=yxwlinfo
rs("sodd")=sodd
rs("tvid")=0
rs("dcid")=0
rs("ly")=ly
rs("kf")=kf
rs("tc")=tc
rs("pfad")=pfad
rs("dlad")=dlad
rs("hdsl")=hdsl
rs("PicW")=PicW
rs("PicH")=PicH
rs("Piclx")=Piclx
rs("PicBW")=PicBW
rs("PicBH")=PicBH
rs("Pictxt")=Pictxt
rs("PicSize")=PicSize
rs("PicColor")=PicColor
rs("PicLocation")=PicLocation
rs("hyzc")=hyzc
rs("hysh")=hysh
rs("zxyp")=zxyp
rs("getpass")=getpass
rs("hylyinfo")=hylyinfo
rs("hyypinfo")=hyypinfo
rs("hyorderinfo")=hyorderinfo
rs("maillock")=maillock
rs("MailServer")=MailServer
rs("MailAdd")=MailAdd
rs("MailPassword")=MailPassword
rs("MailUserName")=MailUserName
rs("ykgw")=ykgw
rs("MailUserNameen")=MailUserNameen
rs("PostMail")=PostMail
rs("sendmail")=sendmail
rs("newguest")=newguest
rs("newyp")=newyp
rs("newdd")=newdd
rs("hfguest")=hfguest
rs("hfyp")=hfyp
rs("hfdd")=hfdd
rs("zchy")=zchy
rs("sytp")=sytp
rs("notzc")=notzc
rs("map")=map
rs("JMailDisplay")=JMailDisplay
rs("zxys")=zxys
rs("bq")=bq
rs("qq")=qq
rs("alipayemaill")=alipayemaill
rs("alipaypid")=alipaypid
rs("alipaykey")=alipaykey
rs("alipayapi")=alipayapi
rs("blank")=blank
rs("waplogo")=waplogo
rs("waplogowidth")=waplogowidth
rs("waplogoheight")=waplogoheight
rs("wapnewinfo")=wapnewinfo
rs("wapProductInfo")=wapProductInfo
rs("wapalInfo")=wapalInfo
rs("wapdowninfo")=wapdowninfo
rs("waptvinfo")=waptvinfo
rs("wapzpinfo")=wapzpinfo
rs("waplyinfo")=waplyinfo
rs("wapmap")=wapmap
rs("wapmc")=wapmc
rs("wapcontent")=wapcontent
rs("ewm")=ewm
rs("ewmwidth")=ewmwidth
rs("ewmheight")=ewmheight
rs("wxewm")=wxewm
rs("wxewmwidth")=wxewmwidth
rs("wxewmheight")=wxewmheight
rs("mapphone")=mapphone
rs("kq")=kq
rs("skins")=skins
rs("mde")=mde
rs("minfo")=minfo
rs("mpro")=mpro
rs("mph")=mph
rs("mtv")=mtv
rs("mdo")=mdo
rs("mzp")=mzp
rs("tvid")=tvid
rs("dcid")=dcid
rs("qyid")=qyid
rs("m1")=m1	
rs("m2")=m2
rs("m3")=m3
rs("m4")=m4
rs("m6")=m6
rs("zid")=zid
rs("ljid")=ljid
rs("lang")=langbs1

rs.update
    rs.Close
    Set rs = Nothing


set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ad where id=1",conn,1,1
width=rs("width")
height=rs("height")
pfpic=rs("pfpic")
pfurl=rs("pfurl")
pfwidth=rs("pfwidth")
pfheight=rs("pfheight")
pic=rs("pic")
url=rs("url")
top=rs("top")
left1=rs("left1")
right1=rs("right1")
pic1=rs("pic1")
url1=rs("url1")
lf=rs("lf")
rf=rs("rf")
rs.close:set rs=Nothing

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ad  "
rs.Open sql, conn, 1, 3
rs.addnew

rs("width")=width
rs("height")=height
rs("pfpic")=pfpic
rs("pfurl")=pfurl
rs("pfwidth")=pfwidth
rs("pfheight")=pfheight
rs("pic")=pic
rs("url")=url
rs("top")=top
rs("left1")=left1
rs("right1")=right1
rs("pic1")=pic1
rs("url1")=url1
rs("lf")=lf
rs("rf")=rf


rs("lang")=langbs1
rs.update
    rs.Close
    Set rs = Nothing




set rs=server.CreateObject("adodb.recordset")
rs.open "select * from hygroup where id=1",conn,1,1
GroupID=rs("GroupID")
GroupName=rs("GroupName")
GroupLevel=rs("GroupLevel")
Explain=rs("Explain")
rs.close:set rs=Nothing


Set rs = server.CreateObject("adodb.recordset")
sql = "select * from hygroup  "
rs.Open sql, conn, 1, 3
rs.addnew

rs("GroupID")=GroupID
rs("GroupName")=GroupName
rs("GroupLevel")=GroupLevel
rs("Explain")=Explain
rs("adddate")=Now()

rs("lang")=langbs1
rs.update
    rs.Close
    Set rs = Nothing


set rs=server.CreateObject("adodb.recordset")
rs.open "select * from hygroup where id=2",conn,1,1
GroupID=rs("GroupID")
GroupName=rs("GroupName")
GroupLevel=rs("GroupLevel")
Explain=rs("Explain")
rs.close:set rs=Nothing


Set rs = server.CreateObject("adodb.recordset")
sql = "select * from hygroup  "
rs.Open sql, conn, 1, 3
rs.addnew

rs("GroupID")=GroupID
rs("GroupName")=GroupName
rs("GroupLevel")=GroupLevel
rs("Explain")=Explain
rs("adddate")=Now()

rs("lang")=langbs1
rs.update
    rs.Close
    Set rs = Nothing



    
	response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""lang.asp""</script>"

	
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
	 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>
<script language=javascript>
function deltp()
{

var parent=document.getElementById("mysmaimg");
var child=document.getElementById("sc");
var child1=document.getElementById("lj");
parent.removeChild(child);
parent.removeChild(child1);
document.getElementById("weepic").value="";
}

 </script>

</head>

<body>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加语言</div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="lang.asp" target="main">多语言管理</a></div></td>
 

 </tr>
</table>
<form  name="myform" class="checkform" method="post">
<table width="99%"  border="0" cellpadding="0" cellspacing="0" class="dtable4" align="center">

<tr> 
<td width="12%" >语言名称:</td>
<td width="88%">
  <input name="langname" type="text" id="langname" size="40" maxlength="50" datatype="*" nullmsg="语言名称不能为空"></td>
</tr>
<tr>
  <td >语言标识:</td><td>
    <input name="langbs" type="text" id="langbs" value="" size="40" maxlength="50" datatype="*" nullmsg="语言标识不能为空"></td>
  </tr>

 <tr >

			    <td >语言图标：
				 	        </td>
                <td colspan="2">  <input name="SmallImage"  type="hidden" id="weepic"  value="" size="50"  /> <br/> <br/>  	 <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
		        <div id="mysmaimg"></div></td>
	      </tr>


  <tr>
  <td >货币符号:</td><td>
    <input name="huobi" type="text" id="langbs" value="" size="40" maxlength="50" datatype="*" nullmsg="货币符号不能为空"></td>
  </tr>
  <tr>
  <td >货币文字:</td><td>
    <input name="huobi1" type="text" id="langbs" value="" size="40" maxlength="50" datatype="*" nullmsg="货币文字不能为空"></td>
  </tr>
 <tr>
  <td >排序:</td><td>
    <input name="px" type="text" id="px" value="0" size="40" maxlength="50" datatype="*" ></td>
  </tr>
<tr> 
<td  colspan="2"><input name="Submit" type="submit" class="bnt" value="添加"></td>
</tr>


</table>
</form>
</body>
</html>
