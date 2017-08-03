<!--#include file="../sub.asp"-->
<%
Response.Charset="utf-8"
if DateDiff("s",session("time"),now())<5 then  
WriteMsg 2,""&word(370)&"","member/membercenter/?m="&request.QueryString("m")&"&Language="&Language&"","wrong"
response.end
else  
session("time")=now()
end if
if chkpost=0 then
WriteMsg 1,""&word(491)&"","","wrong"
response.end
end if
dim mMemName,mRealName,mSex,mPassword,mQuestion,mAnswer,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail
dim vPassword,checkcode
dim rs,sql
mMemName=trim(request.form("MemName"))
mRealName=trim(request.form("RealName"))
mSex=trim(request.form("Sex"))
mPassword=trim(request.form("Password"))
vPassword=trim(request.form("vPassword"))
mQuestion=trim(request.form("Question"))
mAnswer=trim(request.form("Answer"))
mCompany=trim(request.form("Company"))
mAddress=trim(request.form("Address"))
mZipCode=trim(request.form("ZipCode"))
mFax=trim(request.form("Fax"))
mMobile=trim(request.form("Mobile"))
mEmail=trim(request.form("Email"))
checkCode=trim(request.form("checkCode"))
qqapi=trim(request.form("qqapi"))
sh=trim(request.form("sh"))
if maillock=1 and zchy=1 then 
SendStat   =   Jmail(""&mEmail&"",""&word(381)&""&mc&""&word(382)&"",""&word(383)&""&mc&""&word(3824)&"<br> "&word(385)&""&mMemName&"<br>"&word(386)&""&mPassword&"<br>"&word(387)&""&mQuestion&"<br>"&word(388)&""&mAnswer&"<br><br>"&word(389)&"<br><br>==============================<br>"&mc&"<br>"&word(390)&""&ym&"<br>"&word(391)&""&yx&"<br>"&word(392)&""&sj&"<br>"&word(393)&""&cz&"<br>"&word(394)&""&lxdz&"","GB2312","text/html")   

end if
set rs = server.createobject("adodb.recordset")
sql="select * from Members"
rs.open sql,conn,1,3
rs.addnew
rs("MemName")=mMemName
rs("RealName")=mRealName
rs("Sex")=mSex
rs("Password")=Md5(mPassword)
rs("Question")=mQuestion
rs("Answer")=mAnswer
rs("Company")=mCompany
rs("Address")=mAddress
rs("ZipCode")=mZipCode
rs("Fax")=mFax
rs("Mobile")=mMobile
rs("Email")=mEmail
rs("lang")=language
rs("qqopenapi")=qqapi
if hysh=1 then
rs("GroupID")="999177"
rs("GroupName")=GroupName("999177")
else
rs("GroupID")="1777680"
rs("GroupName")=GroupName("1777680")
end if
rs("AddTime")=now()
if hysh=1 then
rs("sh")=0
else
rs("sh")=1
end if
rs("working")=1
rs.update
rs.close
set rs=nothing

if qqapi="" then
if hysh=1 then
if maillock=1 and zchy=1 then
WriteMsg 2,""&word(395)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
else
writemsg 2, ""&word(396)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
end if
else
if maillock=1 and zchy=1 then
WriteMsg 2,""&word(397)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
else
writemsg 2, ""&word(398)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
end if
end if


else


if hysh=1 then
if maillock=1 and zchy=1 then
WriteMsg 2,""&word(597)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
else
writemsg 2, ""&word(598)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
end if
else
if maillock=1 and zchy=1 then
WriteMsg 2,""&word(599)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
else
writemsg 2, ""&word(600)&"","member/memberuserlogin/?m="&request.QueryString("m")&"&Language="&Language&"","right"
end if
end if

end if

function GroupName(GroupID)
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from hyGroup where GroupID='"&GroupID&"' and Lang='"&Language&"'"
  rs.open sql,conn,1,1
  GroupName=rs("GroupName")
  rs.close
  set rs=nothing
end function
%>


