<!--#include file="../sub.asp"-->
<%

Response.Charset="utf-8"
if DateDiff("s",session("time"),now())<5 then  
WriteMsg 2,""&word(370)&"","member/memberinfo/?m="&request.QueryString("m")&"&Language="&Language&"","wrong"
response.end
else  
session("time")=now()
end if
if chkpost=0 then
WriteMsg 1,""&word(491)&"","","wrong"
response.end
end if
dim ID,mRealName,mSex,mPassword,vPassword,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,mHomePage
dim checkcode
dim rs,sql
ID=request.QueryString("ID")
mRealName=trim(request.form("RealName"))
mSex=trim(request.form("Sex"))
mPassword=trim(request.form("Password"))
vPassword=trim(request.form("vPassword"))
mCompany=trim(request.form("Company"))
mAddress=trim(request.form("Address"))
mZipCode=trim(request.form("ZipCode"))
mFax=trim(request.form("Fax"))
mMobile=trim(request.form("Mobile"))
mEmail=trim(request.form("Email"))
mHomePage=trim(request.form("HomePage"))
checkcode=trim(request.form("checkcode"))
dim ErrMessage,ErrMsg(8),FindErr(8),i
ErrMsg(0)=""&word(372)&""
ErrMsg(1)=""&word(373)&""
ErrMsg(2)=""&word(374)&""
ErrMsg(3)=""&word(375)&""
ErrMsg(4)=""&word(376)&""
ErrMsg(5)=""&word(377)&""
ErrMsg(6)=""&word(378)&""
ErrMsg(7)=""&word(379)&""
if len(mPassword)>0 then
   if not (6<=len(mPassword) and len(mPassword)<=18) then 
   WriteMsg 1,ErrMsg(0),"","wrong"
   response.end
   end if
   if mPassword<>vPassword then 
   WriteMsg 1,ErrMsg(1),"","wrong"
   response.end
   end if
end if
if len(mMobile)>12  then 
WriteMsg 1,ErrMsg(4),"","wrong"
response.end
end if
if not IsValidEmail(mEmail) then 
WriteMsg 1,ErrMsg(5),"","wrong"
response.end
end if
if not conn.execute("select MemName from Members where ID<>"&ID&" and Email='" & mEmail & "'").eof then 
WriteMsg 1,ErrMsg(6),"","wrong"
response.end
end if
if TestCaptcha("ASPCAPTCHA", Request.Form("captchacode")) then
session("checkcode")=""
else
WriteMsg 1,ErrMsg(7),"","wrong"
response.end
end if
set rs = server.createobject("adodb.recordset")
sql="select * from Members where ID="&ID
rs.open sql,conn,1,3
rs("RealName")=StrReplace(mRealName)
rs("Sex")=mSex
if len(mPassword)>0 then rs("Password")=Md5(mPassword)
rs("Company")=StrReplace(mCompany)
rs("Address")=StrReplace(mAddress)
rs("ZipCode")=StrReplace(mZipCode)
rs("Fax")=StrReplace(mFax)
rs("Mobile")=StrReplace(mMobile)
rs("Email")=mEmail
rs("HomePage")=StrReplace(mHomePage)
rs.update
rs.close
set rs=nothing
WriteMsg 2,""&word(380)&"","member/memberinfo/?m="&request.QueryString("m")&"&Language="&Language&"","right"
%>
 