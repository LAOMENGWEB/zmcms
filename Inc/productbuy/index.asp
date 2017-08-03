<!--#include file="../sub.asp"-->
<%
'初始化加密类
call initCodecs 
Response.Charset="utf-8"
if DateDiff("s",session("time"),now())<5 then  
WriteMsg 1,""&word(370)&"","","wrong"
response.end
else  
session("time")=now()
end if
if chkpost=0 then
WriteMsg 1,""&word(491)&"","","wrong"
response.end
end if
Randomize()
ranNum	=	int(90000*rnd)+10000
orderddh	=	year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
dim mOrderName,mRemark
dim mMemID,mRealName,mSex,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail,checkCode
dim rs,sql
mOrderName=trim(request.form("OrderName"))
zffs=trim(request.form("zffs"))
mRemark=request.form("Remark")
mMemID=request.QueryString("MemberID")
mRealName=trim(request.form("RealName"))
mSex=trim(request.form("Sex"))
mCompany=trim(request.form("Company"))
mAddress=trim(request.form("Address"))
mZipCode=trim(request.form("ZipCode"))
mFax=trim(request.form("Fax"))
mMobile=trim(request.form("Mobile"))
mEmail=trim(request.form("Email"))
checkCode=trim(request.form("checkCode"))
strShipment=split(Trim(Request("shipment")),"|")
shipmentID=strShipment(0)
shipmentMoney=strShipment(1)
shipmentName=strShipment(2)
If CheckCart = False Then
Response.Write ""&word(446)&""
Response.end
End If
aryCart = Session("Cart")
ii = 0
For mm = 0 to Ubound(aryCart,2)
set rsdetail = server.createobject("adodb.recordset")
sqldetail="select * from Ordercp"
rsdetail.open sqldetail,conn,1,3
rsdetail.AddNew 
rsdetail("productname")=aryCart(ii+2,mm)
rsdetail("productslt")=aryCart(ii+5,mm)
rsdetail("bookcount")=aryCart(ii+1,mm)
rsdetail("MemID")=mMemID
rsdetail("p_id")=aryCart(ii+6,mm)
rsdetail("hyj")=formatnumber2(aryCart(ii+4,mm))
rsdetail("ordersj")=now()
rsdetail("orderddh")=orderddh
rsdetail("lang")=language
rsdetail.Update
Next
set rs = server.createobject("adodb.recordset")
sql="select * from Orderdd"
rs.open sql,conn,1,3
rs.addnew
rs("OrderName")=StrReplace(mOrderName)
rs("Remark")=StrReplace(mRemark)
rs("MemID")=mMemID
rs("Linkman")=StrReplace(mRealName)
rs("Sex")=mSex
rs("Company")=StrReplace(mCompany)
rs("Address")=StrReplace(mAddress)
rs("ZipCode")=StrReplace(mZipCode)
rs("Fax")=StrReplace(mFax)
rs("Mobile")=StrReplace(mMobile)
rs("Email")=mEmail
rs("AddTime")=now()
rs("shipmentID") =shipmentID
rs("shipmentMoney") =shipmentMoney
rs("shipmentName") =shipmentName
rs("orderddh")=orderddh
rs("lang")=language
if sqlno=2 then
rs("isdate")=0
end if
rs("zffs")=zffs
rs("aumont")=session("auuont")
rs.update
rs.close
set rs=nothing
'清空购物车
Session("Cart") = Null
Session("Cart") = ""
session("orderddh")=orderddh

if maillock=1 and newdd=1 then 
  SendStat   =   Jmail(""&PostMail&"",""&word(447)&""&mRealName&""&word(448)&""&mOrderName&""&word(449)&"","<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><title>"&mOrderName&""&word(449)&"</title></head><body>"&word(450)&""&orderddh&"<br>"&word(451)&""&msex&"<br> "&word(452)&""&mCompany&"<br>"&word(453)&""&mAddress&"<br>"&word(454)&""&mZipCode&"<br>"&word(455)&""&mFax&"<br>"&word(456)&""&mEmail&"<br>"&word(457)&""&mMobile&"<br></body></html>","GB2312","text/html")   
end if
writemsg 2,""&word(458)&"","template/"&skins&"/member/pay.asp?orderddh="&base64Encode(orderddh)&"&m="&request.QueryString("m")&"&Language="&Language&"","right"
%>








