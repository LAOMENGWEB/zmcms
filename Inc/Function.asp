<%
Function str_rnd()
    Dim ran_num, dt_now, tmp
    dt_now = Now()
    Randomize
    ran_num = Int( (90000 * Rnd) + 10000 )
    tmp = Year(dt_now) & Right("0"&Month(dt_now), 2) & Right("0"&Day(dt_now), 2) & Right("0"&Hour(dt_now), 2) &_
    Right("0"&Minute(dt_now), 2) & Right("0"&Second(dt_now), 2) & ran_num
    str_rnd = tmp
End Function
'函数：三元IF
'外部提交
Function ChkPost()
  Dim url1,url2
  chkpost=true
  url1=Cstr(Request.ServerVariables("HTTP_REFERER"))
  url2=Cstr(Request.ServerVariables("SERVER_NAME"))
  If Mid(url1,8,Len(url2))<>url2 Then
    chkpost=false
    exit function
  End If
End function
Function IIf(Exp, v1, v2)
    Dim tmp
    tmp = v2
    If Exp Then tmp = v1
    IIf = tmp
End Function
'函数:读取ASP类型文件的全部内容

Function file_iread(Path)
    Dim Str
    Str = file_read(Path)
    Dim Pattern
    Pattern = "<\!--#include[ ]+?file[ ]*?=[ ]*?""(\S+?)""--\>"
    Dim matches
    Set matches = str_execute(Pattern, Str)
    Dim m, f, tmp
    For Each m in matches
        f = Mid(Path, 1, instrrev(Path, "/"))&m.submatches(0)
        tmp = file_read(f)
        If str_test(Pattern, tmp) Then tmp = file_iread(f) '处理子包含
        Str = Replace(Str, m.Value, tmp)
    Next
    Pattern = "<%@[ ]*?LANGUAGE[ ]*?=[ ]*?""[a-zA-Z]+?""[ ]+?CODEPAGE[ ]*?=[ ]*?""[0-9]+?""[ ]*?%\>"
    Str = str_replace(Pattern, Str, "")
    file_iread = Str
End Function

Sub echo(Str)
    response.Write(Str)
End Sub

'过程：结束页面并输出字符串

Sub die(Str)
    response.Write(Str)
    response.End()
End Sub


Function nohtml(str) 
dim re 
Set re=new RegExp 
re.IgnoreCase =true 
re.Global=True 
re.Pattern="{(.[^}]*)}" 
str=re.replace(str,"") 
nohtml=str 
set re=nothing 
End Function
Function nohtml1(str) 
dim re 
Set re=new RegExp 
re.IgnoreCase =true 
re.Global=True 
re.Pattern="(\[.[^\[]*\])" 
str=re.replace(str,"") 
re.Pattern="(\[\/[^\[]*\])" 
str=re.replace(str,"") 
nohtml1=str 
set re=nothing 
End Function
Function FormatImg(content) 
dim re 
Set re=new RegExp 
re.IgnoreCase =true 
re.Global=True 
re.Pattern="(script)" 
Content=re.Replace(Content,"script") 
re.Pattern="<img.[^>]*src(=| )(.[^>]*)>" 
Content=re.replace(Content,"<a href=$2 class=zmcmsfdtp target=""_blank""><img src=$2 title=""追梦工作室"" border=""0""/></a>") 
set re = nothing 
FormatImg = content 
End Function
function unrepreg(str)
	unrepreg = replace(replace(inull(str),"@@@(@@@","{"),"@@@)@@@","}")
end function
function replaceGlobal(templateStr)
	on error resume next
	  Dim regEx, Match, Matches,RetStr
	  Set regEx = New RegExp
	  regEx.Pattern = "{([^}]+)}"
	  regEx.IgnoreCase = False 
	  regEx.Global = True
	  Set Matches = regEx.Execute(templateStr)
	  RetStr = templateStr
	  For Each Match in Matches  '标签会引起死循环吗？
		RetStr = replace(RetStr,Match.Value,repreg(T(Match.SubMatches(0))))  '为防止循环只能套一次
	  Next
	replaceGlobal = unrepreg(RetStr)
	if err.number <> 0 then
		err.clear
		replaceGlobal = templateStr
	end if
end function
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function
function ChangeChr(str) 
ChangeChr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","") 
end function 
Function CutStrX(ByVal Str,ByVal StrLen)
Dim l,t,c,i,r     
'过滤全部HTML标记
Set r=New RegExp
r.Global=True
r.MultiLine=True
r.Pattern="(</?[A-Za-z][A-Za-z0-9]*[^>]*>)"
str=r.Replace(str," ")
Set r=Nothing    
l=Len(str)
t=0
For i=1 To l
c=AscW(Mid(str,i,1))
If c<0 Or c>255 Then t=t+2 Else t=t+1
IF t>=StrLen Then
CutStrX=Left(Str,i)&"..."
Exit For
Else
CutStrX=Str
End If
Next
End Function


'动态包含,适用与语言包加载-语言包编码UTF-8
Function include(filename)
		Dim re,content,fso,f,aspStart,aspEnd
		Set objStm = Server.CreateObject("ADODB.Stream")
			objStm.charset = "UTF-8"
			objStm.open
			objStm.LoadFromFile server.mappath(filename)
			content = objStm.Readtext
		Set objStm = Nothing
		set re=new RegExp
		re.pattern="^\s*="
		aspEnd=1
		aspStart=inStr(aspEnd,content,"<%")+2
		do while aspStart>aspEnd+1
		Response.write Mid(content,aspEnd,aspStart-aspEnd-2)
		aspEnd=inStr(aspStart,content,"%\>")+2
		Executeglobal(re.replace(Mid(content,aspStart,aspEnd-aspStart-2),"Response.Write "))
		aspStart=inStr(aspEnd,content,"<%")+2
		loop
		Response.write Mid(content,aspEnd)
		set re=nothing
End Function

	





Function DateStringFromNow(Byval sTheDate)
' 格式化显示时间为几个月,几天前,几小时前,几分钟前,或几秒前    
Dim iSeconds, iMinutes, iHours, iDays    
iSeconds = DateDiff("s", sTheDate, Now())  'd/h/n/s    
iMinutes = Int(iSeconds/60)    
iHours = Int(iSeconds/3600)    
iDays = Int(iSeconds/86400)
   
If iDays >= 60 Then       
 DateStringFromNow = sTheDate   
  ElseIf iDays >= 30 Then       
   DateStringFromNow = "1个月前"   
    ElseIf iDays >= 14 Then        
	DateStringFromNow = "2周前"    
	ElseIf iDays >= 7 Then        
	DateStringFromNow = "1周前"    
	ElseIf iDays >= 1 Then        
	DateStringFromNow = iDays & "天前"    
	ElseIf iHours >= 1 and iDays < 1 Then        
	DateStringFromNow = iHours & "小时前"   
	 ElseIf iMinutes >= 1 and iHours<1 Then        
	 DateStringFromNow = iMinutes & "分钟前"    
	 ElseIf iMinutes < 1 Then        
	 DateStringFromNow = iSeconds & "秒前"   
	  Else       
	  DateStringFromNow = "1秒前"   
	   End If
	   End Function
	
	'**************************************************
'函数作用:内容页TAB正则查找出标题
'**************************************************	 
Function Retitle(inpstr) 
Dim oRe, oMatch, oMatches
  Set oRe = New RegExp
'正则表达式
  oRe.Pattern = "\{标题:(.+?)}"
  ' 得到 Matches 集合
  Set oMatches = oRe.Execute(inpStr)
  ' 得到 Matches 集合中的第一项
  Set oMatch = oMatches(0)
'得到第一个匹配项的第一个匹配值
 Retitle= oMatch. SubMatches(0)
End Function
Function Deltitle(strng) '正则删掉内容中标签
Set   regEx   =   New   RegExp
regEx.Pattern   =  "\{标题:(.+?)}"
regEx.IgnoreCase   =   True
regEx.Global   =   True
Set   Matches   =   regEx.Execute(strng)
For   Each   Match   in   Matches
strng   =   regEx.Replace(strng,   "")
 Deltitle= (strng& "")
Next
End Function
	Function FilterHTML(str)
    Dim re,cutStr
    Set re=new RegExp
    re.IgnoreCase =True
    re.Global=True
    re.Pattern="<(.[^>]*)>"
    str=re.Replace(str,"")   
    set re=Nothing
	str=Replace(str,"script","")
	str=Replace(str,"alert","")	
	str=Replace(str,"%3E","")
	str=Replace(str,"%3C","")
    str=Replace(str,chr(10),"")
    str=Replace(str,chr(13),"")
    Dim l,t,c,i
    l=Len(str)
    t=0
    For i=1 to l
        c=Abs(Asc(Mid(str,i,1)))
        If c>255 Then
            t=t+2
        Else
            t=t+1
        End If
        cutStr=str
    Next
    
    FilterHTML=cutStr
End Function
	
	
	
	
	
	
	   
'================================================
 '函数名：FormatDate
 '作  用：格式化日期
 '参  数：DateAndTime   ----原日期和时间
 '        para   ----日期格式
 '返回值：格式化后的日期
 '================================================
Public Function FormatDate(DateAndTime, para)
On Error Resume Next
Dim y, m, d, h, mi, s, strDateTime
FormatDate = DateAndTime
If Not IsNumeric(para) Then Exit Function
If Not IsDate(DateAndTime) Then Exit Function
y = CStr(Year(DateAndTime))
m = CStr(Month(DateAndTime))
If Len(m) = 1 Then m = "0" & m
d = CStr(Day(DateAndTime))
If Len(d) = 1 Then d = "0" & d
h = CStr(Hour(DateAndTime))
If Len(h) = 1 Then h = "0" & h
mi = CStr(Minute(DateAndTime))
If Len(mi) = 1 Then mi = "0" & mi
s = CStr(Second(DateAndTime))
If Len(s) = 1 Then s = "0" & s
Select Case para
Case "1"
strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
Case "2"
strDateTime = y & "-" & m & "-" & d
Case "3"
strDateTime = y & "/" & m & "/" & d
Case "4"
strDateTime = y & "年" & m & "月" & d & "日"
Case "5"
strDateTime = m & "-" & d & " " & h & ":" & mi
Case "6"
strDateTime = m & "/" & d
Case "7"
strDateTime = m & "月" & d & "日"
Case "8"
strDateTime = y & "年" & m & "月"
Case "9"
strDateTime = y & "-" & m
Case "10"
strDateTime = y & "/" & m
Case "11"
strDateTime = right(y,2) & "-" &m & "-" & d & " " & h & ":" & mi
Case "12"
strDateTime = right(y,2) & "-" &m & "-" & d
Case "13"
strDateTime = m & "-" & d
Case Else
strDateTime = DateAndTime
End Select
FormatDate = strDateTime
End Function



'留言过滤脏话
Function ChkSB(str)
ChkSB = True
For i=0 To Ubound(Split(KillWord,","))
If Instr(Str,Split(KillWord,",")(i)) > 0 Then
ChkSB = False
Exit Function
End If
Next
End Function

function WriteMsg1(type1,Message,url,c)

response.Redirect(""&sitepath&"msginfo/?type1="&type1&"&message="&Message&"&url="&url&"&c="&c&"")

end function

function WriteMsg(type1,Message,url,c)
%>
<style type="text/css">
body, h1, p, a, img {margin:0;padding:0;}
body{text-align:center;font-family:Arial, Helvetica, sans-serif, "宋体";}
h1,p{padding-left:100px; text-align:left;}
h1{font-size:14px; margin-top:12px;}
p{font-size:12px; line-height:150%;padding:10px 10px 10px 100px;line-height:18px;}
.box_border{margin:180px auto 0;width:420px; padding:20px 0;}
#right{border:1px solid #A6DFA6; background:#EEF9EE url(<%=sitepath%>images/bg_right.gif) no-repeat 10px center;}
#right h1{color:#237E29;}
#right a:link,a:visited{color: #237E29;text-decoration: none;}
#right a:hover,a:active{color: #237E29;text-decoration: underline;}
#wrong{border:1px solid #FDBD77; background:#FFFDD7 url(<%=sitepath%>images/bg_wrong.gif) no-repeat 10px center;}
#wrong h1{color:#f30;}
#wrong a:link,a:visited{color: #f30;text-decoration: none;}
#wrong a:hover,a:active{color: #0066FF;text-decoration: underline;}
#right a{color: #237E29;}
#wrong a{color: #f30;}
</style>

<div class="box_border"  id="<%=c%>">
<h1><%=Message%></h1>
<p>
<%if type1=1 then%>
<a href="javascript:history.back(-1);" ><%=word(622)%></a>
<script language="javascript">setTimeout(function(){history.back();},3000);</script>
<%end if%>
<%if type1=2 then%>
<a href="<%=sitepath%><%=url%>"><%=word(622)%></a>
	<script language="javascript">setTimeout("location.href='<%=sitepath%><%=url%>';",3000);</script> 
<%end if%>
</p>
</div>
<%
end function

'会员屏蔽词
Function pbhy(str)
pbhy = True
For i=0 To Ubound(Split(notzc,","))
If Instr(Str,Split(notzc,",")(i)) > 0 Then
pbhy = False
Exit Function
End If
Next
End Function
function nbq(xx,yy)
set reg=new regexp
reg.IgnoreCase = False 
reg.Global = True 
reg.pattern=""&yy&"=([a-zA-Z0-9_,]+)"
set matches=reg.execute(xx)
for each match in matches
aa=match.submatches(0)
next
nbq=aa
end Function

function zmcms(str)
set reg=new regexp
reg.IgnoreCase = False 
reg.Global = True 
reg.Pattern = "{zmcms:([0-9]+)?}"
set aa=reg.execute(str)
for each match in aa
dbq=match.submatches(0)
ybq=match
next
zmcms=str
end function


function lmdh(style)
mb=getTagStr(style)
zmcms1=zmcms(mb)
zmcms1=memu(zmcms1):zmcms1=Memu1(zmcms1):zmcms1=Memu2(zmcms1):zmcms1=lanmu(zmcms1):zmcms1=lanmu1(zmcms1):zmcms1=lanmu2(zmcms1):zmcms1=lanmu3(zmcms1):zmcms1=lanmu4(zmcms1):zmcms1=ireplace(zmcms1)
lmdh=zmcms1
end function

function strLength(str)
ON ERROR RESUME NEXT
dim WINNT_CHINESE
WINNT_CHINESE    = (len("会员")=2)
if WINNT_CHINESE then
dim l,t,c
dim i
l=len(str)
t=l
for i=1 to l
c=asc(mid(str,i,1))
if c<0 then c=c+65536
if c>255 then
t=t+1
end if
next
strLength=t
else 
strLength=len(str)
end if
if err.number<>0 then err.clear
end function
function IsValidEmail(email)
dim names, name2, i, c
IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
IsValidEmail = false
exit function
end if
End Function
Function rdnum
randomize(time)
ood=fix(Rnd*1000000000000000)+999999999
rdnum=ood
End Function
Function CheckHtml(Str)
Dim Sos
Sos=Trim(Str)
if InStr(1,Sos," ",vbTextCompare)<>0 or InStr(1,Sos,".",vbTextCompare)<>0 or InStr(1,Sos,"<",vbTextCompare)<>0 or InStr(1,Sos,">",vbTextCompare)<>0 or InStr(1,Sos,"&",vbTextCompare)<>0 then
CheckHtml=true
else
 CheckHtml=false
end if      
End Function
function UBB_flv(strText)
	dim strContent
	dim re,Test
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[flv=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		strContent=re.replace(strContent, chr(1) & "flv=$1,$2" & chr(2))
		re.Pattern="\[\/flv\]"
		Test=re.Test(strContent)
		if Test then
			strContent=re.replace(strContent, chr(1) & "/flv" & chr(2))
				re.Pattern="\x01flv=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/flv\x02"
				strContent=re.Replace(strContent,"<div style=""text-align:center;""><EMBED pluginspage=http://www.macromedia.com/go/getflashplayer src=../images/flvplayer.swf width=$1 height=$2 type=application/x-shockwave-flash allowfullscreen=true flashvars=""file=$3&amp;autostart=true"" quality=high play=true loop=true></div>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
		end if
		re.Pattern="\x01"
		strContent=re.replace(strContent, "[")
	end if
	set re=Nothing
	UBB_flv=strContent
end function
Function UBBCode(strContent)
UBBCode=ReplaceKey(UBB_flv(strContent))
End Function


Function ReplaceKey(ByVal Str)
		If IsNull(Str) Then Exit Function
		Dim RsKeyword
		sql="Select * From [gjz] Order By [Num] Desc"
		Set RsKeyword=Conn.Execute(sql)
		do while not (RsKeyword.eof or err)
		If InStr(Str,RsKeyword("Title")) > 0 Then
		oReplace = RsKeyword("Replace")
		If oReplace=0 then oReplace=-1
		Str = p_replace(Str,RsKeyword("Title"),"<a href="""&RsKeyword("Url")&""" target=""_blank"">"&RsKeyword("Title")&"</a>",1,oReplace,1)
		End if	
		RsKeyword.movenext
		loop
		ReplaceKey = Str
		RsKeyword.Close:Set RsKeyword=nothing
End Function

Function p_replace(byval content,byval asp,byval htm,byval aa,byval Rnum,byval bb)
dim Matches,objRegExp,strs,i
strs=content
Set objRegExp = New Regexp
objRegExp.Global = True
objRegExp.IgnoreCase = True
objRegExp.Pattern = "(\<a[^<>]+\>.+?\<\/a\>)|(\<img[^<>]+\>)"'
Set Matches =objRegExp.Execute(strs)
i=0
Dim MyArray()
For Each Match in Matches
ReDim Preserve MyArray(i)
MyArray(i)=Mid(Match.Value,1,len(Match.Value))
strs=replace(strs,Match.Value,"<"&i&">",1,Rnum,1)
i=i+1
Next
if i=0 then
 content=replace(content,asp,htm,1,Rnum,1)
 p_replace=content
 exit Function
end if
strs=replace(strs,asp,htm,1,Rnum,1)
for i=0 to ubound(MyArray)
strs=replace(strs,"<"&i&">",MyArray(i),1,Rnum,1)
next
p_replace=strs
end Function

Private Function UrlEncode_GBToUtf8(szInput)
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    If szInput = "" Then
        UrlEncode_GBToUtf8= szInput
        Exit Function
    End If
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)
        If nAsc < 0 Then nAsc = nAsc + 65536
        If wch = "+" then
            szRet = szRet & "%2B"
        ElseIf wch = "%" then
            szRet = szRet & "%25"
        ElseIf (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    UrlEncode_GBToUtf8 = szRet
End Function

function StrReplace(Str)'表单存入替换字符
  if Str="" or isnull(Str) then 
    StrReplace=""
    exit function 
  else
    StrReplace=replace(str," ","&nbsp;") '"&nbsp;"
    StrReplace=replace(StrReplace,chr(13),"&lt;br&gt;")'"<br>"
    StrReplace=replace(StrReplace,"<","&lt;")' "&lt;"
    StrReplace=replace(StrReplace,">","&gt;")' "&gt;"
  end if
end function

function IsValidMemName(memname)
  dim i, c
  IsValidMemName = true
  if not (3<=len(memname) and len(memname)<=16) then
    IsValidMemName = false
    exit function
  end if  
  for i = 1 to Len(memname)
    c = Mid(memname, i, 1)
    if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-", c) <= 0 and not IsNumeric(c) then
      IsValidMemName = false
      exit function
    end if
  next
end function
function IsValidEmail(email)
  dim names, name, i, c
  IsValidEmail = true
  names = Split(email, "@")
  if UBound(names) <> 1 then
    IsValidEmail = false
    exit function
  end if
  for each name in names
	if Len(name) <= 0 then
	  IsValidEmail = false
      exit function
    end if
    for i = 1 to Len(name)
      c = Mid(name, i, 1)
      if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-.", c) <= 0 and not IsNumeric(c) then
        IsValidEmail = false
        exit function
      end if
	next
	if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
    end if
  next
  if InStr(names(1), ".") <= 0 then
    IsValidEmail = false
    exit function
  end if
  i = Len(names(1)) - InStrRev(names(1), ".")
  if i <> 2 and i <> 3 then
    IsValidEmail = false
    exit function
  end if
  if InStr(email, "..") > 0 then
    IsValidEmail = false
  end if
end function

function ViewNoRight(GroupID,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from hyGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  ViewNoRight=true
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then   ' 当用户权限数字小于模块权限数字, 则无权访问
	    ViewNoRight=false
	  end if
    case "="
      if not session("GroupLevel") = GroupLevel then
	    ViewNoRight=false
      end if
  end select
end function


function HtmlStrReplace(Str)'写入Html网页替换字符
  if Str="" or isnull(Str) then 
    HtmlStrReplace=""
    exit function 
  else
    HtmlStrReplace=replace(Str,"&lt;br&gt;","<br>")'"<br>"
  end if
end function
Function RndNumber(Min,Max) 
Randomize 
RndNumber=Int((Max - Min + 1) * Rnd() + Min) 
End Function


Function Alert(message,gourl) 
    message = replace(message,"'","\'")
    If gourl="-1" then
        Response.Write ("<script language=javascript>alert('" & message & "');history.go(-1)</script>")
    ElseIf gourl="-2" then
        Response.Write ("<script language=javascript>alert('" & message & "');history.go(-2)</script>")
    ElseIf gourl="Close" then
		Response.Write ("<script language=javascript>alert('" & message & "');window.opener=null;window.close();</script>")
	Else
        Response.Write ("<script language=javascript>alert('" & message & "');location='" & gourl &"'</script>")
    End If
    Response.End()
End Function


%>


<%
Function   Jmail(mailTo,mailTopic,mailBody,mailCharset,mailContentType)   
Dim   ConstFromNameCn,ConstFromNameEn,ConstFrom,ConstMailDomain,ConstMailServerUserName,ConstMailServerPassword   
    
          ConstFromNameCn   =   ""&MailUserName&""   
          ConstFromNameEn   =   ""&MailUserNameen&""   
          ConstFrom   =   ""&sendmail&""   
          ConstMailDomain   =   ""&MailServer&""   
          ConstMailServerUserName   =   ""&MailAdd&""   
          ConstMailServerPassword   =   ""&MailPassword&""'   
  
          On   Error   Resume   Next   
          Dim   myJmail   
          Set   myJmail   =   Server.CreateObject("JMail.Message")   
          myJmail.Logging   =   False 
          myJmail.ISOEncodeHeaders   =   False 
          myJmail.ContentTransferEncoding   =   "base64"   
          myJmail.AddHeader   "Priority","3"   
          myJmail.AddHeader   "MSMail-Priority","Normal"   
          myJmail.AddHeader   "Mailer","Microsoft   Outlook   Express   6.00.2800.1437"   
          myJmail.AddHeader   "MimeOLE","Produced   By   Microsoft   MimeOLE   V6.00.2800.1441"   
          myJmail.Charset   =   mailCharset   
          myJmail.ContentType   =   mailContentType   
    
          If   UCase(mailCharset)   =   "GB2312"   Then   
                  myJmail.FromName   =   ConstFromNameCn   
          Else   
                  myJmail.FromName   =   ConstFromNameEn   
          End   If   
    
          myJmail.From   =   ConstFrom   
          myJmail.Subject   =   mailTopic   
          myJmail.Body   =   mailBody   
          myJmail.AddRecipient   mailTo   
          myJmail.MailDomain   =   ConstMailDomain   
          myJmail.MailServerUserName   =   ConstMailServerUserName   
          myJmail.MailServerPassword   =   ConstMailServerPassword   
          myJmail.Send   ConstMailDomain   
          myJmail.Close   
          Set   myJmail=nothing     
    
          If   Err   Then     
                  Jmail=Err.Description   
                  Err.Clear   
          Else   
                  Jmail="OK"   
          End   If   
    
          On   Error   Goto   0   
End   Function  

function inull(s)
	dim str
	if isnull(s) then
	 str = ""
	else
	 str = s
	end if
	inull = str
end function

function snull(s)
	dim str
	if isnull(s) then
	 str = 0
	else
	 str = s
	end if
	snull = str
end function

function eint(s)
	dim str
	if isnull(s) or s = "" then
	 str = 0
	else
	 on error resume next
	 str = int(s)
	 if err.number <> 0 then
	 	err.clear
		str = 0
	 end if
	end if
	eint = str
end function

Function RegExpRet(strng,patrn,n)
  on error resume next
  Dim regEx, retVal ,RetStr,Matches ,Match
  Set regEx = New RegExp
  regEx.Pattern = patrn
  regEx.IgnoreCase = False
  set Matches  = regEx.Execute(strng)
  set Match = Matches(0)
  if err.number = 0 then
  	RetStr = Match.SubMatches(n)
  else
  	err.clear
  	RetStr = ""
  end if
  RegExpRet = RetStr
End Function

Function RegExpTest(strng, patrn)
  Dim regEx, retVal
  Set regEx = New RegExp 
  regEx.Pattern = patrn
  regEx.IgnoreCase = False '不区分大小写
  retVal = regEx.Test(strng)
  If retVal Then
    RegExpTest = true
  Else
    RegExpTest = false
  End If
End Function
function getTagStr(tagName)
set rs1=server.CreateObject("adodb.recordset")
	on error resume next
	'response.Write("##########" & tagName & "####$")
	rs1.open "select top 1 tagcontent from templateTags where tagName = '"&tagName&"'",Conn3,0,1
	if not rs1.eof then 
	if sqlno=1 then
		getTagStr = inull(rs1(0))
		end if
		if sqlno=2 then
		getTagStr = rs1(0)
		end if
	else
		getTagStr = ""
	end if
	response.Write(err.description)
	rs1.Close
	if err.number <> 0 then getTagStr = ""	
end function

Function RemoveHTML(strHTML)
	on error resume next
	if inull(strHTML) = "" then
		RemoveHTML = ""
		exit function
	end if
	Dim objRegExp, Match, Matches
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = False
	objRegExp.Global = True
	objRegExp.Pattern = "(\<.[^\>]+\>)"
	Set Matches = objRegExp.Execute(strHTML)
	For Each Match in Matches
		strHtml=Replace(strHTML,Match.Value,"")
	Next
	RemoveHTML=strHTML
	Set objRegExp = Nothing
	if err.number <> 0 then RemoveHTML = ""
End Function

Function ReplaceTest(str1,patrn, replStr)
  Dim regEx               ' 建立变量。
  Set regEx = New RegExp               ' 建立正则表达式。
  regEx.Global = True
  regEx.Pattern = patrn               ' 设置模式。
  regEx.IgnoreCase = True               ' 设置是否区分大小写。
  ReplaceTest = regEx.Replace(str1, replStr)         ' 作替换。
End Function


Function RegExpDiyTag(TempStr,diytagValue)
  Dim regEx, Match, Matches,RetStr
  Set regEx = New RegExp
  regEx.Pattern = "{diy:([^}]+)}"
  regEx.IgnoreCase = True 
  regEx.Global = True
  Set Matches = regEx.Execute(TempStr)
  RetStr = TempStr
  For Each Match in Matches
	RetStr = replace(RetStr,Match.Value,RegExpRet("$$$"&diytagValue, "\$\$\$" & trim(Match.SubMatches(0)) & "\|\|([^\$]+)\$\$\$",0))
  Next
  RegExpDiyTag = RetStr
End Function


Function MonthToShortEN(TheMonth)
Dim vvvv
vvvv=split("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",",")
If IsNumeric(TheMonth) Then
  MonthToShortEN = vvvv(TheMonth-1)
Else
  MonthToShortEN = " "
End If
End Function
Function MonthToLongEN(TheMonth)
Dim vvvv
vvvv=split("January,February,March,APRil,May,June,July,August,September,October,November,December",",")
If IsNumeric(TheMonth) Then
  MonthToLongEN = vvvv(TheMonth-1)
Else
  MonthToLongEN = " "
End If
End Function
function FormatTime(s_Time,s_Flag)
	Dim y, m, d, h, mi, f, bbb, ccc, ReturnStr
	FormatTime = ""
	If IsDate(s_Time) = False Then Exit Function
	y = year(s_Time)
	m = month(s_Time)
	bbb=MonthToLongEN(m)
	ccc=MonthToShortEN(m)
	d = day(s_Time)
	h = hour(s_Time)
	mi = minute(s_Time)
	f = second(s_Time)
	ReturnStr = lcase(s_Flag)
	ReturnStr = replace(ReturnStr,"yyyy",y)
	ReturnStr = replace(ReturnStr,"yy",right(y,2))
	ReturnStr = replace(ReturnStr,"mm",right("00" & m,2))
	ReturnStr = replace(ReturnStr,"m",m)	
	ReturnStr = replace(ReturnStr,"bbb",bbb)
	ReturnStr = replace(ReturnStr,"ccc",ccc)
	ReturnStr = replace(ReturnStr,"dd",right("00" & d,2))
	ReturnStr = replace(ReturnStr,"d",d)
	ReturnStr = replace(ReturnStr,"hh",right("00" & h,2))
	ReturnStr = replace(ReturnStr,"h",h)
	ReturnStr = replace(ReturnStr,"nn",right("00" & mi,2))
	ReturnStr = replace(ReturnStr,"n",mi)
	ReturnStr = replace(ReturnStr,"ff",right("00" & f,2))
	ReturnStr = replace(ReturnStr,"f",f)
	FormatTime = ReturnStr
end function
function TestCaptcha(byval valSession, byval valCaptcha)
	dim tmpSession
	valSession = Trim(valSession)
	valCaptcha = Trim(valCaptcha)
	if (valSession = vbNullString) or (valCaptcha = vbNullString) then
		TestCaptcha = false
	else
		tmpSession = valSession
		valSession = Trim(Session(valSession))
		Session(tmpSession) = vbNullString
		if valSession = vbNullString then
			TestCaptcha = false
		else
			valCaptcha = Replace(valCaptcha,"i","I")
			if StrComp(valSession,valCaptcha,1) = 0 then
				TestCaptcha = true
			else
				TestCaptcha = false
			end if
		end if		
	end if
end function





Public Function Safereplace(Str)
	    If Isnull(Str) Then
	    	    Safereplace = ""
	    	    Exit Function
	    End If
	    Str = Replace(Str,";","")
	    Str = Replace(Str,"'","&#180;")
	    Str = Replace(Str,"%","")
	    Str = Replace(Str,"""","&#34;")
        Str = Replace(Str,"<","&lt;")
        Str = Replace(Str,">","&gt;")
	    Str = Replace(Str,"$","")
	    Safereplace = Str
End Function



Public function formatnumber2(num)
		Langdecimal = checknum(Langdecimal,2)
        If IsNull(num) Or Not IsNumeric(num) Then
		        Exit Function
	    End If
		If Langdecimal = 0 Then 
				If num < 1 Then 
						num = formatnumber(num,2)
				Else 
						num = formatnumber(num,Langdecimal)
				End If 
		Else 
				num = formatnumber(num,Langdecimal)
		End If 
	    if num <1 And num >= 0 then
	            num= 0&num
	            num = replace(num,"00.","0.")
		ElseIf num = 0 Then
				num= 0&num
	    end If
	    If Langdecimal = 0 Then num = replace(num,".00","")
	    formatnumber2 = Replace(num,",","")
end Function



Public function judge(a,b)
	    dim res
	    if instr(a,b)>0 then
	    	    res ="checked"
	    end if
	    judge = res
end Function




'图片按比例缩放函数 
'****  Start  ************************************************************************************************************************
Public Function showpicsize(Owidth,Oheight,Vwidth,Vheight)
	    Dim str_1,str_2
        If Owidth="" Then
	    	    showpicsize="100|100"
        Exit Function
        End If
	    If clng(Owidth)<clng(Vwidth) And clng(Oheight)<clng(Vheight) Then
	    	    showpicsize=Owidth&"|"&Oheight
        Exit Function
        End If
        If clng(Owidth)>clng(Oheight) Then
	    	    str_1 = Vwidth
	    	    str_2 = clng(Vwidth * Oheight / Owidth)
        Else
	    	    str_1 = clng(Vheight * Owidth / Oheight)
	    	    str_2 = Vheight
	    End If
        showpicsize=str_1&"|"&str_2
End Function
'****  End   *************************************************************************************************************************

rem 截取字符串长度函数
rem UTF-8编码下 截取中英文混排长度
Public Function cutchar(Title,TLen)
		If title = "" Or IsNull(title) Then
				Exit Function
		End If
		Dim k,i,d,c
		Dim iStr
		Dim ForTotal
		If CDbl(TLen) > 0 Then
				k=0
				d=StrLen(Title)
				iStr=""
				ForTotal = Len(Title)
				For i=1 To ForTotal
						c=Abs(AscW(Mid(Title,i,1)))
						If c>255 Then
								k=k+2
						Else
								k=k+1
						End If
						iStr=iStr&Mid(Title,i,1)
						If CLng(k)>CLng(TLen*2) Then
								iStr=iStr&".."
								Exit For
						End If
				Next
				cutchar=iStr
		Else
				cutchar=""
		End If
		TLen = TLen
End Function







Public Function cut_Right(str)
        If str="" Then
		       Exit Function
	    End If
        cut_Right=Left(str,(Len(str)-1))
End Function

















Function CheckNum(num,back)
		num = Trim(num)
		If Not IsNumeric(num) Or IsEmpty(num) Then
				CheckNum = CLng(back)
		Else
				CheckNum = num
		End If
End Function

function judge(a,b)
        dim res
        if instr(","& a &",",","& b &",")>0 then
                res ="checked"
        end if
        judge = res
end Function

function judge2(a,b)
        dim res
        if instr(","& a &",",","& b &",")>0 then
               res ="selected"
        end if
        judge2= res
end Function





Function getN(str,keyword,splits) 
		getN = 0 
		arr = Split(str,splits) 
		getN = UBound(arr) 
		For nnnn = 0 To UBound(arr) 
		if CStr(arr(nnnn)) = CStr(keyword) then 
				getN = (nnnn-1) + 1 
				Exit Function  
		End If  
		Next  
End Function 

Function Kuadi100API(com,nu,valicode)
	If Split(sys_kuaidi100,"|")(0) = 1 Then 
		SendURL = "http://api.kuaidi100.com/api?id="& Split(sys_kuaidi100,"|")(1) &"&com="& safereplace(com) &"&nu="& safereplace(nu) &"&valicode="& safereplace(valicode) &"&show=2&muti=1&order=asc"
		Kuadi100API = GetHTTPPage(SendURL)
	End If 
End Function 
Function GetHTTPPage(URL) 
    Dim objXML 
    Set objXML=CreateObject("MSXML2.SERVERXMLHTTP.3.0")  '调用XMLHTTP组件，测试空间是否支持XMLHTTP，如果服务不支持，请测试下面两个。
	'Set objXML=Server.CreateObject("Microsoft.XMLHTTP") 
	'Set objXML=Server.CreateObject("MSXML2.XMLHTTP.4.0") 
	objXML.SetTimeouts 5000, 5000, 30000, 10000' 解析DNS名字的超时时间，建立Winsock连接的超时时间，发送数据的超时时间，接收response的超时时间。单位毫秒
    objXML.Open "GET",URL,False 'false表示以同步的方式获取网页代码，了解什么是同步？什么是异步？
    objXML.Send() '发送
	If objXML.Readystate<>4 Then
		Exit Function 
	End If
	'Readystate属性,传回XML文件资料的目前状况,返回值分别有以下：
	'0-UNINITIALIZED：XML 对象被产生，但没有任何文件被加载。
	'1-LOADING：加载程序进行中，但文件尚未开始解析。
	'2-LOADED：部分的文件已经加载且进行解析，但对象模型尚未生效。
	'3-INTERACTIVE：仅对已加载的部分文件有效，在此情况下，对象模型是有效但只读的。
	'4-COMPLETED：文件已完全加载，代表加载成功。
	'GetHTTPPage=objXML.ResponseBody
	GetHTTPPage=BytesToBstr(objXML.ResponseBody)'返回信息，同时用函数定义编码，如果您需要转码请选择 
    Set objXML=Nothing'关闭 
	If Err.number<>0 Then 
'		Response.Write "<p align='center'><font color='red'><b>服务器获取文件内容出错，请稍后再试！</b></font></p>" 
		Err.Clear
	Else
		GetHTTPPage = GetHTTPPage &"<span style=""font-size:12px"">查询数据由：快递100(<a href='http://kuaidi100.com' target='_blank'>www.kuaidi100</a>) 提供</span>"
	End If
End Function
	''通过xmlhtml方式获取远程文件内容
	public function gethttp(byval t0,byval t1)
		select case t1
			case "get","post"
			case else:t1="get"
		end select
		on error resume next
		dim http
		set http=server.createobject("microsoft.xmlhttp")
			http.open t1,t0,false
			http.setRequestHeader "If-Modified-Since","0"
			http.send()
			gethttp=http.responsetext
		set http=nothing
	end function
	function getip()
		getip=request.servervariables("http_x_forwarded_for")
		if getip="" then getip=request.servervariables("remote_addr")
	end function
%>
