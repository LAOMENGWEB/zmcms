<%
Dim query, aa, x, temp
query = Split(Request.ServerVariables("QUERY_STRING"), "&")
For Each x In query
aa = Split(x, "=")
If StrComp(aa(0), "page", vbTextCompare) <> 0 Then
temp = temp & aa(0) & "=" & aa(1) & "&"
End If
Next
'=================================================
'过程名：ManualPagination
'作  用：采用手动分页方式显示文章具体的内容  asp版本
'参  数：str1,str2,str3
'=================================================
Function ManualPagination(str1,str2,str3)
ArticleId = str1
strContent = str2
ContentLen=strContent
CurrentPage=request("Page")
if Instr(strContent,"[zmcmsfy]")<=0 then
ManualPagination_Tmp = ManualPagination_Tmp & strContent
else
arrContent=split(strContent,"[zmcmsfy]")
pages=Ubound(arrContent)+1
pages=Ubound(arrContent)+1
if CurrentPage="" then
CurrentPage=1
else
CurrentPage=Cint(CurrentPage)
end if
if CurrentPage<1 then CurrentPage=1
ManualPagination_Tmp = ManualPagination_Tmp & arrContent(CurrentPage-1)
ManualPagination_Tmp=ManualPagination_Tmp&" <div id=""page"&str3&""" class=""page""></div>"& vbCrLf 
ManualPagination_Tmp=ManualPagination_Tmp&"<script type=""text/javascript"">"& vbCrLf  
ManualPagination_Tmp=ManualPagination_Tmp&"laypage({"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "cont: 'page"&str3&"',"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "pages: """&pages&""","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "skin: '"&word(498)&"',"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "first: "&word(499)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "prev: "&word(500)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "next: "&word(501)&","& vbCrLf
if word(502)<>"" then
ManualPagination_Tmp=ManualPagination_Tmp&  "last: "&word(502)&","& vbCrLf
end if
ManualPagination_Tmp=ManualPagination_Tmp&  "groups: "&word(503)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "skip: "&word(504)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"curr: """&CurrentPage&""","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"jump: function(e, first){"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"if(!first){"& vbCrLf
if yy=1 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
else
if str3=1 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href ='"&aboutname&""&Separated&"" & zmone & ""&Separated&""&ArticleId&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if str3=2 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href ='"&newname&""&Separated&"" & zmone & ""&Separated&""&ArticleId&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if str3=3 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href ='"&ProductName&""&Separated&"" & zmone & ""&Separated&""&ArticleId&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if str3=4 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href ='"&alName&""&Separated&"" & zmone & ""&Separated&""&ArticleId&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if str3=5 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href ='"&tvName&""&Separated&"" & zmone & ""&Separated&""&ArticleId&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if str3=6 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href ='"&downName&""&Separated&"" & zmone & ""&Separated&""&ArticleId&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if str3=7 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href ='"&zpName&""&Separated&""&ArticleId&""&Separated&"'+e.curr+'.html'"& vbCrLf
end if
if str3=8 then
ManualPagination_Tmp=ManualPagination_Tmp&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
end if
end if
ManualPagination_Tmp=ManualPagination_Tmp&"}"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&" }"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"});"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"</script>"& vbCrLf
end if
ManualPagination = ManualPagination_Tmp
end Function



Function mobileManualPagination(str1,str2)
ArticleId = str1
strContent = str2
ContentLen=strContent
CurrentPage=request("Page")
if Instr(strContent,"[zmcmsfy]")<=0 then
ManualPagination_Tmp = ManualPagination_Tmp & strContent
else
arrContent=split(strContent,"[zmcmsfy]")
pages=Ubound(arrContent)+1
pages=Ubound(arrContent)+1
if CurrentPage="" then
CurrentPage=1
else
CurrentPage=Cint(CurrentPage)
end if
if CurrentPage<1 then CurrentPage=1
ManualPagination_Tmp = ManualPagination_Tmp & arrContent(CurrentPage-1)
ManualPagination_Tmp=ManualPagination_Tmp&" <div id=""pagecontent"" class=""page""></div>"& vbCrLf 
ManualPagination_Tmp=ManualPagination_Tmp&"<script type=""text/javascript"">"& vbCrLf  
ManualPagination_Tmp=ManualPagination_Tmp&"laypage({"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "cont: 'pagecontent',"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "pages: """&pages&""","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "skin: '"&word(505)&"',"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "first: "&word(506)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "prev: "&word(507)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "next: "&word(508)&","& vbCrLf
if word(509)<>"" then
ManualPagination_Tmp=ManualPagination_Tmp&  "last: "&word(509)&","& vbCrLf
end if
ManualPagination_Tmp=ManualPagination_Tmp&  "groups: "&word(510)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&  "skip: "&word(511)&","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"curr: """&CurrentPage&""","& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"jump: function(e, first){"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"if(!first){"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"location.href = '?" & temp & "page='+e.curr"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"}"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&" }"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"});"& vbCrLf
ManualPagination_Tmp=ManualPagination_Tmp&"</script>"& vbCrLf
end if
mobileManualPagination = ManualPagination_Tmp
end Function


%>
