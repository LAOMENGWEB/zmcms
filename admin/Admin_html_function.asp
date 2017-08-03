<%
'单页生成开始
Sub Htmlabout
totalrec=Conn.Execute("select count(*) from dy where lang='"&lang&"'")(0)'获取单页的总数量
sql="Select * from dy where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof '循环生成开始
ID=rs("ID")
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")

'获取单页分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from mnclass where id = "&pone&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1

m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&aboutname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&aabout&"/index.asp","pone=",pone,"aboutID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&aboutname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&aabout&"/index.asp","pone=",pone,"aboutID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end if
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush


rs1.close
set rs1=nothing

rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing
end sub
'单页生成结束



'生成新闻分类开始 minfo高亮值
Sub Htmlnewsfl
'所有新闻生成，需要设置minfo值，否则生成不会高亮
totalrec=Conn.Execute("Select count(*) from news where lang='"&lang&"'") (0) '获取新闻的总数   
totalpage=int(totalrec/NewInfo)  '计算分页数
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&iinfo&"/","",""&NewClass&""&Separated&"1.html",""&iinfo&"/index.asp","m=",minfo,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&iinfo&"/","",""&NewClass&""&Separated&""&i&".html",""&iinfo&"/index.asp","m=",minfo,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",i)
next
end If
'所有新闻生成结束


'新闻一级分类生成开始
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from newsclass where lang='"&lang&"' order by px desc"
rs.open sql,conn,1,1
If rs.eof Then 
Class_Num=0
Else 
Class_Num=1
do while not rs.eof
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where ParentClassID = 0 and lang='"&lang&"' "
rs1.open sql1,conn,1,1
do while not rs1.eof
pone=rs1("ID")
m=rs1("m") ' 获取一级分类的高亮值
totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If

creatwjj("../"&rs1("mulu")&"/") '建立新闻分类存放的文件夹

if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&NewClass&""&Separated&""&pone&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&NewClass&""&Separated&""&pone&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",i)
next
End If

Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from newsclass where ParentClassID ="&pone&" and lang='"&lang&"'"
rs2.open sql2,conn,1,1
do while not rs2.eof
ptwo=rs2("ID")
m1=rs2("ID")  ' 获取二级分类的高亮值
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and zmtwo="&ptwo&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&NewClass&""&Separated&""&ptwo&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&NewClass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",i,"","")
next
End If

Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from newsclass where ParentClassID ="&ptwo&" and lang='"&lang&"'"
rs3.open sql3,conn,1,1
do while not rs3.eof
pthree=rs3("ID")
m2=rs3("ID") '获取三级分类的高亮值
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&NewClass&""&Separated&""&pthree&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&NewClass&""&Separated&""&pthree&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",i)
next
End If
rs3.movenext
Loop
rs3.close
set rs3=nothing
rs2.movenext
Loop
rs2.close
set rs2=nothing
rs1.movenext
Loop
rs1.close
set rs1=nothing
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
end sub
'新闻分类结束


'产品分类开始
sub htmlprofl 
totalrec=Conn.Execute("Select count(*) from product where lang='"&lang&"'") (0) '获取产品的总数   
totalpage=int(totalrec/ProductInfo)  '计算分页数
If (totalpage * ProductInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&product&"/","",""&Productclass&""&Separated&"1.html",""&product&"/index.asp","m=",mpro,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&product&"/","",""&Productclass&""&Separated&""&i&".html",""&product&"/index.asp","m=",mpro,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",i)
next
end If
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from cpclass where lang='"&lang&"' order by px desc"
rs.open sql,conn,1,1
If rs.eof Then 
Class_Num=0
Else 
Class_Num=1
do while not rs.eof
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where ParentClassID = 0 and lang='"&lang&"' "
rs1.open sql1,conn,1,1
do while not rs1.eof
pone=rs1("ID")
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from product where zmone="&pone&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/ProductInfo)       
If (totalpage * ProductInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&Productclass&""&Separated&""&pone&""&Separated&"1.html",""&product&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&Productclass&""&Separated&""&pone&""&Separated&""&i&".html",""&product&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",i)
next
End If

Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from cpclass where ParentClassID ="&pone&" and lang='"&lang&"'"
rs2.open sql2,conn,1,1
do while not rs2.eof
ptwo=rs2("ID")
m1=rs2("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from product where zmone="&pone&" and zmtwo="&ptwo&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/ProductInfo)       
If (totalpage * ProductInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&Productclass&""&Separated&""&ptwo&""&Separated&"1.html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&Productclass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",i,"","")
next
End If

Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from cpclass where ParentClassID ="&ptwo&" and lang='"&lang&"'"
rs3.open sql3,conn,1,1
do while not rs3.eof
pthree=rs3("ID")
m2=rs3("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from product where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/ProductInfo)       
If (totalpage * ProductInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&Productclass&""&Separated&""&pthree&""&Separated&"1.html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&Productclass&""&Separated&""&pthree&""&Separated&""&i&".html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",i)
next
End If
rs3.movenext
Loop
rs3.close
set rs3=nothing
rs2.movenext
Loop
rs2.close
set rs2=nothing
rs1.movenext
Loop
rs1.close
set rs1=nothing
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
end sub

'生成图片分类

sub  htmlalfl
totalrec=Conn.Execute("Select count(*) from al where lang='"&lang&"'") (0) '获取新闻的总数   
totalpage=int(totalrec/alinfo)  '计算分页数
If (totalpage * alinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&photo&"/","",""&alclass&""&Separated&"1.html",""&photo&"/index.asp","pone=","","ptwo=","","pthree=","","m=",mph,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&photo&"/","",""&alclass&""&Separated&""&i&".html",""&photo&"/index.asp","pone=","","ptwo=","","pthree=","","m=",mph,"language=",lang,"Page=",i)
next
end If
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from alclass where lang='"&lang&"' order by px desc"
rs.open sql,conn,1,1
If rs.eof Then 
	Class_Num=0
Else 
	Class_Num=1
do while not rs.eof
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from alclass where ParentClassID = 0 and lang='"&lang&"' "
rs1.open sql1,conn,1,1
do while not rs1.eof
pone=rs1("ID")
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from al where zmone="&pone&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/alinfo)       
If (totalpage * alinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&alclass&""&Separated&""&pone&""&Separated&"1.html",""&photo&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&alclass&""&Separated&""&pone&""&Separated&""&i&".html",""&photo&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",i)
next
End If

Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from alclass where ParentClassID ="&pone&" and lang='"&lang&"'"
rs2.open sql2,conn,1,1
do while not rs2.eof
ptwo=rs2("ID")
m1=rs2("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from al where zmone="&pone&" and zmtwo="&ptwo&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/alinfo)       
If (totalpage * alinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&alclass&""&Separated&""&ptwo&""&Separated&"1.html",""&photo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&alclass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&photo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",i,"","")
next
End If

Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from alclass where ParentClassID ="&ptwo&" and lang='"&lang&"'"
rs3.open sql3,conn,1,1
do while not rs3.eof
pthree=rs3("ID")
m2=rs3("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from al where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/alinfo)       
If (totalpage * alinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&alclass&""&Separated&""&pthree&""&Separated&"1.html",""&photo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&alclass&""&Separated&""&pthree&""&Separated&""&i&".html",""&photo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",i)
next
End If
rs3.movenext
Loop
rs3.close
set rs3=nothing
rs2.movenext
Loop
rs2.close
set rs2=nothing
rs1.movenext
Loop
rs1.close
set rs1=nothing
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing

end sub

'生成视频分类
sub  htmltvfl
totalrec=Conn.Execute("Select count(*) from zxsp where lang='"&lang&"'") (0) '获取新闻的总数   
totalpage=int(totalrec/tvinfo)  '计算分页数
If (totalpage * tvinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&video&"/","",""&tvclass&""&Separated&"1.html",""&video&"/index.asp","m=",mtv,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&video&"/","",""&tvclass&""&Separated&""&i&".html",""&video&"/index.asp","m=",mtv,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",i)
next
end If
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from spclass where lang='"&lang&"' order by px desc"
rs.open sql,conn,1,1
If rs.eof Then 
	Class_Num=0
Else 
	Class_Num=1
do while not rs.eof
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where ParentClassID = 0 and lang='"&lang&"' "
rs1.open sql1,conn,1,1
do while not rs1.eof
pone=rs1("ID")
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from zxsp where zmone="&pone&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/tvinfo)       
If (totalpage * tvinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&tvclass&""&Separated&""&pone&""&Separated&"1.html",""&video&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&tvclass&""&Separated&""&pone&""&Separated&""&i&".html",""&video&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",i)
next
End If

Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from spclass where ParentClassID ="&pone&" and lang='"&lang&"'"
rs2.open sql2,conn,1,1
do while not rs2.eof
ptwo=rs2("ID")
m1=rs2("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from zxsp where zmone="&pone&" and zmtwo="&ptwo&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/tvinfo)       
If (totalpage * tvinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&tvclass&""&Separated&""&ptwo&""&Separated&"1.html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&tvclass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",i,"","")
next
End If

Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from spclass where ParentClassID ="&ptwo&" and lang='"&lang&"'"
rs3.open sql3,conn,1,1
do while not rs3.eof
pthree=rs3("ID")
m2=rs3("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from zxsp where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/tvinfo)       
If (totalpage * tvinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&tvclass&""&Separated&""&pthree&""&Separated&"1.html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&tvclass&""&Separated&""&pthree&""&Separated&""&i&".html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",i)
next
End If
rs3.movenext
Loop
rs3.close
set rs3=nothing
rs2.movenext
Loop
rs2.close
set rs2=nothing
rs1.movenext
Loop
rs1.close
set rs1=nothing
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing

end sub

'生成下载分类
sub  htmlxzfl
totalrec=Conn.Execute("Select count(*) from download where lang='"&lang&"'") (0) '获取新闻的总数   
totalpage=int(totalrec/downinfo)  '计算分页数
If (totalpage * downinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&down&"/","",""&downclass&""&Separated&"1.html",""&down&"/index.asp","m=",mdo,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&down&"/","",""&downclass&""&Separated&""&i&".html",""&down&"/index.asp","m=",mdo,"pone=","","ptwo=","","pthree=","","language=",lang,"Page=",i)
next
end If
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from xzclass where lang='"&lang&"' order by px desc"
rs.open sql,conn,1,1
If rs.eof Then 
	Class_Num=0
Else 
	Class_Num=1
do while not rs.eof
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where ParentClassID = 0 and lang='"&lang&"' "
rs1.open sql1,conn,1,1
do while not rs1.eof
pone=rs1("ID")
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from download where zmone="&pone&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/downinfo)       
If (totalpage * downinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&downclass&""&Separated&""&pone&""&Separated&"1.html",""&down&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&downclass&""&Separated&""&pone&""&Separated&""&i&".html",""&down&"/index.asp","pone=",pone,"ptwo=","","pthree=","","m=",m,"language=",lang,"Page=",i)
next
End If

Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from xzclass where ParentClassID ="&pone&" and lang='"&lang&"'"
rs2.open sql2,conn,1,1
do while not rs2.eof
ptwo=rs2("ID")
m1=rs2("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from download where zmone="&pone&" and zmtwo="&ptwo&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/downinfo)       
If (totalpage * downinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&downclass&""&Separated&""&ptwo&""&Separated&"1.html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&downclass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m1,"language=",lang,"Page=",i,"","")
next
End If

Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from xzclass where ParentClassID ="&ptwo&" and lang='"&lang&"'"
rs3.open sql3,conn,1,1
do while not rs3.eof
pthree=rs3("ID")
m2=rs2("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from download where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&" and lang='"&lang&"'")(0)
totalpage=int(totalrec/downinfo)       
If (totalpage * downinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&downclass&""&Separated&""&pthree&""&Separated&"1.html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&downclass&""&Separated&""&pthree&""&Separated&""&i&".html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m2,"language=",lang,"Page=",i)
next
End If
rs3.movenext
Loop
rs3.close
set rs3=nothing
rs2.movenext
Loop
rs2.close
set rs2=nothing
rs1.movenext
Loop
rs1.close
set rs1=nothing
Response.Write "<script>bar_img.width="&Fix((Class_Num/rs.recordcount)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & formatnumber(Class_Num/rs.recordcount*100) & """;</script>"
Response.Flush
rs.movenext
Class_Num=Class_Num+1
Loop
End If
rs.close
set rs=nothing
end sub


'生成新闻开始
sub htmlnews
totalrec=Conn.Execute("select count(*) from News where lang='"&lang&"'")(0)
sql="Select * from News where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone") '一级分类
ptwo=rs("zmtwo") '二级分类
pthree=rs("zmthree") '三级分类

'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass  where id = "&pone&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

	
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&id&""&Separated&"1.html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If


rs1.close
Set rs1=nothing
End If
'一级分类结束




'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where id = "&ptwo&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

	
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&id&""&Separated&"1.html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If


rs1.close
Set rs1=nothing
End If
'二级分类结束


'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where id = "&pthree&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

	
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&id&""&Separated&"1.html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If


rs1.close
Set rs1=nothing
End If
'三级分类结束








Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush

rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing

end sub

'生成新闻结束




'生成产品开始

sub htmlpro
totalrec=Conn.Execute("select count(*) from product where lang='"&lang&"'")(0)
sql="Select * from product where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")
ptwo=rs("zmtwo")
pthree=rs("zmthree")

'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取产品分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where id = "&pone&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showproduct&"/index.asp","pone=",pone,"proID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showproduct&"/index.asp","pone=",pone,"proID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If





'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where id = "&ptwo&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showproduct&"/index.asp","pone=",pone,"proID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showproduct&"/index.asp","pone=",pone,"proID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If




'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where id = "&pthree&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showproduct&"/index.asp","pone=",pone,"proID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showproduct&"/index.asp","pone=",pone,"proID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If









Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush





rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing

end Sub

'生成产品结束


'生成图片开始
sub htmlal
totalrec=Conn.Execute("select count(*) from al where lang='"&lang&"'")(0)
sql="Select * from al where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")
ptwo=rs("zmtwo")
pthree=rs("zmthree")

'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取图片分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from alclass where id = "&pone&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&alname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showphoto&"/index.asp","pone=",pone,"phID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&alname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showphoto&"/index.asp","pone=",pone,"phID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If

'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取三级分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from alclass where id = "&ptwo&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&alname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showphoto&"/index.asp","pone=",pone,"phID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&alname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showphoto&"/index.asp","pone=",pone,"phID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If


'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取三级分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from alclass where id = "&pthree&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&alname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showphoto&"/index.asp","pone=",pone,"phID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&alname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showphoto&"/index.asp","pone=",pone,"phID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If

















Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush

rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing

end sub




'视频分类开始
sub htmltv
totalrec=Conn.Execute("select count(*) from zxsp where lang='"&lang&"'")(0)
sql="Select * from zxsp where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")
ptwo=rs("zmtwo")
pthree=rs("zmthree")


'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取视频分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where id = "&pone&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If




'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取视频分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where id = "&ptwo&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If

'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取视频分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where id = "&pthree&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If



Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush
rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing

end sub

sub htmlxz
totalrec=Conn.Execute("select count(*) from Download where lang='"&lang&"'")(0)
sql="Select * from Download where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")
ptwo=rs("zmtwo")
pthree=rs("zmthree")



'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取下载分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where id = "&pone&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If



rs1.close
Set rs1=nothing
End If



'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取下载分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where id = "&ptwo&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If



rs1.close
Set rs1=nothing
End If



'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取下载分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where id = "&pthree&" and lang='"&lang&"' "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If



rs1.close
Set rs1=nothing
End If












Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush

rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing

end sub
'生成招聘
sub htmlzp
totalrec=Conn.Execute("select count(*) from job where lang='"&lang&"'")(0)
sql="Select * from job where lang='"&lang&"' order by ID desc"
Set Rs=Conn.Execute(sql)
if totalrec=0 then
Detail_Num=0
Else
Detail_Num=1
do while not rs.eof
ID=rs("ID")
content=rs("HrDetail")
zdfy=rs("zdfy")
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&zpname&""&Separated&""&ID&""&Separated&"1.html",""&showjob&"/index.asp","jobid=",ID,"m=",mzp,"language=",lang,"page=",1,"","","","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&zpname&""&Separated&""&ID&""&Separated&""&i&".html",""&showjob&"/index.asp","jobid=",ID,"m=",mzp,"language=",lang,"page=",i,"","","","")
next
end if
Response.Write "<script>bar_img.width="&Fix((Detail_Num/totalrec)*300)&";"
Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & formatnumber(Detail_Num/totalrec*100) & """;</script>"
Response.Flush

rs.movenext
Detail_Num=Detail_Num+1
loop
end if
rs.close
set rs=Nothing

end sub

sub htmlzpfl
totalrec=Conn.Execute("Select count(*) from job where lang='"&lang&"'")(0)
totalpage=int(totalrec/zpInfo)
If (totalpage * zpInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&job&"/","",""&zpclass&""&Separated&"1.html",""&job&"/index.asp","m=",mzp,"language=",lang,"Page=",1,"","","","","","")
else
for i=1 to totalpage
call htmll("/"&job&"/","",""&zpclass&""&Separated&""&i&".html",""&job&"/index.asp","m=",mzp,"language=",lang,"Page=",i,"","","","","","")
Response.Write "<script>bar_img.width="&Fix((i/totalpage)*300)&";"
Response.Write "bar_txt1.innerHTML=""成功生成"&i&"个分类的HTML静态页面。完成比例：" & formatnumber(i/totalpage*100) & """;</script>"
Response.Flush
next
end if
rs.close
set rs=Nothing

Response.Write "<script>bar_img.width=300;bar_txt2.innerHTML='';bar_txt1.innerHTML=""成功生成相关静态文件。完成比例：100"";</script>"
Response.end

end sub
%>

