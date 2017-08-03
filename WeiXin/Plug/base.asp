<!--#include file="../../inc/conn.asp"-->
<!--#include file="../Config.asp"-->
<%

	dim Connstr2,Conn2
	dim reg
	
	
	Connstr2="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(webroot&datapath&"/"&datadbname)
	Set Conn2=server.createobject("adodb.connection")
	Conn2.open Connstr2
	
	''querystring参，t1为0时过滤空格
	public function fget(byval t0,byval t1)
		fget=request.querystring(t0)
		if t1=0 then fget=trim(fget)
	end function	
	
	''转换HTML代码，过滤代码
	public function enhtml(byval t0)
		if isnull(t0) then enhtml="":exit function
		if t0="<p>&nbsp;</p>" then enhtml="":exit function
		t0=replace(t0,"&","&amp;")
		t0=replace(t0,"'","&#39;")
		t0=replace(t0,"""","&#34;")
		t0=replace(t0,"<","&lt;")
		t0=replace(t0,">","&gt;")
		set reg=new regexp
		reg.ignorecase=true
		reg.global=true
		reg.pattern="(w)(here)"
		t0=reg.replace(t0,"$1h&#101;re")
		reg.pattern="(s)(elect)"
		t0=reg.replace(t0,"$1el&#101;ct")
		reg.pattern="(i)(nsert)"
		t0=reg.replace(t0,"$1ns&#101;rt")
		reg.pattern="(c)(reate)"
		t0=reg.replace(t0,"$1r&#101;ate")
		reg.pattern="(d)(rop)"
		t0=reg.replace(t0,"$1ro&#112;")
		reg.pattern="(a)(lter)"
		t0=reg.replace(t0,"$1lt&#101;r")
		reg.pattern="(d)(elete)"
		t0=reg.replace(t0,"$1el&#101;te")
		reg.pattern="(u)(pdate)"
		t0=reg.replace(t0,"$1p&#100;ate")
		reg.pattern="(\s)(or)"
		t0=reg.replace(t0,"$1o&#114;")
		reg.pattern="(java)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		reg.pattern="(j)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		reg.pattern="(vb)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		if instr(t0,"expression")<>0 then
			t0=replace(t0,"expression","e&#173;xpression",1,-1,0)
		end if
		enhtml=t0
	end function
		
	''逆向转换HTML
	public function dehtml(byval t0)
		if isnull(t0) then
			dehtml=""
			exit function
		end if
		t0=replace(t0,"&amp;","&")
		t0=replace(t0,"&#39;","'")
		t0=replace(t0,"&#34;","""")
		t0=replace(t0,"&lt;","<")
		t0=replace(t0,"&gt;",">")
		t0=replace(t0,chr(10),vbcrlf)
		dehtml=t0
	end function
	
	''逆向转换HTML
	public function deurl(byval t0)
		if isnull(t0) then
			deurl=""
			exit function
		end if
		t0=replace(t0,"%2F","/")
		t0=replace(t0,"%2E",".")
		deurl=t0
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
CutStrX=Left(Str,i)
Exit For
Else
CutStrX=Str
End If
Next
End Function
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
function getPicUrl(str)
    dim content,regstr,url
    content=str&""
    regstr="src=.+?.(gif|jpg|jpeg|jpe|png|bmp)"
    url=Replace(Replace(Replace(RegExp_Execute(regstr,content),"'",""),"""",""),"src=","")
    getPicUrl=url
end function
sub DeleteFile(path)
path = server.mappath(path)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then fso.DeleteFile(path)
set fso=nothing
end sub

sub DeleteFile1(path)
path = server.mappath("/" & path)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then fso.DeleteFile(path)
set fso=nothing
end sub
%>