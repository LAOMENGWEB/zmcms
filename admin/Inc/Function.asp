<%
'*************************************************
'函数名：GetUrl
'作  用：获取网址
'*************************************************
Function GetUrl()
GetUrl="http://"&Request.ServerVariables("SERVER_NAME")
If Request.ServerVariables("QUERY_STRING")<>"" Then GetURL=GetUrl&"?"& Request.ServerVariables("QUERY_STRING")
End Function
'*************************************************
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串
'*************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "…"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'**************************************************
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
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

'***************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'***************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
Function zmcmsRequest(ParaName)
     Dim ParaValue
     ParaValue = trim(ParaName)
	 If ParaValue="" Then Exit Function
        If Not isNumeric(ParaValue) Then
		    Response.Write "<li>  参数类型不合法"
			Response.end
	    Else
           zmcmsRequest = ParaValue
        End If
End Function
Function CheckStr(Strer,Num)
        Dim Shield,w
	    If Strer = "" Or IsNull(Strer) Then Exit Function
	    Select Case Num
		  Case 1
	        If IsNumeric(Strer) = 0 Then
	          Response.Write "操作错误"
		      Response.End
	        End If
			Strer = Int(Strer)
		End Select
	    CheckStr = Strer
   End Function

Function oflink_lcasetag(imgstrng)
'将内容html代码里的标签大写转换成小写
Dim regEx, Match, Matches '建立变量。 
Set regEx = New RegExp '建立正则表达式。 
regEx.Pattern = "<.+?\>" '设置模式。 
regEx.IgnoreCase = true '设置是否区分字符大小写。 
regEx.Global = True '设置全局可用性。 
Set Matches = regEx.Execute(imgstrng) '执行搜索。 
imgstrng=imgstrng
For Each Match in Matches '遍历匹配集合。 
imgstrng=replace(imgstrng,Match.Value,lcase(Match.Value))
Next
oflink_lcasetag = imgstrng 
End Function
Function ofink_getsrc(strng) 
strng=oflink_lcasetag(strng)
Dim regEx, Match, Matches,values,if_add '建立变量。 
Set regEx = New RegExp '建立正则表达式。 
regEx.Pattern = "src\=.+?\.(gif|jpg|jpeg|jpe|png)" 
regEx.IgnoreCase = true '设置是否区分字符大小写。 
regEx.Global = True '设置全局可用性。 
Set Matches = regEx.Execute(strng) '执行搜索。 
if_add=0
For Each Match in Matches '遍历匹配集合。 
values=values&Match.Value&"|" 
if_add=1
Next 
values=replace(values,"src=""","")
values=replace(values,"src='","")
values=replace(values,"src=","")
if if_add=1 then values=left(values,len(values)-1)
ofink_getsrc = values 
End Function 

'***********************************************
'函数名：getPicUrl
'作  用：获得信息里的图片地址
'参  数：str  ----信息
'***********************************************
function getPicUrl(str)
    dim content,regstr,url
    content=str&""
    regstr="src=.+?.(gif|jpg|jpeg|jpe|png|bmp)"
    url=Replace(Replace(Replace(RegExp_Execute(regstr,content),"'",""),"""",""),"src=","")
    getPicUrl=url
end function

Function RegExp_Execute(patrn, strng)
Dim regEx, Match, Matches,values '建立变量。
Set regEx = New RegExp '建立正则表达式。
regEx.Pattern = patrn '设置模式。
regEx.IgnoreCase = true '设置是否区分字符大小写。
regEx.Global = True '设置全局可用性。
Set Matches = regEx.Execute(strng) '执行搜索。
For Each Match in Matches '遍历匹配集合。
values=values&Match.Value&","
Next
RegExp_Execute = values
End Function

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


Function ubb(strer)
  strer=strer
strer=replace(strer,",","")
  strer=replace(strer,">","")
  strer=replace(strer,CHR(32),"")    '空格
  strer=replace(strer,CHR(32),"")  
  strer=replace(strer,CHR(9),"")    'table
  strer=replace(strer,CHR(39),"")    '单引号
  strer=replace(strer,CHR(34),"")    '双引号
  strer = Replace(strer, CHR(13), "")
  strer = Replace(strer, CHR(10), "")
  strer = Replace(strer, "[enter]", "")
    strer=replace(strer,vbCrLf,"")
ubb=strer
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
function ReStrReplace(Str)'写入表单替换字符
  if Str="" or isnull(Str) then 
    ReStrReplace=""
    exit function 
  else
    ReStrReplace=replace(Str,"&nbsp;"," ") '"&nbsp;"
    ReStrReplace=replace(ReStrReplace,"<br>",chr(13))'"<br>"
    ReStrReplace=replace(ReStrReplace,"&lt;br&gt;",chr(13))'"<br>"
    ReStrReplace=replace(ReStrReplace,"&lt;","<")' "&lt;"
    ReStrReplace=replace(ReStrReplace,"&gt;",">")' "&gt;"
  end if
end function
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

Function IsFileExist(filespec)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      IsFileExist=true
   Else
      IsFileExist=false
   End If
   set fso=nothing
End Function


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



Dim App_DataCount,App_Language,LanguageName(10),LanguageTxt(10),LanguagePath(10),LanguageFile(10),LanguageMoney(10),LanguageRate(10),LanguageSkin(10),LanguageMark(10),LanguageID(10),LicenseData,ArrData
TotalLang = 1
Function tmp(a,b)
Dim tmp_i
For tmp_i = 1 To a
tmp = tmp & "&nbsp;|&nbsp;&nbsp;"
Next
tmp = tmp & b
End Function
Function Safereplace(Str)
If Isnull(Str) Then
Safereplace = ""
Exit Function
End If
Str = Replace(Trim(Str),"'","&#180;")
Str = Replace(Str,"%","")
Str = Replace(Str,"""","&#34;")
Str = Replace(Str,"<","&lt;")
Str = Replace(Str,">","&gt;")
Str = Replace(Str,"$","")
Safereplace = Str
End Function
Function NoEnter(Str)
If Isnull(Str) Then
NoEnter = ""
Exit Function
End If
Str = Replace(Str,Chr(8),"")
Str = Replace(Str,Chr(9),"")
Str = Replace(Str,Chr(10),"")
Str = Replace(Str,Chr(13),"")
NoEnter = Str
End Function
Function CheckNum(num,back)
num = Trim(num)
If Not IsNumeric(num) Or IsEmpty(num) Then
CheckNum = CLng(back)
Else
CheckNum = num
End If
End Function

Function change(Str)
	    Str = Replace(Str,"{","<%=")
	    Str = Replace(Str,"]","%")
		Str = Replace(Str,"}",">")
	    change = Str
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
Function sortdate(str,str1)
       If str="" Then
	        Exit Function
	   End If
	   news_year=year(str)
       news_month=month(str)
       If Len(news_month)=1 Then news_month="0"&news_month
       news_day=day(str)
       If Len(news_day)=1 Then news_day="0"&news_day
	   If str1=0 Then
       sortdate=CLng(news_month&news_day)
	   Else
	   sortdate=CLng(news_year&news_month&news_day)
	   End If
End Function
function formatnumber2(num)
        If IsNull(num) Or Not IsNumeric(num) Then
		        Exit Function
	    End If
	    num = formatnumber(num,2)
	    if num <1 And num > 0 then
	            num= 0&num
	            num = replace(num,"00.","0.")
		ElseIf num = 0 Then
				num= 0&num
	    end if
	    formatnumber2 = Replace(num,",","")
end Function

function JoinChar(strUrl)
		if strUrl="" then
				JoinChar=""
				exit function
		end if
		if InStr(strUrl,"?")<len(strUrl) then
				if InStr(strUrl,"?")>1 then
						if InStr(strUrl,"&")<len(strUrl) then
								JoinChar=strUrl & "&"
						else
								JoinChar=strUrl
						end if
				else
						JoinChar=strUrl & "?"
				end if
		else
				JoinChar=strUrl
		end if
end function
Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
End Function

Function CutText(str,stralign,cutnum)
		On Error Resume Next 
		If str = "" Or IsNull(str) Then Exit Function
		Select Case stralign
			   Case "0"
					CutText = left(str,Len(str)-cutnum)
			   Case "1"
					CutText = right(str,Len(str)-cutnum)
		End Select
End Function
Function Savehtml(filename,filebody)
		On Error Resume Next
        filename = server.mappath(filename)
		filebody = Replace(filebody,"&lt;","<")
		filebody = Replace(filebody,"&gt;",">")
		filebody = Replace(filebody,"&quot;","""")
		filebody = Replace(filebody,"&amp;","&")
		filebody = Replace(filebody,"&middot;","·")
        adTypeText=2
        adSaveCreateOverWrite=2
        Set objStream = Server.CreateObject("ADODB.Stream")
        With objStream
                .Type = adTypeText
                .Mode = adModeReadWrite
                .Open
                .Charset = "utf-8"
                .Position = objStream.Size
                .WriteText= filebody
                .SaveToFile FileName,adSaveCreateOverWrite
                .Close
        End With
End Function
Function SavehtmlGb2312(filename,filebody)
		On Error Resume Next
        filename = server.mappath(filename)
        adTypeText=2
        adSaveCreateOverWrite=2
        Set objStream = Server.CreateObject("ADODB.Stream")
        With objStream
                .Type = adTypeText
                .Mode = adModeReadWrite
                .Open
                .Charset = "Gb2312"
                .Position = objStream.Size
                .WriteText= filebody
                .SaveToFile FileName,adSaveCreateOverWrite
                .Close
        End With
End Function
Function LoadFile(strFile)
		On Error Resume Next
		Dim strTempPath,objFSO,objStm
		strTempPath = Server.MapPath(strFile)
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set objStm = Server.CreateObject("ADODB.Stream")
			If objFso.FileExists(strTempPath) Then
				objStm.charset = "UTF-8"
				objStm.open
				objStm.LoadFromFile strTempPath
				LoadFile = objStm.ReadText
			End If
		Set objStm = Nothing
		Set objFSO = Nothing
		If Err Then
			    Err.Clear
		End If
End Function
Function AdminLoadStyle(strFile)
		Dim strTempPath,objFSO,objStm
		strTempPath = Server.MapPath(strFile)
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set objStm = Server.CreateObject("ADODB.Stream")
			If objFso.FileExists(strTempPath) Then
				objStm.charset = "UTF-8"
				objStm.open
				objStm.LoadFromFile strTempPath
				AdminLoadStyle = objStm.ReadText
			End If
		Set objStm = Nothing
		Set objFSO = Nothing
End Function

Function CheckDir(byval FolderPath)
        dim fso
		If IsNull(FolderPath) Or InStr(FolderPath,"//") > 0 Then 
				CheckDir = False 
				Exit Function 
		End If 
        folderpath=Server.MapPath(folderpath)
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(FolderPath) then
                CheckDir = True
        Else
                CheckDir = False
        End if
        Set fso = nothing
End Function
Function CheckDir2(byval FolderPath)
        dim fso
        folderpath=Server.MapPath(".")&"\"&folderpath
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(FolderPath) then
                CheckDir2 = True
        Else
                CheckDir2 = False
        End if
        Set fso = nothing
End Function
Function MakeNewsDir2(byval foldername)
        dim fso
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder(Server.MapPath(".") &"\" &foldername)
        If fso.FolderExists(Server.MapPath(".") &"\" &foldername) Then
                MakeNewsDir2 = True
        Else
                MakeNewsDir2 = False
        End If
        Set fso = nothing
End Function
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
        Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
        If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" Then
                DefiniteUrl="$False$"
                Exit Function
        End If
        If Left(ConsultUrl,7)<>"HTTP://" And Left(ConsultUrl,7)<>"http://" Then
                ConsultUrl= "http://" & ConsultUrl
        End If
        ConsultUrl=Replace(ConsultUrl,"://",":\\")
        If Right(ConsultUrl,1)<>"/" Then
                If Instr(ConsultUrl,"/")>0 Then
                        If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then
                        Else
                                ConsultUrl=ConsultUrl & "/"
                        End If
                Else
                        ConsultUrl=ConsultUrl & "/"
                End If
        End If
        ConArray=Split(ConsultUrl,"/")
        If Left(PrimitiveUrl,7) = "http://" then
                DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
        ElseIf Left(PrimitiveUrl,1) = "/" Then
                DefiniteUrl=ConArray(0) & PrimitiveUrl
        ElseIf Left(PrimitiveUrl,2)="./" Then
                DefiniteUrl=ConArray(0) & Right(PrimitiveUrl,Len(PrimitiveUrl)-1)
        ElseIf Left(PrimitiveUrl,3)="../" then
                Do While Left(PrimitiveUrl,3)="../"
                PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
                Pi=Pi+1
                Loop
                For Ci=0 to (Ubound(ConArray)-1-Pi)
                        If DefiniteUrl<>"" Then
                                DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
                        Else
                                DefiniteUrl=ConArray(Ci)
                        End If
                Next
                DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
        Else
                If Instr(PrimitiveUrl,"/")>0 Then
                        PriArray=Split(PrimitiveUrl,"/")
                        If Instr(PriArray(0),".")>0 Then
                                If Right(PrimitiveUrl,1)="/" Then
                                        DefiniteUrl="http:\\" & PrimitiveUrl
                                Else
                                        If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then
                                                DefiniteUrl="http:\\" & PrimitiveUrl
                                        Else
                                                DefiniteUrl="http:\\" & PrimitiveUrl & "/"
                                        End If
                                End If
                        Else
                                If Right(ConsultUrl,1)="/" Then
                                        DefiniteUrl=ConsultUrl & PrimitiveUrl
                                Else
                                        DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
                                End If
                        End If
                Else
                        If Instr(PrimitiveUrl,".")>0 Then
                                If Right(ConsultUrl,1)="/" Then
                                        If right(PrimitiveUrl,3)=".cn" or right(PrimitiveUrl,3)="com" or right(PrimitiveUrl,3)="net" or right(PrimitiveUrl,3)="org" Then
                                                DefiniteUrl="http:\\" & PrimitiveUrl & "/"
                                        Else
                                                DefiniteUrl=ConsultUrl & PrimitiveUrl
                                        End If
                                Else
                                        If right(PrimitiveUrl,3)=".cn" or right(PrimitiveUrl,3)="com" or right(PrimitiveUrl,3)="net" or right(PrimitiveUrl,3)="org" Then
                                                DefiniteUrl="http:\\" & PrimitiveUrl & "/"
                                        Else
                                                DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
                                        End If
                                End If
                        Else
                                If Right(ConsultUrl,1)="/" Then
                                        DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
                                Else
                                        DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
                                End If
                        End If
                End If
        End If
        If Left(DefiniteUrl,1)="/" then
                DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
        End if
        If DefiniteUrl<>"" Then
                DefiniteUrl=Replace(DefiniteUrl,"//","/")
                DefiniteUrl=Replace(DefiniteUrl,":\\","://")
        Else
                DefiniteUrl="$False$"
        End If
End Function
Function ReplaceSaveRemoteFile(ConStr,IncluL,IncluR,SaveTf,SaveFilePath,ThisUrl)
		On Error Resume Next 
'        If ThisUrl = "/" Or ThisUrl = "" Then
'				If request.ServerVariables("SERVER_PORT") <> 80 Then
'						ThisUrl="http://"&Request.ServerVariables("SERVER_NAME")&":"&request.ServerVariables("SERVER_PORT")&"/"
'				Else
'						ThisUrl="http://"&Request.ServerVariables("SERVER_NAME")&"/"
'				End If
'		End If
		If request.ServerVariables("SERVER_PORT") <> 80 Then
				WebUrl = "http://"&Request.ServerVariables("SERVER_NAME")&":"&request.ServerVariables("SERVER_PORT")&"/"
		Else
				WebUrl = "http://"&Request.ServerVariables("SERVER_NAME")&"/"
		End If
        If ConStr="$False$" or ConStr="" or IsNull(ConStr) Then
                ReplaceSaveRemoteFile="$False$"
                Exit Function
        End If
        Dim TempStr,TempStr2,ReF,Matches,Match,Tempi,TempArray,TempArray2,OverTypeArray,StartStr,OverStr
        '图片开始的字符串
        StartStr="src="
        '图片结束的字符串
        OverStr="gif|jpg|bmp|jpeg|png"
        Set ReF = New Regexp
        ReF.IgnoreCase = True
        ReF.Global = True
        ReF.Pattern = "("&StartStr&").+?("&OverStr&")"
        Set Matches =ReF.Execute(ConStr)
        For Each Match in Matches
                If Instr(TempStr,Match.Value)=0 Then
                        If TempStr<>"" then
                                TempStr=TempStr & "$Array$" & Match.Value
                        Else
                                TempStr=Match.Value
                        End if
                End If
        Next
        Set Matches=nothing
        Set ReF=nothing
        If TempStr="" or IsNull(TempStr)=True Then
                ReplaceSaveRemoteFile=ConStr
                Exit function
        End if
        If IncluL=False then
                TempStr=Replace(TempStr,StartStr,"")
        End if
        If IncluR=False then
                If Instr(OverStr,"|")>0 Then
                        OverTypeArray=Split(OverStr,"|")
                        For Tempi=0 To Ubound(OverTypeArray)
                                TempStr=Replace(TempStr,OverTypeArray(Tempi),"")
                        Next
                Else
                        TempStr=Replace(TempStr,OverStr,"")
                End If
        End if
        TempStr=Replace(TempStr,"""","")
        TempStr=Replace(TempStr,"'","")
        Dim RemoteFile,RemoteFileurl,SaveLocalFileName,SaveFileType,ArrSaveLocalFileName,RanNum
        If Right(SaveFilePath,1)="/" then
                SaveFilePath=Left(SaveFilePath,Len(SaveFilePath)-1)
        End If
        If SaveTf=True then
                If CheckDir2(SaveFilePath)=False Then
                        If MakeNewsDir2(SaveFilePath)=False Then
                                SaveTf=False
                        End If
                End If
        End If
        SaveFilePath= RemoteFileCreatePath(SaveFilePath)
        '图片转换/保存
        TempArray=Split(TempStr,"$Array$")
        For Tempi=0 To Ubound(TempArray)
                RemoteFileurl=DefiniteUrl(TempArray(Tempi),WebUrl)
				If InStr(RemoteFileurl,WebUrl)=0 Then
                If RemoteFileurl<>"$False$" And SaveTf=True Then'保存图片
                        ArrSaveLocalFileName = Split(RemoteFileurl,".")
                        SaveFileType=ArrSaveLocalFileName(Ubound(ArrSaveLocalFileName))'文件类型
                        RanNum=Int(900*Rnd)+100
                        SaveLocalFileName = SaveFilePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&SaveFileType
						dim Ads,Retrieval,GetRemoteData
						Set Retrieval = CreateObject("Micro"&"soft"&"."&"XML"&"HTTP")
						With Retrieval
						.Open "Get", RemoteFileurl, False, "", ""
						.Send
						GetRemoteData = .ResponseBody
						End With
						Set Retrieval = Nothing
						Set Ads = Server.CreateObject("ado"&"db"&"."&"str"&"eam")
						With Ads
						.Type = 1
						.Open
						.Write GetRemoteData
						.SaveToFile server.MapPath(SaveLocalFileName),2
						.Cancel()
						.Close()
						End With
						Set Ads=nothing
						ConStr=Replace(ConStr,TempArray(Tempi),ThisUrl &Replace(SaveLocalFileName,"../",""))
                ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'不保存图片
                        SaveLocalFileName=RemoteFileUrl
                        ConStr=Replace(ConStr,TempArray(Tempi),SaveLocalFileName)
                End If
                If RemoteFileUrl<>"$False$" Then
                        If UploadFiles="" then
                                UploadFiles=SaveLocalFileName
                        Else
                                UploadFiles=UploadFiles & "|" & SaveLocalFileName
                        End if
                End If
				End If
        Next
        ReplaceSaveRemoteFile=ConStr
End Function
Function CreatePath(PathValue)
		Dim objFSO,Fsofolder,uploadpath
		yy=year(Date()):mm=month(Date())
		if len(mm)=1 then mm="0"&mm
		uploadpath = yy&mm
		If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
		On Error Resume Next
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
				objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
		End If
		If Err.Number = 0 Then
				CreatePath = PathValue & uploadpath & "/"
		End If
		Set objFSO = Nothing
End Function
Function RemoteFileCreatePath(PathValue)
		Dim objFSO,Fsofolder,uploadpath
		yy=year(Date()):mm=month(Date())
		if len(mm)=1 then mm="0"&mm
		uploadpath = yy&mm
		If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
		On Error Resume Next
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
				objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
		End If
		If Err.Number = 0 Then
				RemoteFileCreatePath = PathValue & uploadpath & "/"
		End If
		Set objFSO = Nothing
End Function
sub SaveRemoteFile(LocalFileName,RemoteFileUrl)
		On Error Resume Next
        dim Ads,Retrieval,GetRemoteData
        Set Retrieval = CreateObject("Micro"&"soft"&"."&"XML"&"HTTP")
        With Retrieval
        .Open "Get", RemoteFileUrl, False, "", ""
        .Send
        GetRemoteData = .ResponseBody
        End With
        Set Retrieval = Nothing
        Set Ads = Server.CreateObject("ado"&"db"&"."&"str"&"eam")
        With Ads
        .Type = 1
        .Open
        .Write GetRemoteData
        .SaveToFile server.MapPath(LocalFileName),2
        .Cancel()
        .Close()
        End With
        Set Ads=nothing
end sub
Function GetImg(str,strpath)
        set objregEx = new RegExp
        objregEx.IgnoreCase = true
        objregEx.Global = true
        objregEx.Pattern = "src=(.+?)\.(jpg|gif|png|bmp)"
        set matches = objregEx.execute(str)
        for each match in matches
                sImgone = sImgone & "|"& Match.Value
        next
        if sImgone<>"" then
                Imglist=split(sImgone,"|")(1)
'                Imgone=replace(Imgone,"src='","")
				Imgone=replace(Imglist,"src=""","")
				If Left(Imgone,1) = Eshop_url Then Imgone = Right(Imgone,Len(Imgone)-Len(Eshop_url))
				GetImg=Imgone
'				If Left(Imgone,"1") = "/" Then
'					GetImg=Imgone
'				Else
'					GetImg=strpath & Imgone
'				End If
        else
                GetImg=""
        end if
end function
Function getPicNewSize(sStr,pw,ph)
		Dim oldpw,oldph,pw_r,ph_r,newpw,newph
		newpw = 0
		newph = 0
		If instr(sStr,"|")>0 Then
				oldpw = CDbl(split(sStr,"|")(0))
				oldph = CDbl(split(sStr,"|")(1))
				pw_r = pw/oldpw
				ph_r = ph/oldph
				If pw_r > ph_r Then
						newpw = oldpw * ph_r
						newph = oldph * ph_r
				Else
						newpw = oldpw * pw_r
						newph = oldph * pw_r
				End If
		Else
				newpw = pw
				newph = ph
		End If
		 getPicNewSize = CInt(newpw)&"|"&CInt(newph)
End Function
rem	<-{ 数据下拉菜单
Function dbHTMLSelect(mySql,showStr,valueStr,selectedStr)
		dim valStr,seledStr,optionStr,myRsSelect
		set myRsSelect = conn.execute(mysql)
		while not myRsSelect.eof
				if valueStr <> "" then
						valStr = "value="""&myRsSelect(valueStr)&""""
						if selectedStr <> "" and CStr(selectedStr)=CStr(myRsSelect(valueStr)) then
								seledStr = "selected"
						end if
				end if
				optionStr = optionStr & " <option "&valStr&" "&seledStr&">"&myRsSelect(showStr)&"</option> "
				valStr = ""
				seledStr = ""
				myRsSelect.movenext
		wend
		myRsSelect.close:set myRsSelect = nothing
		dbHTMLSelect = optionStr
End Function

Sub OptionClass(a,b,c)
		sql="select classid,classname From [Lebi_newsclass] where ParentID  = "& a &" order by B_order desc"
		Set Rs1=Conn.Execute(sql)
		do while not rs1.eof
		If a = 0 Then
				Response.write("<option value='"& clng(rs1(0)) &"'")
		if InStr(c,","&rs1(0)&",")>0 Then Response.write("Selected")
				Response.write(">+ "& rs1(1) &"</option>")
		Else
				Response.write("<option value='"& clng(rs1(0)) &"'")
		If InStr(c,","&rs1(0)&",")>0 Then Response.write("Selected")
				Response.write(">"& tmp(b,"├ ") &" "& rs1(1) &"</option>")
		End If
		b=b+1
		OptionClass rs1(0),b,c
		b=b-1
		rs1.movenext
		loop
		rs1.Close : Set Rs1 = Nothing
End Sub

Sub Debug(str,stype)
		Response.Write(str)
		If stype = 1 Then Response.End
End Sub

Function Splits(strBlock)
		If strBlock = "" Or IsNull(strBlock) Then Exit Function
		Dim strContent,strValue,strValuestr
		strContent = Replace(strBlock,"[[","'")
		strContent = Replace(strContent,"&#180;","'")
        Dim p_RegExp
        Set p_RegExp = New RegExp
		Dim objRegExp,objValue

		p_RegExp.ignorecase = True
		p_RegExp.global = True
		p_RegExp.pattern = "{\[(.+?)\]}"
		p_RegExp.pattern = "(<(script|link|img|input|embed|param|object|base|area|map|table|td|a|form).+?(href|src|background|action)=.+?)(.+?\>)"

		Set objRegExp = p_RegExp.Execute(strContent)
		For Each objValue In objRegExp
			If instr(LCase(objValue.value),"http") = 0 And instr(LCase(objValue.value),"msnim:") = 0 And instr(LCase(objValue.value),"tencent:") = 0 And instr(LCase(objValue.value),"skype:") = 0 And instr(LCase(objValue.value),"wangwang:") = 0 Then
			        If ESHOP_URL = "/" Then
							strValue	= p_RegExp.Replace(objValue.value,"$1" & ESHOP_URL & "$4")
					Else
							strValue	= p_RegExp.Replace(Replace(objValue.value,ESHOP_URL,""),"$1" & ESHOP_URL & "$4")
					End If
					strContent = Replace(strContent,objValue.value,strValue)
			End If
		Next
		Set objRegExp = Nothing
		Splits = strContent
End Function
Function TplShowReplace(str)
		TplShowReplace = Replace(str,"[[","'")
		TplShowReplace = Replace(TplShowReplace,"&#180;","'")
End Function

Function GetMaxNum(Arr_data)
        MaxNum = 0
		If Arr_data = "" Then Exit Function
		Arr = Split(Left(Arr_data,Len(Arr_data)-1),",")
        For i = 0 To ubound(Arr)
		        If CLng(Arr(i)) > CLng(MaxNum) Then
						MaxNum = Arr(i)
				End If
		Next
		GetMaxNum = MaxNum
End Function

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




Function bytes2BSTR(vIn)
		dim strReturn
		dim i,ThisCharCode ,NextCharCode
		strReturn = ""
		For i = 1 To LenB(vIn)
				ThisCharCode = AscB(MidB(vIn,i,1))
				If ThisCharCode < &H80 Then
						strReturn = strReturn & Chr(ThisCharCode )
				Else
						NextCharCode = AscB(MidB(vIn,i+1,1))
						strReturn = strReturn & Chr(CLng(ThisCharCode ) * &H100 + CInt(NextCharCode ))
						i = i + 1
				End If
		Next
		bytes2BSTR = strReturn
End Function
Function GetHttp(U)
		On Error Resume Next
		If Lcase(Left(U,7))<>"http://" Then Exit Function
		Set D=Server.CreateObject("Microsoft.XMLHTTP"):D.Open "GET",U,false:D.send
		If D.status=200 Then GetHttp=D.ResponseBody
		Set D=Nothing
End Function
Function SaveAs(B,P)
		On Error Resume Next
		Set D=server.createobject("ADODB.Stream"):D.type=1:D.Mode=3:D.Open:D.Write B
		If D.size>0 Then D.SaveToFile P,2:ErrN=0:
		D.Close:Set D=nothing
End Function
Function UnPack(file)
	   On Error Resume Next 
	   str = Server.MapPath("../") &"\"
	   Set rs = CreateObject("ADODB.RecordSet")
	   Set stream = CreateObject("ADODB.Stream")
	   Set conn = CreateObject("ADODB.Connection")
	   connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../"& file)
	   conn.Open connStr
	   rs.Open "FileData", conn, 1, 1
	   stream.Open
	   stream.Type = 1
	   UnPack = 0
	   Do Until rs.Eof
				If InStr(LCase(rs("P")),"update\default.asp") > 0 Then
						UnPack = UnPack + 1
				End If
				theFolder = Left(rs("P"), InStrRev(rs("P"), "\"))
				If AutoCreateFolder(str & theFolder) Then
'						 response.write str & theFolder
'						 oFso.CreateFolder(str & theFolder)
				End If
				stream.SetEOS()
				If IsNull(rs("fileContent")) = False Then stream.Write rs("fileContent")
				stream.SaveToFile str & rs("P"), 2
				rs.MoveNext
	   Loop
	   rs.Close
	   conn.Close
	   stream.Close
	   Set ws = Nothing
	   Set rs = Nothing
	   Set stream = Nothing
	   Set conn = Nothing
End Function
Function AutoCreateFolder(strPath) 'As Boolean 
        On Error Resume Next 
        Dim astrPath, ulngPath, i, strTmpPath 
        Dim objFSO 
        If InStr(strPath, "\") <=0 Or InStr(strPath, ":") <= 0 Then 
                AutoCreateFolder = False 
                Exit Function 
        End If 
        Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
        If objFSO.FolderExists(strPath) Then 
                AutoCreateFolder = True 
                Exit Function 
        End If 
        astrPath = Split(strPath, "\") 
        ulngPath = UBound(astrPath) 
        strTmpPath = "" 
        For i = 0 To ulngPath 
                strTmpPath = strTmpPath & astrPath(i) & "\" 
                If Not objFSO.FolderExists(strTmpPath) Then 
                        '创建 
                        objFSO.CreateFolder(strTmpPath) 
                End If 
        Next 
        Set objFSO = Nothing 
        If Err = 0 Then 
                AutoCreateFolder = True 
        Else 
                AutoCreateFolder = False 
        End If 
End Function 
Function ReplaceFolder(Str)
	    If Isnull(Str) Then
	    	    Safereplace = ""
	    	    Exit Function
	    End If
	    Str = Replace(Trim(Str),"*","")
	    Str = Replace(Str,"?","")
	    Str = Replace(Str,"""","")
        Str = Replace(Str,"\","")
        Str = Replace(Str,"/","")
	    Str = Replace(Str,"|","")
        Str = Replace(Str,":","")
	    Str = Replace(Str,"<","")
	    Str = Replace(Str,">","")
	    ReplaceFolder = Str
End Function
Function copyfile(startfile,tofile)
		Set MyFileObject=Server.CreateObject("Scripting.FileSystemObject")
		If MyFileObject.FolderExists(startfile) Or MyFileObject.FolderExists(tofile) Then  
				Exit Function 
		End If 
		Set MyFolder=MyFileObject.GetFolder(startfile)
		domain=Split(startfile,"\")(UBound(Split(startfile,"\")))
		For Each thing in MyFolder.Files
				s=Split(thing,"\")
				a=UBound(s)
				s3=Split(thing,"\")(a)
				MyFileObject.CopyFile thing,tofile&"\"&s3
		Next
		For Each thing in MyFolder.SubFolders
				s=Split(thing,"\")
				a=UBound(s)
				s3=Split(thing,"\")(a)
				MyFileObject.copyFolder thing,tofile&"\"&s3
		Next
End Function 
Function getN(str,keyword,splits) 
		getN = 0 
		arr = Split(str,splits) 
		getN = UBound(arr) 
		For n = 0 To UBound(arr) 
		if CStr(arr(n)) = CStr(keyword) then 
				getN = (n-1) + 1 
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
		Response.Write "<p align='center'><font color='red'><b>服务器获取文件内容出错，请稍后再试！</b></font></p>" 
		Err.Clear
	Else
		GetHTTPPage = GetHTTPPage &"<span style=""font-size:12px"">查询数据由：快递100(<a href='http://kuaidi100.com' target='_blank'>www.kuaidi100</a>) 提供</span>"
	End If
End Function
Function BytesToBstr(body) 
	On Error Resume Next 
	If IsNull(body) Then 
		BytesToBstr = ""
		Exit Function 
	End If 
    Dim objstream 
    Set objstream = Server.CreateObject("Adodb.Stream") '//调用adodb.stream组件
    objstream.Type = 1 
    objstream.Mode =3 
    objstream.Open 
    objstream.Write body 
    objstream.Position = 0 
    objstream.Type = 2 
    objstream.Charset = "UTF-8" '转换原来默认的编码转换成GB2312编码，否则直接用XMLHTTP调用有中文字符的网页得到的将是乱码 
    BytesToBstr = objstream.ReadText 
    objstream.Close 
    Set objstream = Nothing 
End Function
Function CheckUrl(url)
	If InStr(url,"http") = 0 Then
		CheckUrl = url
	Else
		If InStr(url,"http://"& Request.ServerVariables("SERVER_NAME")) > 0 Then 
			CheckUrl = url
		Else
			CheckUrl = ""
			Exit Function 
		End If 
	End If 
End Function 
Function MoveSave(Rstr)
	Dim i,SpStr
	SpStr = Split(Rstr,"|")
	For i = 0 To Ubound(Spstr)
		If instr("|"& MoveSave &"|","|"& SpStr(i) &"|") = 0 Then
			MoveSave = MoveSave & SpStr(i)
			If CLng(i) < CLng(Ubound(Spstr)) And CLng(Ubound(Spstr)) > 0 Then MoveSave = MoveSave & "|"
		End If
	Next
	If Right(MoveSave,1) = "|" Then MoveSave = CutText(MoveSave,0,1)
End Function
Function Urlencode(str)
		If IsNull(str) Then
				Exit Function
		End If
		Urlencode = server.urlencode(str)
End Function
rem 去掉最后一位
Public Function cut_Right(str)
        If str="" Then
		       Exit Function
	    End If
        cut_Right=Left(str,(Len(str)-1))
End Function

function creatwjj(str)
set fso = Server.CreateObject("scripting.filesystemobject") 
If fso.folderexists(Server.MapPath(""&str&"")) Then
Else 
fso.CreateFolder(Server.MapPath(""&str&""))
End If 
end function

function delwjj(str1)
set fso = Server.CreateObject("scripting.filesystemobject") 
If fso.folderexists(Server.MapPath(""&str1&"")) Then
fso.deletefolder(Server.MapPath(""&str1&""))
Else 
End If 
end function
function getip()
		getip=request.servervariables("http_x_forwarded_for")
		if getip="" then getip=request.servervariables("remote_addr")
	end function

'===========================================================
    '检查数据表是否存在
    '===========================================================

Function CheckTable(TabName)
CheckTable = False
Dim iRs
Set iRs = conn.openSchema(20)
Do While Not iRs.EOF
If LCase(iRs(3)) = "table" Then
If LCase(iRs(2)) = LCase(TabName) Then
CheckTable = True
End If
End If
iRs.MoveNext
Loop
End Function


    '===========================================================
    '创建表
    '===========================================================

   Function CreateTable(TableName)
       
     if  not CheckTable(TableName) then
       
            If sqlno = 1 Then
                SQL = "CREATE Table ["& TableName &"] (Id INT Identity (1, 1) Primary Key)"
            Else
                SQL = "CREATE Table ["& TableName &"] (Id INT IDENTITY (1, 1) NOT NULL PRIMARY KEY)"
            End If
            conn.execute(SQL)
      else
	  
	  response.Write "<script language=javascript>alert(""非常抱歉！此数据表已经存在！请更换数据表"");window.location.href=""bd.asp""</script>"
	  response.End()
      end if
	  
    End Function
	
	
	
		    '===========================================================

    '判断字段是否存在通用函数
    '===========================================================

   Function IsFieldsText(TableName,ColumnName)
        Dim FieldsRs, imNum, iFoundErr
        iFoundErr = False
        Set FieldsRs = conn.execute("Select Top 1 * From " & TableName & "")
        For imNum = 0 To FieldsRs.Fields.Count - 1
            If FieldsRs.Fields(imNum).Name = ColumnName Then
                iFoundErr = True
                Exit For
            End If
        Next
        IsFieldsText = iFoundErr
    End Function
	   Function AddColumn(TableName,ColumnName,ColumnType)
        Dim Result, ErrMsg
        Result = conn.execute("Alter Table "& TableName &" Add "& ColumnName &" "& ColumnType &" ")
        If Result = 0 Then
            ErrMsg = "新建 "& TableName &" 表中字段<span style=""color:#00F"">错误</span>，请手动将数据库中 <b>"& ColumnName &"</b> 字段建立，属性为 <b>"& ColumnType &"</b>"
        Else
            ErrMsg = "新建 "& TableName &" 表中字段 "& ColumnName &" 成功"
        End If
        response.write ErrMsg
    End Function



'===========================================================
    '删除字段通用函数
    '===========================================================

  Function DelColumn(TableName,ColumnName)
        Dim Result, ErrMsg
        If sqlno=1 then
        Result = conn.execute("Alter Table "& TableName &" Drop "& LCase(ColumnName) &" ")
        Else
        Result = conn.execute("Alter Table "& TableName &" Drop COLUMN "& LCase(ColumnName) &" ") 
        End if
        If Result = 0 Then
            ErrMsg = "删除 "& TableName &" 表中字段<span style=""color:#00F"">错误</span>，请手动将数据库中 <b>"& ColumnName &"</b> 字段删除"
        Else
            ErrMsg = "删除 "& TableName &" 表中字段 "& ColumnName &" 成功"
        End If
        DelColumn = ErrMsg
    End Function



    '===========================================================
    '更改字段通用函数
    '===========================================================

 Function ModColumn(TableName,ColumnName,ColumnType)
        Dim Result, ErrMsg
        Result = conn.execute("Alter Table "& TableName &" Alter Column "& ColumnName &" "& ColumnType &"")
        If Result = 0 Then
            ErrMsg = "更改 "& TableName &" 表中字段属性<span style=""color:#00F"">错误</span>，请手动将数据库中 <b>"& ColumnName &"</b> 字段更改为 <b>"& ColumnType &"</b>"
        Else
            ErrMsg = "更改 "& TableName &" 表中字段属性 "& ColumnName &" 成功"
        End If
        ModColumn = ErrMsg
    End Function







	    
%>





