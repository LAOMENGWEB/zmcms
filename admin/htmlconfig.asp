
<script language="VBSCRIPT" runat="server">

Function htmll(mulu, htmlmulu, FileName, filefrom, htmla, htmlb, htmlc, htmld, htmle, htmlf, htmlg, htmlh,htmli,htmlj,htmlk,htmlm)
SysRootDir=""&sitepath&""
if mulu="" then mulu="/"
if htmlmulu="" then htmlmulu="/"
mulu=replace(SysRootDir&mulu, "//", "/")
htmlmulu=replace(SysRootDir&htmlmulu, "//", "/")
FilePath=Server.MapPath(mulu)&"\"&FileName
Do_Url="http://"
Do_Url=Do_Url&Request.ServerVariables("server_name")&htmlmulu&filefrom
Do_Url=Do_Url&"?"&htmla&htmlb&"&"&htmlc&htmld&"&"&htmle&htmlf&"&"&htmlg&htmlh&"&"&htmli&htmlj&"&"&htmlk&htmlm
strUrl=Do_Url
set objXmlHttp=Server.createObject("Microsoft.XMLHTTP")
objXmlHttp.open "GET",strUrl,false
objXmlHttp.send()
binFileData=objXmlHttp.responseBody
Set objXmlHttp=Nothing
set objAdoStream=Server.CreateObject("Adodb." & "Stream")
objAdoStream.Type=1
objAdoStream.Open()
objAdoStream.Write(binFileData)
objAdoStream.SaveToFile FilePath,2
objAdoStream.Close()
set objAdoStream=nothing
End Function
</script>
