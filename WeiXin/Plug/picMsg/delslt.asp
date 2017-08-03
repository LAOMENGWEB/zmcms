
<%
sub DeleteFile1(path)
path = server.mappath("/" & path)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then fso.DeleteFile(path)
set fso=nothing
end sub
sltdz=request.QueryString("sltdz")
session("sltdz")=sltdz
call deletefile1(sltdz)

%>

<script LANGUAGE="JavaScript">
top.window.opener = top;
top.window.open('','_self','');
top.window.close();

 </script>