
<%
sub DeleteFile(path)
path = server.mappath(path)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then fso.DeleteFile(path)
set fso=nothing
end sub
id=request.QueryString("id")
call deletefile(id)

%>

