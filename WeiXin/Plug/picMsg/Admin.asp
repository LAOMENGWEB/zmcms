<!--#include file="../base.asp"-->
<!--#include file="../../lib/db.asp"-->
<!--#include file="../../lib/function.asp"-->
<!--#include file="../../lib/weixin_class.asp"-->
<!--#include file="../../lib/md5.asp"-->
<!--#include file="class_page.asp"-->

<%
Result=request.QueryString("Result")
keytag=request.Form("k1")
pone=request.QueryString("pone")
If request.Form("button") = "批量删除" Then
if Request("id")="" then
Response.Write "<script>alert('错误!请选择要操作的记录!');window.location.href='admin.asp';</script>" 
response.end()
end if
for y=1 to request("ID").count
if request("ID").count=1 then
ArticleID=request("ID")
else
ArticleID=replace(request("id")(y),"'","")
end if	


set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Plug_PicMsg where P_ID in ("&ArticleID&")",conn2,0,1

p_Pic=rs("p_Pic")
dim picUrl
dim picUrlArray
dim x,y
picUrl = getPicUrl(rs("P_content"))
if picUrl <> "" then
picUrl = left(picUrl,len(picUrl)-1)
picUrlArray = split(picUrl,",")
for x = 0 to ubound(picUrlArray)
if instr(picUrlArray(x),"") > 0 then
call DeleteFile(picUrlArray(x))
end if
next
end if

dim picUrl1
dim picUrlArray1
dim x1,y1
picUrl1 = rs("ImagePath")
if picUrl1 <> "" then
picUrlArray1 = split(picUrl1,"|")
for x1 = 0 to ubound(picUrlArray1)
call DeleteFile(picUrlArray1(x1))
next
end if
call deletefile1(p_Pic)


next
rs.close:set rs=Nothing

conn2.Execute ( "delete from sys_KeyWord where k_plugName='picMsg' and k_plugParam in ("&Request("id")&")" )





conn2.Execute ( "delete from Plug_PicMsg where P_id in ("&Request("id")&")" )

response.Write "<script language=javascript>alert(""删除成功！"");window.location.href=""admin.asp""</script>"
end if




%>





<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="../../css/wx.css" rel="stylesheet" type="text/css">

<script language="javascript"> 
<!-- 
function CheckAll(){ 
 for (var i=0;i<eval(form1.elements.length);i++){ 
  var e=form1.elements[i]; 
  if (e.name!="allbox") e.checked=form1.allbox.checked; 
 } 
} 
--> 
</script> 

</head>

<body>

      <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >自定义单图文素材管理</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 

 <div class="dmeun"><a href="Admin_add.asp" target="main" style="color:#ffffff">添加自定义素材</a></div>
 <div class="dmeun"><a href="Admin_class.asp" target="main" style="color:#ffffff">素材分类管理</a></div>
 </td>
 

 </tr>
</table>	
<table width="99%"  cellpadding="0" cellspacing="0" class="dtable1" align="center">
  		
          <tr   >
            <td width="11%">按照分类查找：</td>
            <td width="54%"><select class="wbk" name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from Plug_PicMsg_class "   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn2,1,1 
  Response.Write("<option value=?pone=0")
		 If request.QueryString("pone")=0 then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">所有分类 </option>")  & VbCrLf
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?pone=" & Rsp("c_ID") & """" & "")
		 If int(request("pone"))=Rsp("c_ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">" & Rsp("c_ClassName") & " </option>")  & VbCrLf
		 
      Rsp.Movenext   
      Loop   
   End if
	rsp.close
	set rsp=nothing
	%>
        </select></td>
            <form id="form2" name="form2" method="post" action="admin.asp?Result=ok">
            <td width="8%"  >
            快速搜索：                      </td>
            <td width="27%"  > <input name="k1" type="text" id="k1" />
             
              <input class="bnt" type="submit" name="button" id="button" value="搜索" /> </td>
        </form>
  </tr>
</table>


<%
if Result="ok"   then
 DB_Sql = "select * from Plug_PicMsg where p_title like '%"&keytag&"%' "
else

if pone=0 then
DB_Sql = "select * from Plug_PicMsg "
end if
if pone<>0  then
DB_Sql = "select * from Plug_PicMsg where p_classid="&Pone&" "
end if
end if
DB_Sql = DB_Sql & "order by P_ID desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn2,1,1
If Not(rs.Bof and rs.Eof) Then


'================================================
'调用分页类开始
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 20 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum

if Result="ok" then
Obj_Page.Setpage_Pathurl = "admin.asp?Result=ok" '执行分页的页面
else
if pone=0 then
Obj_Page.Setpage_Pathurl = "admin.asp" '执行分页的页面
else
Obj_Page.Setpage_Pathurl = "admin.asp?pone="&pone&"" '执行分页的页面
end if
end if

Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页

%>

<form action="admin.asp" method="post" name="form1" id="form1"> 

            
              <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
               
                <tr class="wz">
                  <td width="66" >全选</td>
                  <td width="47"    >编号</td>
                  <td width="183"    >类型</td>
                  <td width="144"     >分类</td>
                  <td width="150"    >标题</td>
     
                  <td width="196"    >添加时间</td>
                  <td width="402"> 操作</td>
                </tr>
              <%

'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>
                <tr  >
                  <td  ><input name="ID" type="checkbox" id="ID" value="<%=rs("P_id")%>" class="none"/></td>
                  <td  ><%=rs("P_id")%></td>
                  <td  ><%=rs("P_Type")%> </td>
              
                 
                  <td  ><%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from Plug_PicMsg_class",conn2,1,1
		  while not rsc.eof
		    if rs("p_classid")=rsc("c_id") then
			response.Write("" & rsc("c_classname") & "")
			end if
			rsc.movenext
		wend
		rsc.close
		set rsc=nothing
		  %></td>
		   <td   ><%=rs("P_Title")%></td>
   
                  <td  ><%=rs("P_addtime")%></td>
                  <td >				  
                    
                    &nbsp;<a href="admin_edit.asp?ID=<%=rs("p_ID")%>">修改</a>&nbsp; 
                    &nbsp;<a href="admin_Del.asp?ID=<%=rs("p_ID")%>" onClick="return confirm('确定要删除此素材和与之关联的关键字吗？');" >删除</a>
				  </td>
                </tr>
			
 <%
rs.MoveNext
Next
	
%>      
<tr  >
                  <td   > 
                   
                      <input type="checkbox" name="allbox" onClick="CheckAll()" class="none"/>
     </td>
      <td  colspan="7"  style="padding:5px; text-align:left"><input class="bnt" type="submit" name="button" id="button" value="批量删除"  onClick="return confirm('确定要批量删除选中的素材和与之关联的关键字吗？');"/>
	  </td>
</tr>
  </table>

  <%If Obj_Page.OutErrInfo <> "" Then
		Response.Write Obj_Page.OutErrInfo '报错信息显示
		Response.End
		else
	Call Obj_Page.Page_Navigation() '分页导航
	end if
    else

	%>
	<table width="99%"  cellpadding="0" cellspacing="0" >
  		
          <tr   >
            <td  height="30"  style="padding-left:10px;border-bottom:none; text-align:center">  <%Response.Write "目前暂无任何信息"
Response.end%> </td>       
  </table>

</form><%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</body>
</html>
