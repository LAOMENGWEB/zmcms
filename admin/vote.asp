<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="class_page.asp"-->
<%
if Instr(session("AdminPurview"),"|1081,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>

<SCRIPT language=javascript>
    function unselectall(thisform){
        if(thisform.chkAll.checked){
            thisform.chkAll.checked = thisform.chkAll.checked&0;
        }   
    }
    function CheckAll(thisform){
        for (var i=0;i<thisform.elements.length;i++){
            var e = thisform.elements[i];
            if (e.Name != "chkAll"&&e.disabled!=true)
                e.checked = thisform.chkAll.checked;
        }
    }

</SCRIPT>
<script language="javascript" type="text/javascript" src="js/getdate/WdatePicker.js"></script>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>

<%
	if request("action") = "add" then 
		call add()
	elseif request("action")="edit" then
		call edit()
	elseif request("action")="savenew" then
		call savenew()
		
	elseif request("action")="savedit" then
		call savedit()
	
	elseif request("action")="delAll" then
		call delAll()
	elseif request("action")="show" then
		call show()
	else
		call List()
	end if

sub list
%>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">调查管理 <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


 <table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
   <tr>
     <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
         <div class="dmeun"><a href="vote.asp?action=add" target="main">添加调查</a></div></td>
   </tr>
 </table>
 <form name="myform" method="POST" action="Vote.asp?action=delAll">
<%
DB_Sql = "select * from Vote where lang='"&lang&"' "
DB_Sql = DB_Sql & "order by ID desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open DB_Sql,Conn,1,1
If Not(rs.Bof and rs.Eof) Then


'================================================
'调用分页类开始
Dim Obj_Page, Obj_PageListnum, II
Obj_PageListnum = 20 

Set Obj_Page = New HHYY_Pageclass

Obj_Page.Setpage_RS = rs
Obj_Page.Setpage_Listnum = Obj_PageListnum

Obj_Page.Setpage_Pathurl = "vote.asp" '执行分页的页面
Obj_Page.Setpage_Navtype = 1 '分页导航样式(可选1文字,2数字)

Call Obj_Page.Page_Init() '初始化分页
%>
<table width="99%" border="0"  align="center" cellpadding="0" cellspacing="0"  class="dtable2">

  <tr class="wz" > 
    <td width="11%" >全选</td>
    <td width="26%" >投票标题</td>
    <td width="9%" >编号</td>
    <td width="21%" >投票时间</td>
    <td width="11%" >状态</td>
    <td width="5%" >排序</td>
    <td width="17%" >操 作</td>
  </tr>

<%
'/循环开始
For II = 1 to Obj_PageListnum
If rs.Eof Then
	Exit For
End If
%>

    <tr >
    <td ><input type="checkbox" value="<%=rs("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td ><%=rs("Title")%></td>
    <td ><%=rs("ID")%></td>
    <td ><%=rs("StartTime")%>至<%=rs("EndTime")%></td>
    <td >  <%if rs("yn")=1 then
        Response.Write "<font color='blue'>显示</font>" & vbCrLf
      else
        Response.Write "<font color='red'>隐藏</font>" & vbCrLf
	  end If%></td>
    <td height="25" align="center"><%=rs("Px")%></td>
    <td align="center" style="border-right:none"><a href="?action=edit&id=<%=rs("ID")%>">编辑</a> <a href="?action=show&id=<%=rs("ID")%>">查看结果</a></td>
    </tr>
 <%
rs.MoveNext
Next	
%> 
<tr>
  <td ><input name="Action" type="hidden"  value="Del"><input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0;"></td>
  <td colspan="7" style="padding:5px; text-align:left"><div align="left">
    <input name="Del" type="submit" class="bnt" id="Del" value="隐藏">
    <input name="Del" type="submit" class="bnt" id="Del" value="显示">
    <input name="Del" type="submit" class="bnt" id="Del" value="删除">
  </div></td>
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
	<table width="99%"  cellpadding="0" cellspacing="0" class="table1" style="margin-top:10px">
  		
          <tr   >
            <td  height="30"  style="padding-left:10px;border-bottom:none; text-align: center;">  <%Response.Write "目前暂无任何信息"
Response.end%> </td>       
  </table>
<%Set Obj_Page = Nothing

RS.Close
Set RS = Nothing
End If
%>
</form>

<%

end sub

Sub add()
%>



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加调查  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


 <table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
   <tr>
     <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
	 <div class="dmeun"><a href="vote.asp" target="main">调查管理</a></div>
        </td>
   </tr>
 </table>


  <form name="add" method="post" action="?action=savenew" class="checkform">
  <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0"  class="dtable4">

    
    <tr>
      <td width="11%"  >投票名称：      </td>
      <td width="89%" ><input name="t0" type="text" class="input" id="t0" size="50" maxlength="50" datatype="*" nullmsg="名称不能为空"></td>
    </tr>
    <tr>
      <td >投票选项：      </td>
      <td><input name='t1' type='radio' class="noborder" value='1'  checked />单选 <input  name='t1' type='radio' class="noborder" value='0' />多选</td>
    </tr>
    <tr>
      <td >项目内容：<br><br>
      <font class="note">格式：<br><br>项目标题|票数</font></td>
      <td >
      <textarea name="votes" id="votes" cols="50" rows="10">投票选项一|0
投票选项二|0</textarea></td>
    </tr>
    <tr>
      <td  >投票时间：</td>
      <td ><input name="StartTime" type="text" class="input" value="<%=Now%>" size="20" onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})">
        -
        <input name="EndTime" type="text" class="input" value="<%=Now+30%>" size="20" onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})"></td>
    </tr>
    <tr>
      <td >排序：</td>
      <td ><input name="Px" type="text" id="Px" value="1" size="6" maxlength="5">
      <span class="note">数字小的排在前面</span></td>
    </tr>
 
    <tr>
	  <td colspan="2" style="border-right:none;border-bottom:none"><input name="Submit" type="submit" class="bnt" value="添 加">
      <input name="Submit23" type="button" class="bnt" onClick="history.go(-1)" value="返 回"></td>
    </tr>
	
</table>
</form>
<%
end sub

sub edit
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select id,title,stype,vote,result,StartTime,EndTime,Px from Vote where id="& id &" and lang='"&lang&"'"
rs.open sql,conn,1,1

If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""vote.asp""</script>"
    response.End
End if
%>



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改调查  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



 <table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
   <tr>
     <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
	 <div class="dmeun"><a href="vote.asp" target="main">调查管理</a></div>
        </td>
   </tr>
 </table>


 <form name="add" method="post" action="?action=savedit&id=<%=id%>" class="checkform">
  <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
 

    <tr>
      <td width="11%" align="center" >投票名称：      </td>
      <td width="89%" ><input name="t0" type="text" class="input" id="t0" value="<%=rs(1)%>" size="50" maxlength="50" datatype="*" nullmsg="名称不能为空"></td>
    </tr>
    <tr>
      <td align="center" >投票选项：      </td>
      <td ><input name='t1' type='radio' class="noborder" value='1' <%If rs(2)=1 then Response.Write("checked") end if%>/>
      单选 <input  name='t1' type='radio' class="noborder" value='0' <%If rs(2)=0 then Response.Write("checked") end if%> />
      多选</td>
    </tr>
    <tr>
      <td align="center" >项目内容：</td>
      <td >
      <textarea name="votes" id="votes" cols="50" rows="10"><%
result=split(rs(4),"|")
	for i=0 to ubound(result)
next
vote=split(rs(3),"|")
for i=0 to ubound(vote)-1
	Response.Write CHR(10)&""&vote(i)&"|"&result(i)&""
next
%></textarea></td>
    </tr>
    <tr>
      <td align="center" >投票时间：</td>
      <td ><input name="StartTime" type="text" class="input" value="<%=rs(5)%>" size="20" onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})">
        -
        <input name="EndTime" type="text" class="input" value="<%=rs(6)%>" size="20" onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})"></td>
    </tr>
    <tr>
      <td align="center" >排序：</td>
      <td ><input name="Px" type="text" id="Px" value="<%=rs(7)%>" size="6" maxlength="5">
        <span class="note">数字小的排在前面</span></td>
    </tr>
	 
    <tr>
	  <td colspan="2" style="padding-left:10px"><input name="Submit" type="submit" class="bnt" value="保 存">
     </td>
    </tr>

  </table>
  	</form>
<%
rs.close
set rs=nothing
end sub

Sub savenew()

t0			=trim(request("t0"))
t1			=trim(request("t1"))
votes		=trim(request("votes"))
StartTime	=trim(request("StartTime"))
EndTime		=trim(request("EndTime"))
Px			=trim(request("Px"))

If votes="" then
Call Alert ("项目内容不能为空!",-1)
End if
 
votes=split(votes,CHR(10))
for i=0 to ubound(votes)
if not instr(trim(votes(i)),"|")>0 then
Call Alert ("投票选项"&i+1&"有错误!",-1)
end if
 somevote=split(trim(votes(i)),"|")
   for j=0 to ubound(somevote)
   s=trim(somevote(0))
   s1=trim(somevote(1))
   if not isnumeric(s1) then
   Call Alert ("投票选项"&i+1&"有错误!",-1)
   end if
   next  
   vote=vote&s&"|"
   result=result&s1&"|"
next

	set rs = server.CreateObject ("adodb.recordset")
	sql="Select * from Vote where Title='"& Title &"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("Title")			=t0
		rs("vote")			=vote
		rs("result")		=result
		rs("stype")			=t1
		rs("StartTime")		=StartTime
		rs("EndTime")		=EndTime
		rs("Px")			=Px
		rs("yn")			=1
		rs("lang")          =lang
		rs.update
		call addlog("添加投票成功")
		Call Alert("恭喜,添加成功！","Vote.asp")
	else
	call addlog("添加投票失败")
		Call Alert("添加失败，该投票已经存在！",-1)
	end if
	rs.close
	set rs=nothing
	
End sub

Sub savedit()

t0=trim(request("t0"))
t1=trim(request("t1"))
votes=trim(request("votes"))
ID=trim(request("id"))
StartTime	=trim(request("StartTime"))
EndTime		=trim(request("EndTime"))
Px			=trim(request("Px"))

If votes="" then
Call Alert ("项目内容不能为空!",-1)
End if

votes=split(votes,CHR(10))
for i=0 to ubound(votes)
	if not instr(trim(votes(i)),"|")>0 then
	Call Alert ("投票选项"&i+1&"有错误!",-1)
	end if
	somevote=split(trim(votes(i)),"|")
	for j=0 to ubound(somevote)
	s=trim(somevote(0))
	s1=trim(somevote(1))
		if not isnumeric(s1) then
		Call Alert ("投票选项"&i+1&"有错误!",-1)
		end if
	next  
	vote=vote&s&"|"
	result=result&s1&"|"
next

	set rs = server.CreateObject ("adodb.recordset")
	sql="Select * from Vote where ID="& ID &""
	rs.open sql,conn,1,3
		rs("Title")			=t0
		rs("vote")			=vote
		rs("result")		=result
		rs("stype")			=t1
		rs("StartTime")		=StartTime
		rs("EndTime")		=EndTime
		rs("Px")			=Px
		rs("yn")			=1
		rs("lang")          =lang
		rs.update
		call addlog("修改投票")
		Call Alert("恭喜,修改成功！","Vote.asp")
		
	rs.close
	set rs=nothing
	
End sub

sub show
id=request("id")
set rs=conn.execute("select id,title,vote,result from vote where id="&id&" and lang='"&lang&"'")
if rs.eof then
  response.Write "<script language=javascript>window.location.href=""vote.asp""</script>"
    response.End
else
vote=rs(2)
result=rs(3)
total_vote=0
vote=split(vote,"|")
result=split(result,"|")
for i=0 to ubound(result)
if not result(i)="" then total_vote=result(i)+total_vote
next
end if
%>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"> 调查结果显示  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


 <table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
   <tr>
     <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
	 <div class="dmeun"><a href="vote.asp" target="main">调查管理</a></div>
        </td>
   </tr>
 </table>




<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" class="dtable2">
  <tr class="wz">
      <td width="43%" ><div align="center">选性</div></td>
      <td width="57%" ><div align="center">结果</div></td>
  </tr>
<%for i=0 to (ubound(vote)-1)%>
  
    <tr>
	  <td height="25" align="center" ><%Response.Write ""&i+1&". "&vote(i)&""%></td>
      <td align="center" style="border-right:none"><%Response.Write ""&result(i)&" "%>
        票</td>
    </tr>
<%next%>
<tr>
    <td colspan="2"   style="border-bottom:none;padding-left:10px; text-align:left"><input name="Submit2" type="button" class="bnt" onClick="history.go(-1)" value="返 回"></td>
  </tr>
</table>
<%
End Sub

Sub delAll
ID=Trim(Request("ID"))
page=request("page")
If ID="" Then
	  Call Alert("请选择记录!",-1)
ElseIf Request("Del")="隐藏" Then
   set rs=conn.execute("Update Vote set yn = 0 where ID In(" & ID & ")")
   call addlog("隐藏投票")
   Call Alert ("操作成功！","Vote.asp")
   	
ElseIf Request("Del")="显示" Then
   set rs=conn.execute("Update Vote set yn = 1 where ID In(" & ID & ")")
   call addlog("显示投票")
   Call Alert ("操作成功！","Vote.asp")
ElseIf Request("Del")="删除" Then
	set rs=conn.execute("delete from Vote where ID In(" & ID & ")")
	call addlog("删除投票")
	Call Alert ("操作成功！","Vote.asp")
End If
End Sub
%>
</body>
</html>
