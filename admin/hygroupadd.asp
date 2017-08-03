<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!-- #include file="admin.asp" -->
<%
if Instr(session("AdminPurview"),"|1002,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
dim RanNum,GroupID

If request.Form("submit") = "确认添加" Then
 
	  dim groupname
groupname=trim(request("groupname"))
set rs = server.createobject("adodb.recordset")
sql="select * from hyGroup where groupname='" & groupname & "'"
      rs.open sql,conn,1,3
	  if not (rs.bof and rs.EOF) then
		response.Write "<script language=javascript>alert(""会员组别已经存在"");window.history.back(-1);</script>"
        response.End
		end if
      rs.addnew
	  rs("GroupID")=Request.Form("GroupID")
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("GroupLevel")=trim(Request.Form("GroupLevel"))
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("Adddate")=now()
	  rs("lang")=lang
	  rs.update
       rs.close
	   call addlog("添加会员组别")
	response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""hygroup.asp""</script>"
	end if
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
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>
<FORM name="form_admin" method="post"  class="checkform">	

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加会员组  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="huiyuan.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>

            <table width="99%"   align="center"  class="dtable4">
               
            
				<%
				Randomize() 
				ranNum=int(90000*rnd)+10000
				GroupID=second(now)&ranNum
				%>
                <tr > 
                  <td > 会员组别号：  </td><td>                 <input name="GroupID" type="text" id="GroupID" value="<%=groupID%>" size="16" maxlength="20"
>
                  (组别号必须是唯一的)</td>
                </tr>
                <tr >
                  <td >会员组别名：</td><td> 
               
                  <input name="Groupname" type="text" id="Groupname" size="16" maxlength="20" datatype="*" nullmsg="会员组名称不能为空"/></td>
                </tr>
                <tr > 
                  <td > 权限值：   </td><td>                 <input name="GroupLevel" type="text" id="GroupLevel" size="16" maxlength="20" 
datatype="*" nullmsg="权限值不能为空">
                    权限值中必须输入整数（每一个权限值 一定是不一样的  一定要注意！）</td>
                </tr>
                
                
                
                
                <tr >
                  <td >备注说明:</td><td> 
                    <label>
                    <textarea name="Explain" cols="50" rows="6" id="Explain"></textarea>
                  </label></td>
                </tr>
                
                <tr > 
                  <td  colspan="2"><div align="left">
                    <input name=Submit type=submit class="bnt" id="Submit" value='确认添加' /></div>                  </td>
                </tr>
         </table>
</form>

</body>
</html>
