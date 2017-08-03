<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|1131,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
If request.Form("submit") = "确认添加" Then
    mc = request.Form("mc")
    dz = request.Form("dz")
    px = request.Form("px")
	lx = request.Form("lx")
	logourl= request.Form("logourl")
	zmone=Trim(request.form("zmone"))



    Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From yqlj ",conn,1,3
    rs.addnew
    rs("mc") = mc
    rs("dz") = dz
    rs("px") = px
	rs("lx") = lx
	rs("zmone")=zmone
	rs("logourl") = logourl
	rs("lang")=lang
    rs.update
    rs.Close
    Set rs = Nothing
   call addlog("添加链接")
    response.Write "<script language=javascript>alert(""链接添加成功"");window.location.href=""yq.asp""</script>"
    response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>
<script type="text/javascript">

function deltp()
{
var parent=document.getElementById("mysmaimg");
var child=document.getElementById("sc");
var child1=document.getElementById("lj");
parent.removeChild(child);
parent.removeChild(child1);
document.getElementById("weepic").value="";
}

 </script>
 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加链接  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="lianjie.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="yq.asp" target="main">链接管理</a></div></td>
 

 </tr>
</table>
 <form method="post" name="yqlj" class="checkform">
<table width="99%"  border="0" align="center" cellpadding="0" cellspacing="0" class="dtable4">
       
		 
          			  <tr  >
				    <td width="11%" >所属分类: </td>
				    <td width="89%"><select name="zmone" id="select" datatype="*" nullmsg="所属分类不能为空">
            <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from yqljclass",conn,1,1
		  Response.Write("<option value="""">请选择分类</option>") 
    If Rsc.Eof and Rsc.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
    Else
		  while not rsc.eof
			response.Write("<option value=""" & rsc("id") & """>" & rsc("classname") & "</option>")
			rsc.movenext
		wend
		end if
		rsc.close
		set rsc=nothing
		  %>
                    </select></td>
			      </tr>
          <tr>
            <td  >链接名称： </td><td>
              <input name="mc" type="text" id="mc"  size="20" datatype="*" nullmsg="名称不能为空"/>
             </td>
          </tr>
		   

			  <tr>
			    <td >链接图片： </td><td>
				   <input name="logourl"  type="text" id="weepic"  value="" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><div style="width:100px;padding:10px; background:#28b576;color:#ffffff" align="center">logo上传</div></a><br/><br/>
		        <div id="mysmaimg"></div>		        </td>
	      </tr>
          <tr>
            <td >链接地址： </td><td>
              <input name="dz" type="text" id="dz"  datatype="*" nullmsg="链接地址不能为空"/>
             </td>
          </tr>
         <tr>
            <td>排序： </td><td>
              <input name="px" type="text" id="px"  value="1"/>
       </td>
          </tr>
        
		
          <tr >
            <td colspan="2" ><input name="submit" type="submit" class="bnt" value="确认添加" /></td>
          </tr>
     
</table>
   </form>
</body>
</html>
<% conn.Close
    Set conn = Nothing%>