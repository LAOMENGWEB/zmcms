<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
if Instr(session("AdminPurview"),"|71,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request.Form("submit") = "确定" Then
HrName=Trim(Request("HrName"))
HrRequireNum=Trim(Request("HrRequireNum"))
HrAddress=Trim(Request("HrAddress"))
HrSalary=Trim(Request("HrSalary"))
ksDate=Trim(Request("ksDate"))
jsDate=Trim(Request("jsDate"))
Hits=trim(request.form("Hits"))
btcolor=trim(request.form("btcolor"))
HrDetail=Request("Content")
zpdw=Request("zpdw")
zdfy=trim(request.form("zdfy"))
dh=Request("dh")
dz=Request("dz")
emaill=Request("emaill")

if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
		end if
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from job "
rs.open sql,conn,1,3
rs.addnew
rs("HrName")=HrName
rs("HrRequireNum")=HrRequireNum
rs("HrAddress")=HrAddress
rs("HrSalary")=HrSalary
rs("ksDate")=ksDate
rs("jsDate")=jsDate
rs("HrDetail")=HrDetail
rs("zpdw")=zpdw
rs("zdfy")=zdfy
rs("dh")=dh
rs("lang")=lang
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=7 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=7 order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd2=request.form(""&zdyzd1&"")
rs(""&zdyzd1&"")=zdyzd2
rsObj.movenext
loop
end if
'自定义字段结束
rs("dz")=dz
rs("emaill")=emaill
rs("hits")=hits
rs("btcolor")=btcolor
rs.update


rs.close
set rs=nothing

'获取添加新闻的ID  生成html开始
set rs=conn.execute("select top 1 id from [job] order by id desc")
zpid=rs(0)
rs.close:set rs=Nothing




if yy=2 then

sql="Select * from job where id="&zpid&" order by ID desc"
Set Rs=Conn.Execute(sql)
content=rs("HrDetail")
zdfy=rs("zdfy")
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&zpname&""&Separated&""&zpid&""&Separated&"1.html",""&showjob&"/index.asp","jobid=",zpid,"m=",mzp,"language=",lang,"page=",1,"","","","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&zpname&""&Separated&""&zpid&""&Separated&""&i&".html",""&showjob&"/index.asp","jobid=",zpid,"m=",mzp,"language=",lang,"page=",i,"","","","")
next
end if


rs.close
set rs=Nothing
















call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&job&"/","",""&zpclass&".html",""&job&"/index.asp","m=",mzp,"language=",lang,"","","","","","","","")

end if
call addlog("发布"&zhaopin&"")
 response.Write "<script language=javascript>alert("""&zhaopin&"信息成功发布"");window.location.href=""zp.asp""</script>"
    response.End

end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="javascript" type="text/javascript" src="js/getdate/WdatePicker.js"></script>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<!--#include file="ys.asp"-->
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
<form method="post"  name="myform" class="checkform">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">发布<%=zhaopin%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="zhaopin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 <div class="dmeun"><a href="zp.asp" target="main" style="color:#ffffff"><%=zhaopin%>管理</a></div>
 </td>
 

 </tr>
</table>
           
                <table width="99%"  border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
				 
				  <tr> 
                    <td  > 
                    <%=zhaopin%>单位:</td><td><input name="zpdw" type="text" id="zpdw" datatype="*" nullmsg="<%=zhaopin%>单位不能为空"> <span style="color:#ff0000">*</span></td>
                  </tr>
                  <tr> 
                    <td  > 
                    <%=zhaopin%>职位:</td><td><input 
name="HrName" type="text" id="HrName" datatype="*" nullmsg="<%=zhaopin%>职位不能为空">  <span style="color:#ff0000">*</span> </td>
                  </tr>
				  
				  <tr> 
                    <td  > 
                    标题颜色:</td>
				  <td><input name="btcolor"type="text" id="color"   value="" size="8" /> 
 <input   type="button" id="colorpicker" value="打开取色器" /></td>
				  </tr>
				  
				  
                  <tr> 
                    <td  > 
                    <%=zhaopin%>人数:</td><td><input 
name="HrRequireNum" type="text" id="HrRequireNum" datatype="*" nullmsg="<%=zhaopin%>人数不能为空"> 10人 <span style="color:#ff0000">*</span></td>
                  </tr>
                  <tr> 
                    <td  > 
                    工作地点:</td><td><input 
name="HrAddress" type="text" id="HrAddress" datatype="*" nullmsg="工作地点不能为空"> <span style="color:#ff0000">*</span></td>
                  </tr>
                  <tr> 
                    <td  > 
                    工资待遇:</td><td><input 
name="HrSalary" type="text" id="HrSalary" datatype="*" nullmsg="工资待遇不能为空"> 1000元/月 <span style="color:#ff0000">*</span></td>
                  </tr>
				                    <tr> 
                    <td  > 
                      联系电话:</td><td><input 
name="dh" type="text" id="dh" datatype="*" nullmsg="联系电话不能为空"> <span style="color:#ff0000">*</span></td>
                  </tr>
                  <tr> 
                    <td  > 
                    联系地址:</td><td><input 
name="dz" type="text" id="dz" datatype="*" nullmsg="联系地址不能为空"> <span style="color:#ff0000">*</span></td>
                  </tr>
                  <tr> 
                    <td  > 
                    投递邮箱:</td><td><input 
name="emaill" type="text" id="emaill" datatype="*" nullmsg="邮箱不能为空"> <span style="color:#ff0000">*</span></td>
                  </tr>

                  <tr> 
                    <td  > 
                    开始日期:</td><td><input name="ksdate"type="text" id="ksdate"  onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})"
 value="<% if ksDate="" then response.write now() else response.write (ksDate) end if%>" />                    
                    <span style="color:#ff0000">*</span>&nbsp;默认为当前时间，可手动输入日期格式 [提示：可以定时发布<%=zhaopin%>，点击边框设置时间] </td>
                  </tr>
                  <tr>
                    <td  >结束日期:</td><td>
                      <input name="jsDate" type="text" class="textfield" id="jsDate"  value="<% if jsDate="" then response.write (DateAdd("m",1,now())) else response.write (jsDate) end if%>"   onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})"
>
                    <span style="color:#ff0000">*</span>&nbsp;默认为1个月，可手动输入日期格式</td>
                  </tr>
                	 <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="content"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
 afterBlur: function(){this.sync();}
 });
 	K('input[name=appendHtml]').click(function(e) {
					editor.insertHtml('[zmcmsfy]');
				});
			
			prettyPrint();
 });
 </script>
                  <tr> 
                    <td  > 
                    <%=zhaopin%>要求:</td><td><textarea datatype="*" nullmsg="<%=zhaopin%>要求不能为空" name="content" style="width:700px;height:300px;visibility:hidden;"></textarea></td>
                  </tr>
				  
				  
				  
				   <tr >
		     <td >手动分页：</td>
			 <td><INPUT type=button class="bnt" name="appendHtml" value="插入分页符"></td>
			 </tr>
			 <tr >
			<input name="zdfy" type="hidden" id="zdfy"  value="0" size="10">
	
				  
				  
				  
				  
				  	<%'自定义字段%>
<%call showColumn("job",7,id)%>
				  
				  
				  
				  
				  
                  <tr> 
                    <td colspan="2"> 
                       
                        <input name="submit"type="submit" id="submit"  value="确定" class="bnt">
                       
                    </td>
                  </tr>
				  
				  
				  
				  
				  
            </table>
           
          </form>
</body>
</html>
