
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%

id=request.QueryString("id")
If request.Form("submit") = "修改" Then
HrName=Trim(Request("HrName"))
HrRequireNum=Trim(Request("HrRequireNum"))
HrAddress=Trim(Request("HrAddress"))
HrSalary=Trim(Request("HrSalary"))
ksDate=Trim(Request("ksDate"))
jsDate=Trim(Request("jsDate"))
btcolor=trim(request.form("btcolor"))
HrDetail=Request("Content")
zpdw=Request("zpdw")
zdfy=Request("zdfy")
dh=Request("dh")
dz=Request("dz")
emaill=Request("emaill")


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from job where id="&id
rs.open sql,conn,1,3
rs("lang")=lang
rs("HrName")=HrName
rs("HrRequireNum")=HrRequireNum
rs("HrAddress")=HrAddress
rs("HrSalary")=HrSalary
rs("ksDate")=ksDate
rs("jsDate")=jsDate
rs("zdfy")=zdfy
rs("emaill")=emaill
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=7 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=7  order by Sequence"
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
rs("HrDetail")=HrDetail
rs("zpdw")=zpdw
rs("dh")=dh
rs("dz")=dz
rs("emaill")=emaill
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
call htmll("/html/","",""&zpname&""&Separated&""&ID&""&Separated&"1.html",""&showjob&"/index.asp","jobid=",ID,"m=",mzp,"language=",lang,"page=",1,"","","","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&zpname&""&Separated&""&ID&""&Separated&""&i&".html",""&showjob&"/index.asp","jobid=",ID,"m=",mzp,"language=",lang,"page=",i,"","","","")
next
end if


rs.close
set rs=Nothing



call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&job&"/","",""&zpclass&".html",""&job&"/index.asp","m=",mzp,"language=",lang,"","","","","","","","")

end if

call addlog("修改"&zhaopin&"")
response.Write "<script language=javascript>alert(""信息修改成功"");window.location.href=""zp.asp""</script>"
    response.End
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From job where id="&id&" and lang='"&lang&"'", conn,1,3
	If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""zp.asp""</script>"
    response.End
End if
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
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改<%=zhaopin%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="zhaopin.asp" target="main">返回菜单</a></div>
 
 <div class="dmeun"><a href="zp.asp" target="main"><%=zhaopin%>管理</a></div>
 </td>
 

 </tr>
</table>      

                <table width="99%" border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
				 
				   <tr> 
                    <td  > 
                     用人单位:</td>
					 <td>
                     <input 
name="zpdw" type="text" id="zpdw" value="<%=rs("zpdw")%>" size="30" maxlength="30" datatype="*" nullmsg="单位不能为空"/></td>
                  </tr>
                  <tr> 
                    <td  > 
                    <%=zhaopin%>职位 :</td>
					 <td>
                      <input 
name="HrName" type="text" id="HrName" value="<%=rs("HrName")%>" size="30" maxlength="30" datatype="*" nullmsg="职位不能为空">                     </td>
                  </tr>
				  
				  <tr> 
                    <td  > 
                    标题颜色 :</td>
					 <td> <input name="btcolor"type="text" id="color"   value="<%=rs("btcolor")%>" size="8" /> 
 <input   type="button" id="colorpicker" value="打开取色器" /> </td>
				  
				  </tr>
				  
                  <tr> 
                    <td > 
                    <%=zhaopin%>人数：</td>
					 <td>
                      <input 
name="HrRequireNum" type="text" id="HrRequireNum" value="<%=rs("HrRequireNum")%>" size="5" maxlength="30" datatype="*" nullmsg="人数不能为空">
                    人</td>
                  </tr>
                  <tr> 
                    <td > 
                    工作地点:</td>
					 <td> 
                      <input 
name="HrAddress" type="text" id="HrAddress" value="<%=rs("HrAddress")%>" size="50" maxlength="60" datatype="*" nullmsg="工作地点不能为空">                     </td>
                  </tr>
                  <tr> 
                    <td > 
                    工资待遇:</td>
					 <td> 
                      <input 
name="HrSalary" type="text" id="Daiy" value="<%=rs("HrSalary")%>" size="20" maxlength="50" datatype="*" nullmsg="工资待遇不能为空">                    </td>
                  </tr>
				    <tr> 
                    <td > 
                      联系电话:</td>
					 <td> 
                      <input 
name="dh" type="text" id="Daiy" value="<%=rs("dh")%>" size="20" maxlength="50" datatype="*" nullmsg="联系电话不能为空">                    </td>
                  </tr>
				    <tr> 
                    <td > 
                      联系地址 :</td>
					 <td>
                      <input 
name="dz" type="text" id="Daiy" value="<%=rs("dz")%>" size="20" maxlength="50" datatype="*" nullmsg="联系地址不能为空">                    </td>
                  </tr>
				    <tr> 
                    <td > 
                      投递邮箱 :</td>
					 <td>
                      <input 
name="emaill" type="text" id="Daiy" value="<%=rs("emaill")%>" size="20" maxlength="50" datatype="*" nullmsg="邮箱不能为空">                    </td>
                  </tr>
                  <tr> 
                    <td > 
                    开始日期:</td>
					 <td>
                    <input name="ksdate"type="text" id="ksdate"  onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})"
 value="<%=rs("ksdate")%>" /></td>
                  </tr>
                   <tr> 
                    <td > 
                     结束日期:</td>
					 <td>
                     <input name="jsDate" type="text" class="textfield" id="jsDate" 
 value="<%=rs("jsdate")%>" onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})" /></td>
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
                    <td > 
                    <%=zhaopin%>要求:</td>
					 <td>
                    <textarea name="content" style="width:100%;height:380px;visibility:hidden;" datatype="*" nullmsg="<%=zhaopin%>要求不能为空"><%=rs("HrDetail")%></textarea></td>
                  </tr>
				  <tr>
      <td >手动分页：</td>
			   <td><INPUT type=button class="bnt" name="appendHtml" value="插入分页符">
			   </td>
          </tr>
        
     
        <input name="zdfy" type="hidden" id="zdfy"  value="<%=rs("zdfy")%>" size="10" />
       
				  				  
				  	<%'自定义字段%>
<%call showColumn("job",7,id)%>
                  <tr> 
                    <td colspan="2"> 
                       
                        <input name="submit" type="submit" value="修改" class="bnt">
                                      </td>
                  </tr>
              </table>
            
          </form>
</body>
</html>
