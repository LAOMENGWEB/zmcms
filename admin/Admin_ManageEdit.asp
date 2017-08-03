<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/sha1.asp"-->
<!--#include file="inc/function.asp"-->
<%
If request.Form("submit") = "修改" Then

ID=request.QueryString("ID")
working=Request.Form("Working")
Explain=Request.Form("Explain")
glyname=Request.Form("glyname")
UserName=Request.Form("UserName")
OldPassword=replace(trim(Request("OldPassword")),"'","")
NewPassword=replace(trim(Request("NewPassword")),"'","")
set rs=server.createobject("adodb.recordset")
sql="select * from admin where Id="&id
rs.open sql,conn,1,3

if trim(Request.Form("oldPassword"))<>"" then
if strLength(NewPassword)>16 or strLength(NewPassword)<6 then
  response.write  "<script language=javascript>alert('请输入的密码位数不能小于6位或大于16位!');history.go(-1);</script>"
  response.End
end if
if md5(sha1(OldPassword&"zmcms"))<>rs("PassWord")  then
  response.write  "<script language=javascript>alert('原密码错误，请返回重新输入!');history.go(-1);</script>"
  response.End
else    
  rs("PassWord")=MD5(sha1(NewPassword&"zmcms"))
end if 
end if 


if working=1 then
        rs("Working")=1
	  else
        rs("Working")=0
	  end if
	    rs("Explain")=Explain
	  rs("glyname")=glyname
	  rs("UserName")=UserName
	  
	rs("AdminPurview")=Request.Form("Purview01") & Request.Form("Purview02") & Request.Form("Purview03") &_
	  
	                     Request.Form("Purview11") & Request.Form("Purview12") & Request.Form("Purview13") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") & Request.Form("Purview23") & Request.Form("Purview24") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") & Request.Form("Purview34") & Request.Form("Purview35") & Request.Form("Purview36") &_
						 Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") & Request.Form("Purview44") & Request.Form("Purview45") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
						 Request.Form("Purview511") & Request.Form("Purview512") & Request.Form("Purview513") & Request.Form("Purview514") & Request.Form("Purview515") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") & Request.Form("Purview63") &_
						 	 Request.Form("Purview611") & Request.Form("Purview612") & Request.Form("Purview613") & Request.Form("Purview614") & Request.Form("Purview615") &_
	                     Request.Form("Purview71") & Request.Form("Purview72") & Request.Form("Purview73") &_
	                     Request.Form("Purview81") &_
	                     Request.Form("Purview91") &_
	                     Request.Form("Purview101") & Request.Form("Purview102") &_
	                     Request.Form("Purview103") &_
	                     Request.Form("Purview104") & Request.Form("Purview10411") & Request.Form("Purview10412") & Request.Form("Purview10413") & Request.Form("Purview10414") & Request.Form("Purview10415") &_
	                     Request.Form("Purview105") &_
						 		 Request.Form("Purview1000001") & Request.Form("Purview1000002") & Request.Form("Purview1000003") & Request.Form("Purview1000004") &_
	                     Request.Form("Purview106") & Request.Form("Purview107") &_
	                     Request.Form("Purview108") & Request.Form("Purview109") & Request.Form("Purview1091") & Request.Form("Purview1081") &_
	                     Request.Form("Purview110") &_
						 Request.Form("Purview111") &_
	                    Request.Form("Purview112") &_
	                    Request.Form("Purview113") & Request.Form("Purview1131") & Request.Form("Purview1132") &_
	                    Request.Form("Purview114") &_
	                     Request.Form("Purview201") & Request.Form("Purview202") &_
						 Request.Form("Purview1000") & Request.Form("Purview1001") & Request.Form("Purview1002") &_
						 Request.Form("Purview9999") &_
						 Request.Form("Purview301")
rs.update
rs.close
set rs=nothing
call addlog("修改管理员账号")
 response.Write "<script language=javascript>alert(""修改成功"");window.location.href=""admin_manage.asp""</script>"
    response.End
end if

id=request("id")
set rs=server.createobject("adodb.recordset")
sql="select * from admin where id=" & id
rs.open sql,conn,1,1

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="JavaScript" src="../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<SCRIPT language=javascript>
<!--
function myform_onsubmit()
{
if (document.myform.NewPassword.value!=document.myform.ConPassword.value)
{
alert ("新密码和确认密码不一致。");
document.myform.NewPassword.value='';
document.myform.ConPassword.value='';
document.myform.NewPassword.focus();
return false;
}
}
//-->
</SCRIPT>
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
</head>

<body>

<FORM name=myform  onsubmit="return myform_onsubmit()"  method=post class="checkform">   
<input type=hidden name=id value=<%=id%>>        



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改管理员账号</div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="zhanghao.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="admin_manage.asp" target="main">管理员账号管理</a></div></td>
 

 </tr>
</table> 
           <table width="99%" align="center"  class="dtable4" cellpadding="0" cellspacing="0">
        
                <tr > 
                  <td width="9%"  > 管理员帐号： </td>    <td width="91%"> <input name="UserName" type="text" value="<%=rs("UserName")%>" /></td>
                </tr>
				  <tr > 
                  <td  > 管理员名：</td>    <td> 
                    <input name="glyname" type="text" value="<%=rs("glyname")%>" /></td>
                </tr>
                <tr >
                  <td  >生效:</td>    <td> 
                    <label>
                    <input name="working" type="checkbox" id="working" value="1" <%if rs("Working")=1 then response.write ("checked")%> class="none">
                  </label></td>
                </tr>
                <tr >
                  <td  >旧密码：</td>    <td> <input name="OldPassword" type="password" size="16" maxlength="20"
> 为空则不修改密码</td>
                </tr>
                <tr > 
                  <td  > 新密码：</td>    <td> <input name="NewPassword" type="password" size="16" maxlength="20"
></td>
                </tr>
                <tr > 
                  <td  > 密码确认：</td>    <td> <input name="ConPassword" type="password" size="16" maxlength="20"
></td>
                </tr>
                <tr >
                  <td >操作权限:</td>    <td> 
				  
				  
				  基本设置：
				    <input name="Purview91" type="checkbox" id="Purview91" value="|91,"  <% if Instr(rs("AdminPurview"),"|91,")>0 then response.write ("checked") end  if%>/>
                  基本设置
				  <input name="Purview101" type="checkbox" id="Purview101" value="|101,"  <% if Instr(rs("AdminPurview"),"|101,")>0 then response.write ("checked") end  if%>/>
                  导航管理
				  <input name="Purview102" type="checkbox" id="Purview102" value="|102,"  <% if Instr(rs("AdminPurview"),"|102,")>0 then response.write ("checked") end  if%>/>
                  导航参数
				    <input name="Purview103" type="checkbox" id="Purview103" value="|103,"  <% if Instr(rs("AdminPurview"),"|103,")>0 then response.write ("checked") end  if%>/>
                  分享设置
				   
				  <input name="Purview81" type="checkbox" id="Purview81" value="|81,"  <% if Instr(rs("AdminPurview"),"|81,")>0 then response.write ("checked") end  if%>/>
                  留言管理
				  
							  <input name="Purview10411" type="checkbox" id="Purview10411" value="|10411,"  <% if Instr(rs("AdminPurview"),"|10411,")>0 then response.write ("checked") end  if%>/>
                  咨询管理
				  
				  				  <input name="Purview10412" type="checkbox" id="Purview10412" value="|10412,"  <% if Instr(rs("AdminPurview"),"|10412,")>0 then response.write ("checked") end  if%>/>
                  邮件订阅管理
				  
								  <input name="Purview10413" type="checkbox" id="Purview10413" value="|10413,"  <% if Instr(rs("AdminPurview"),"|10413,")>0 then response.write ("checked") end  if%>/>
                  邮件群发 	  
				  
				  
				  
				  
				  
				  
				  
				  
				  <input name="Purview10414" type="checkbox" id="Purview10414" value="|10414,"  <% if Instr(rs("AdminPurview"),"|10414,")>0 then response.write ("checked") end  if%>/>
                  热门搜索
				  
				  <input name="Purview104" type="checkbox" id="Purview104" value="|104,"  <% if Instr(rs("AdminPurview"),"|104,")>0 then response.write ("checked") end  if%>/>
                  模板管理 
				  
				  
				    <input name="Purview9999" type="checkbox" id="Purview9999" value="|9999,"  <% if Instr(rs("AdminPurview"),"|9999,")>0 then response.write ("checked") end  if%>/>
                  微信功能 
				  
				  
				  
				  
				  <br /><br />
				      wap 设置：
				  
				   <input name="Purview615" type="checkbox" id="Purview615" value="|615,"  <% if Instr(rs("AdminPurview"),"|615,")>0 then response.write ("checked") end  if%>/>
				   wap设置<br /><br />
				  
				  单页管理：
				  <input name="Purview11" type="checkbox" id="Purview11" value="|11,"  <% if Instr(rs("AdminPurview"),"|11,")>0 then response.write ("checked") end  if%>/>
                  单页分类
                  <input name="Purview12" type="checkbox" id="Purview12" value="|12,"  <% if Instr(rs("AdminPurview"),"|12,")>0 then response.write ("checked") end  if%>/>
                  单页增加
                  <input name="Purview13" type="checkbox" id="Purview13" value="|13,"  <% if Instr(rs("AdminPurview"),"|13,")>0 then response.write ("checked") end  if%>/>
                  单页管理 <br /><br />
				  
				  新闻管理：
				  <input name="Purview21" type="checkbox" id="Purview21" value="|21,"  <% if Instr(rs("AdminPurview"),"|21,")>0 then response.write ("checked") end  if%>/>
                  新闻分类
                  <input name="Purview22" type="checkbox" id="Purview22" value="|22,"  <% if Instr(rs("AdminPurview"),"|22,")>0 then response.write ("checked") end  if%>/>
                  新闻增加
                  <input name="Purview23" type="checkbox" id="Purview23" value="|23,"  <% if Instr(rs("AdminPurview"),"|23,")>0 then response.write ("checked") end  if%>/>
                  新闻管理
				  
				     <input name="Purview24" type="checkbox" id="Purview24" value="|24,"  <% if Instr(rs("AdminPurview"),"|24,")>0 then response.write ("checked") end  if%>/>
                  新闻幻灯参数
				  
				  
				    <br /><br />
				  
				  产品管理：
				   <input name="Purview31" type="checkbox" id="Purview31" value="|31,"  <% if Instr(rs("AdminPurview"),"|31,")>0 then response.write ("checked") end  if%>/>
                  产品分类
                  <input name="Purview32" type="checkbox" id="Purview32" value="|32,"  <% if Instr(rs("AdminPurview"),"|32,")>0 then response.write ("checked") end  if%>/>
                  产品增加
                  <input name="Purview33" type="checkbox" id="Purview33" value="|33,"  <% if Instr(rs("AdminPurview"),"|33,")>0 then response.write ("checked") end  if%>/>
                  产品管理  
				  <input name="Purview34" type="checkbox" id="Purview34" value="|34,"  <% if Instr(rs("AdminPurview"),"|34,")>0 then response.write ("checked") end  if%>/>
                  订单管理 
				    <input name="Purview35" type="checkbox" id="Purview35" value="|35,"  <% if Instr(rs("AdminPurview"),"|35,")>0 then response.write ("checked") end  if%>/>
                  产品幻灯参数 
				  
				     <input name="Purview36" type="checkbox" id="Purview36" value="|36,"  <% if Instr(rs("AdminPurview"),"|36,")>0 then response.write ("checked") end  if%>/>
                  产品附属属性管理
				  
				  
				  
				  <br /><br />
				  
				  图片管理：
				     <input name="Purview41" type="checkbox" id="Purview41" value="|41,"  <% if Instr(rs("AdminPurview"),"|41,")>0 then response.write ("checked") end  if%>/>
                  图片分类
                  <input name="Purview42" type="checkbox" id="Purview42" value="|42,"  <% if Instr(rs("AdminPurview"),"|42,")>0 then response.write ("checked") end  if%>/>
                  图片增加
                  <input name="Purview43" type="checkbox" id="Purview43" value="|43,"  <% if Instr(rs("AdminPurview"),"|43,")>0 then response.write ("checked") end  if%>/>
                  图片管理
				  <input name="Purview44" type="checkbox" id="Purview44" value="|44,"  <% if Instr(rs("AdminPurview"),"|44,")>0 then response.write ("checked") end  if%>/>
                  图片幻灯参数 
		
				  
				  
				  
				    <br /><br />
				  
				  视频管理：
				  <input name="Purview51" type="checkbox" id="Purview51" value="|51,"  <% if Instr(rs("AdminPurview"),"|51,")>0 then response.write ("checked") end  if%>/>
                  视频分类
                  <input name="Purview52" type="checkbox" id="Purview52" value="|52,"  <% if Instr(rs("AdminPurview"),"|52,")>0 then response.write ("checked") end  if%>/>
                  视频增加
                  <input name="Purview53" type="checkbox" id="Purview53" value="|53,"  <% if Instr(rs("AdminPurview"),"|53,")>0 then response.write ("checked") end  if%>/>
                  视频管理  <br /><br />
				  
				  下载管理：
				  <input name="Purview61" type="checkbox" id="Purview61" value="|61,"  <% if Instr(rs("AdminPurview"),"|61,")>0 then response.write ("checked") end  if%>/> 
                  下载分类
                  <input name="Purview62" type="checkbox" id="Purview62" value="|62,"  <% if Instr(rs("AdminPurview"),"|62,")>0 then response.write ("checked") end  if%>/>
                  下载增加
                  <input name="Purview63" type="checkbox" id="Purview63" value="|63,"  <% if Instr(rs("AdminPurview"),"|63,")>0 then response.write ("checked") end  if%>/>
                  下载管理  <br /><br />
				  
				  招聘管理：
				    <input name="Purview71" type="checkbox" id="Purview71" value="|71,"  <% if Instr(rs("AdminPurview"),"|71,")>0 then response.write ("checked") end  if%>/>

                  发布招聘
                  <input name="Purview72" type="checkbox" id="Purview72" value="|72,"  <% if Instr(rs("AdminPurview"),"|72,")>0 then response.write ("checked") end  if%>/>
                  招聘管理
                  <input name="Purview73" type="checkbox" id="Purview73" value="|73,"  <% if Instr(rs("AdminPurview"),"|73,")>0 then response.write ("checked") end  if%>/>
                  应聘管理 <br /><br />
				  
				  
				  配送方式：
				  
				 <input name="Purview511" type="checkbox" id="Purview511" value="|511,"  <% if Instr(rs("AdminPurview"),"|511,")>0 then response.write ("checked") end  if%>/> 
                  配送方式管理
				  <input name="Purview512" type="checkbox" id="Purview512" value="|512,"  <% if Instr(rs("AdminPurview"),"|512,")>0 then response.write ("checked") end  if%>/> 
                  配送方式增加
				  <input name="Purview513" type="checkbox" id="Purview513" value="|513,"  <% if Instr(rs("AdminPurview"),"|513,")>0 then response.write ("checked") end  if%>/> 
                  配送方式修改
				  <input name="Purview514" type="checkbox" id="Purview514" value="|514,"  <% if Instr(rs("AdminPurview"),"|514,")>0 then response.write ("checked") end  if%>/> 
                  配送方式删除<br /><br />
				  
				  支付配置：
				  
				   <input name="Purview515" type="checkbox" id="Purview515" value="|515,"  <% if Instr(rs("AdminPurview"),"|515,")>0 then response.write ("checked") end  if%>/> 
                  支付配置 <br /><br />
				  标签样式设置：
				  
				   <input name="Purview611" type="checkbox" id="Purview611" value="|611,"  <% if Instr(rs("AdminPurview"),"|611,")>0 then response.write ("checked") end  if%>/> 
                  标签样式管理
				  <input name="Purview612" type="checkbox" id="Purview612" value="|612,"  <% if Instr(rs("AdminPurview"),"|612,")>0 then response.write ("checked") end  if%>/> 
                  标签样式增加
				  <input name="Purview613" type="checkbox" id="Purview613" value="|613,"  <% if Instr(rs("AdminPurview"),"|613,")>0 then response.write ("checked") end  if%>/> 
                  标签样式修改
				  <input name="Purview614" type="checkbox" id="Purview614" value="|614,"  <% if Instr(rs("AdminPurview"),"|614,")>0 then response.write ("checked") end  if%>/> 
                  标签样式删除<br /><br />
				  
				  
				  
				  
				  
				  
				  
				  
                  
                  自定义字段：
				    <input name="Purview10415" type="checkbox" id="Purview10415" value="|10415,"  <% if Instr(rs("AdminPurview"),"|10415,")>0 then response.write ("checked") end  if%>/>

                  自定义字段管理 <br /><br />
                  
                  
                  
                  
				  
				  其他管理：
				  <input name="Purview110" type="checkbox" id="Purview110" value="|110,"  <% if Instr(rs("AdminPurview"),"|110,")>0 then response.write ("checked") end  if%>/>
                 客服管理添加以及参数设置
				  <input name="Purview111" type="checkbox" id="Purview111" value="|111,"  <% if Instr(rs("AdminPurview"),"|111,")>0 then response.write ("checked") end  if%>/>
                 幻灯管理添加以及参数设置
				 <input name="Purview112" type="checkbox" id="Purview112" value="|112,"  <% if Instr(rs("AdminPurview"),"|112,")>0 then response.write ("checked") end  if%>/>
                 广告管理 <br /><br />
				 
				 安全设置：
				 
				 <input name="Purview109" type="checkbox" id="Purview109" value="|109,"  <% if Instr(rs("AdminPurview"),"|109,")>0 then response.write ("checked") end  if%>/>
                 SQL注入查看以及设置 
				 <input name="Purview1091" type="checkbox" id="Purview1091" value="|1091,"  <% if Instr(rs("AdminPurview"),"|1091,")>0 then response.write ("checked") end  if%>/>
                 后台目录以及数据库名称修改
				 
				 <br /><br />
				 
				 
				 
				  
				  插件管理：
                   <input name="Purview1000003" type="checkbox" id="Purview1000003" value="|1000003," <% if Instr(rs("AdminPurview"),"|1000003,")>0 then response.write ("checked") end  if%>/>
                 营销区域管理
                 <input name="Purview1000004" type="checkbox" id="Purview1000004" value="|1000004," <% if Instr(rs("AdminPurview"),"|1000004,")>0 then response.write ("checked") end  if%>/>
                 营销区域增加
				  <input name="Purview201" type="checkbox" id="Purview201" value="|201,"  <% if Instr(rs("AdminPurview"),"|201,")>0 then response.write ("checked") end  if%>/>
                 营销网络管理
				 <input name="Purview202" type="checkbox" id="Purview202" value="|202,"  <% if Instr(rs("AdminPurview"),"|202,")>0 then response.write ("checked") end  if%>/>
                 营销网络添加
				   <input name="Purview105" type="checkbox" id="Purview105" value="|105,"  <% if Instr(rs("AdminPurview"),"|105,")>0 then response.write ("checked") end  if%>/>
                 站内链接管理
				  <input name="Purview106" type="checkbox" id="Purview106" value="|106,"  <% if Instr(rs("AdminPurview"),"|106,")>0 then response.write ("checked") end  if%>/>
                 站内链接添加
				 <input name="Purview107" type="checkbox" id="Purview107" value="|107,"  <% if Instr(rs("AdminPurview"),"|107,")>0 then response.write ("checked") end  if%>/>
                自定义表单管理
                <!--
				  <input name="Purview108" type="checkbox" id="Purview108" value="|108,"  <% if Instr(rs("AdminPurview"),"|108,")>0 then response.write ("checked") end  if%>/>
                 自定义表单添加
                 -->
				 <input name="Purview1081" type="checkbox" id="Purview1081" value="|1081,"  <% if Instr(rs("AdminPurview"),"|1081,")>0 then response.write ("checked") end  if%>/>
                 投票管理以及增加
			 
				  <input name="Purview113" type="checkbox" id="Purview113" value="|113,"  <% if Instr(rs("AdminPurview"),"|113,")>0 then response.write ("checked") end  if%>/>
                 链接管理
				 <input name="Purview1131" type="checkbox" id="Purview1131" value="|1131,"  <% if Instr(rs("AdminPurview"),"|1131,")>0 then response.write ("checked") end  if%>/>
                 链接添加  
				 
				  <input name="Purview1132" type="checkbox" id="Purview1132" value="|1132,"  <% if Instr(rs("AdminPurview"),"|1132,")>0 then response.write ("checked") end  if%>/>
                 链接分类管理
				 
				 
				 <br /><br />
				  
				  数据库管理：
				  <input name="Purview301" type="checkbox" id="Purview301" value="|301,"  <% if Instr(rs("AdminPurview"),"|301,")>0 then response.write ("checked") end  if%>/>
                 数据库备份压缩恢复 <br /><br />
				  
				  会员管理：
				  <input name="Purview1000" type="checkbox" id="Purview1000" value="|1000,"  <% if Instr(rs("AdminPurview"),"|1000,")>0 then response.write ("checked") end  if%>/>
                  会员组别管理
                  <input name="Purview1001" type="checkbox" id="Purview1001" value="|1001,"  <% if Instr(rs("AdminPurview"),"|1001,")>0 then response.write ("checked") end  if%>/>
                  会员组别增加
                   <input name="Purview1002" type="checkbox" id="Purview1002" value="|1002,"  <% if Instr(rs("AdminPurview"),"|1002,")>0 then response.write ("checked") end  if%>/>
                  会员管理  <br /><br />
				  			  
				  管理员管理：
                 <input name="Purview01" type="checkbox" id="Purview01" value="|01,"  <% if Instr(rs("AdminPurview"),"|01,")>0 then response.write ("checked") end  if%>/>
                  管理员管理
                  <input name="Purview02" type="checkbox" id="Purview02" value="|02,"  <% if Instr(rs("AdminPurview"),"|02,")>0 then response.write ("checked") end  if%>/>
                  管理员增加
                   <input name="Purview03" type="checkbox" id="Purview03" value="|03,"  <% if Instr(rs("AdminPurview"),"|03,")>0 then response.write ("checked") end  if%>/>
                  管理员登录日志  <br /><br />
                     多语言配置管理:
                  
                   <input name="Purview1000001" type="checkbox" id="Purview1000001" value="|1000001," <% if Instr(rs("AdminPurview"),"|1000001,")>0 then response.write ("checked") end  if%>/>
                  多语言配置管理
                  <input name="Purview1000002" type="checkbox" id="Purview1000002" value="|1000002," <% if Instr(rs("AdminPurview"),"|1000002,")>0 then response.write ("checked") end  if%>/>
                  多语言配置增加 <br /><br />
                  
				   静态生成：
                <input name="Purview114" type="checkbox" id="Purview114" value="|114,"  <% if Instr(rs("AdminPurview"),"|114,")>0 then response.write ("checked") end  if%>/>
                 生成html	   <br /><br />

				 
				 </td>
                </tr>
				
                <tr >
                  <td >备注:</td>    <td> 
                    <label>
                    <textarea name="Explain" cols="50" rows="6" id="Explain"><%=rs("explain")%></textarea>
                    </label></td>
                </tr>
                <tr > 
                  <td  colspan="2">
                    <input class="bnt" type="submit" name="Submit" value="修改" >                 </td>
                </tr>
  </table>
           </td>
</form>
</body>
</html>
