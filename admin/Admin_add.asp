<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|02,")=0 then 
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

<script language="JavaScript" src="../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
	 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>






<script language="javascript">
<!--


function form_onsubmit()
{
if (document.form_admin.Password.value!=document.form_admin.ConPassword.value)
{
alert ("新密码和确认密码不一致。");
document.form_admin.Password.value='';
document.form_admin.ConPassword.value='';
document.form_admin.Password.focus();
return false;
}
}
//-->
</SCRIPT>

</head>

<body>

<FORM name="form_admin" method="post"  onsubmit="return form_onsubmit()" action="Admin_ManageSave.asp" >	



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">添加管理员账号</div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="zhanghao.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="admin_manage.asp" target="main">管理员账号管理</a></div></td>
 

 </tr>
</table>





            <table width="99%"  align="center"  class="dtable4" cellpadding="0" cellspacing="0">

                <tr > 
                  <td width="11%" > 管理员帐号：</td>    
                  <td width="89%">               <input name="UserName" type="text" id="UserName" size="16" maxlength="20"
></td>
                </tr>
                <tr >
                  <td >是否生效：</td>    <td> 
                  <label>
                  <input name="working" type="checkbox" id="working" value="1" class="none"/>
                  </label></td>
                </tr>
                <tr > 
                  <td > 管理员密码： </td>    <td>                   <input name="Password" type="password" size="16" maxlength="20"
></td>
                </tr>
                <tr > 
                  <td > 密码确认：</td>    <td>                    <input name="ConPassword" type="password" size="16" maxlength="20"
></td>
                </tr>
                <tr >
                  <td >管理员名:</td>    <td> 
                    <label>
                    <input name="glyname" type="text" id="glyname" />
                  </label></td>
                </tr>
                <tr >
                  <td >操作权限: </td>    <td> 
				  基本设置
				  ：
				    <input name="Purview91" type="checkbox" id="Purview91" value="|91," />
					
                  基本设置
				  
				  <input name="Purview101" type="checkbox" id="Purview101" value="|101," />
                  导航管理
				  
				  <input name="Purview102" type="checkbox" id="Purview102" value="|102," />
                  导航参数
				  
				    <input name="Purview103" type="checkbox" id="Purview103" value="|103," />
                  分享设置
				  
				   
				  <input name="Purview81" type="checkbox" id="Purview81" value="|81," />
                  留言管理
				  
				  <input name="Purview10411" type="checkbox" id="Purview10411" value="|10411," />
                  咨询管理
				  
				  <input name="Purview10412" type="checkbox" id="Purview10412" value="|10412," />
				  邮件订阅管理
				  
				  <input name="Purview10413" type="checkbox" id="Purview10413" value="|10413," />
				  邮件群发管理
				  
				  <input name="Purview10414" type="checkbox" id="Purview10414" value="|10414," />
                  热门搜索
				  
				
				  
				  <input name="Purview104" type="checkbox" id="Purview104" value="|104," />
                  模板管理 
				  
				  <input name="Purview9999" type="checkbox" id="Purview9999" value="|9999," />
                  微信功能
				  
				  
				  
				  
				  <br /><br />
				    wap 设置：
				  
				   <input name="Purview615" type="checkbox" id="Purview615" value="|615," />
				   wap设置<br /><br />
				  
				  单页管理：
				  <input name="Purview11" type="checkbox" id="Purview11" value="|11," />
                  单页分类
                  <input name="Purview12" type="checkbox" id="Purview12" value="|12," />
                  单页增加
                  <input name="Purview13" type="checkbox" id="Purview13" value="|13," />
                  单页管理 <br /><br />
				  
				  新闻管理：
				  <input name="Purview21" type="checkbox" id="Purview21" value="|21," />
                  新闻分类
                  <input name="Purview22" type="checkbox" id="Purview22" value="|22," />
                  新闻增加
                  <input name="Purview23" type="checkbox" id="Purview23" value="|23," />
                  新闻管理  
				  <input name="Purview24" type="checkbox" id="Purview24" value="|24," />
                  新闻幻灯参数
				  
				  <br /><br />
				  
				  产品管理：
				   <input name="Purview31" type="checkbox" id="Purview31" value="|31," />
                  产品分类
                  <input name="Purview32" type="checkbox" id="Purview32" value="|32," />
                  产品增加
                  <input name="Purview33" type="checkbox" id="Purview33" value="|33," />
                  产品管理  
				  <input name="Purview34" type="checkbox" id="Purview34" value="|34," />
                  订单管理
				  <input name="Purview35" type="checkbox" id="Purview35" value="|35," />
                  产品幻灯参数
				  <input name="Purview36" type="checkbox" id="Purview36" value="|36," />
                  产品附属属性属性管理
				  
				  
				  
				   <br /><br />
				  
				  图片管理：
				     <input name="Purview41" type="checkbox" id="Purview41" value="|41," />
                  图片分类
                  <input name="Purview42" type="checkbox" id="Purview42" value="|42," />
                  图片增加
                  <input name="Purview43" type="checkbox" id="Purview43" value="|43," />
                  图片管理  
				  <input name="Purview44" type="checkbox" id="Purview44" value="|44," />
                  图片幻灯参数
				 
				  
				  <br /><br />
				  
				  视频管理：
				  <input name="Purview51" type="checkbox" id="Purview51" value="|51," />
                  视频分类
                  <input name="Purview52" type="checkbox" id="Purview52" value="|52," />
                  视频增加
                  <input name="Purview53" type="checkbox" id="Purview53" value="|53," />
                  视频管理  <br /><br />
				  
				  下载管理：
				  <input name="Purview61" type="checkbox" id="Purview61" value="|61," /> 
                  下载分类
                  <input name="Purview62" type="checkbox" id="Purview62" value="|62," />
                  下载增加
                  <input name="Purview63" type="checkbox" id="Purview63" value="|63," />
                  下载管理  <br /><br />
				  
				  招聘管理：
				    <input name="Purview71" type="checkbox" id="Purview71" value="|71," />
                  发布招聘
                  <input name="Purview72" type="checkbox" id="Purview72" value="|72," />
                  招聘管理
                  <input name="Purview73" type="checkbox" id="Purview73" value="|73," />
                  应聘管理 <br /><br />
                  
				  配送方式：
				  
				 <input name="Purview511" type="checkbox" id="Purview511" value="|511," /> 
                  配送方式管理
				  <input name="Purview512" type="checkbox" id="Purview512" value="|512," /> 
                  配送方式增加
				  <input name="Purview513" type="checkbox" id="Purview513" value="|513," /> 
                  配送方式修改
				  <input name="Purview514" type="checkbox" id="Purview514" value="|514," /> 
                  配送方式删除<br /><br />
				  
				  支付配置：
				  
				   <input name="Purview515" type="checkbox" id="Purview515" value="|515," /> 
                  支付配置 <br /><br />
				  标签样式设置：
				  
				   <input name="Purview611" type="checkbox" id="Purview611" value="|611," /> 
                  标签样式管理
				  <input name="Purview612" type="checkbox" id="Purview612" value="|612," /> 
                  标签样式增加
				  <input name="Purview613" type="checkbox" id="Purview613" value="|613," /> 
                  标签样式修改
				  <input name="Purview614" type="checkbox" id="Purview614" value="|614," /> 
                  标签样式删除<br /><br />
				  
				
				  
				  
				  
				  
				  
				  
                  
                  自定义字段：
				    <input name="Purview10415" type="checkbox" id="Purview10415" value="|10415," />
                  自定义字段管理
                  <br /><br />
                  
                  
                  
                  
                  
                  
                  
                  
				  
				  其他管理：
				  <input name="Purview110" type="checkbox" id="Purview110" value="|110," />
                 客服管理添加以及参数设置
				  <input name="Purview111" type="checkbox" id="Purview111" value="|111," />
                 幻灯管理添加以及参数设置
				 <input name="Purview112" type="checkbox" id="Purview112" value="|112," />
                 广告管理 <br /><br />
				 
				 安全设置：
				 
				 <input name="Purview109" type="checkbox" id="Purview109" value="|109," />
                 SQL注入查看以及设置 
				 <input name="Purview1091" type="checkbox" id="Purview1091" value="|1091," />
                 后台目录以及数据库名称修改
				 
				 <br /><br />
				 
				 
				 
				  
				  插件管理：
                  
                 <input name="Purview1000003" type="checkbox" id="Purview1000003" value="|1000003," />
                 营销区域管理
                 <input name="Purview1000004" type="checkbox" id="Purview1000004" value="|1000004," />
                 营销区域增加
                  
				 <input name="Purview201" type="checkbox" id="Purview201" value="|201," />
                 营销网络管理
				 <input name="Purview202" type="checkbox" id="Purview202" value="|202," />
                 营销网络添加
				   <input name="Purview105" type="checkbox" id="Purview105" value="|105," />
                 站内链接管理
				  <input name="Purview106" type="checkbox" id="Purview106" value="|106," />
                 站内链接添加
				 <input name="Purview107" type="checkbox" id="Purview107" value="|107," />
                自定义表单管理
                <!--
				  <input name="Purview108" type="checkbox" id="Purview108" value="|108," />
                 自定义表单添加
                 -->
				 <input name="Purview1081" type="checkbox" id="Purview1081" value="|1081," />
                 投票管理以及增加
				  
				  <input name="Purview113" type="checkbox" id="Purview113" value="|113," />
                 链接管理
				 <input name="Purview1131" type="checkbox" id="Purview1131" value="|1131," />
                 链接添加  
				 <input name="Purview1132" type="checkbox" id="Purview1131" value="|1132," />
                 链接分类管理
				 <br /><br />
				  
				  数据库管理：
				  <input name="Purview301" type="checkbox" id="Purview301" value="|301," />
                 数据库备份压缩恢复 <br /><br />
				  
				  会员管理：
				  <input name="Purview1000" type="checkbox" id="Purview1000" value="|1000," />
                  会员组别管理
                  <input name="Purview1001" type="checkbox" id="Purview1001" value="|1001," />
                  会员组别增加
                   <input name="Purview1002" type="checkbox" id="Purview1002" value="|1002," />
                  会员管理  <br /><br />
				  
				  
				  
				  
				  
				  管理员管理：
                 <input name="Purview01" type="checkbox" id="Purview01" value="|01," />
                  管理员管理
                  <input name="Purview02" type="checkbox" id="Purview02" value="|02," />
                  管理员增加
                   <input name="Purview03" type="checkbox" id="Purview03" value="|03," />
                  管理员登录日志  <br /><br />
                  
                  
                  多语言配置管理:
                  
                   <input name="Purview1000001" type="checkbox" id="Purview1000001" value="|1000001," />
                  多语言配置管理
                  <input name="Purview1000002" type="checkbox" id="Purview1000002" value="|1000002," />
                  多语言配置增加 <br /><br />
                  
                  
                  
                  
				   静态生成：
                <input name="Purview114" type="checkbox" id="Purview114" value="|114," />
                 生成html	   <br /><br />
				 
				 
				 
				 			  </td>
                </tr>
                
                <tr >
                  <td >备注说明:</td>    <td> 
                    <label>
                    <textarea name="Explain" cols="50" rows="6" id="Explain"></textarea>
                  </label></td>
                </tr>
                
                <tr > 
                  <td colspan="2">
                    <input type=submit value='确认添加' name=Submit2 class="bnt" />                 </td>
                </tr>
         </table>
</form>

</body>
</html>
