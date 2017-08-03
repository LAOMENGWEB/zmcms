<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<style type="text/css">
*{	margin: 0;padding: 0;border: 0;outline: 0;font-size: 100%;vertical-align: baseline;background: transparent;}
.clear { clear: both; }
body {
background-color:#284174;font-family: "微软雅黑",Arial, Helvetica;font-size: 12px;color: #646464;text-align: center;}
#login {margin:80px auto; width:400px; height:440px; background-image:url(../images/login_bg.jpg); border:3px  #ffffff solid}
.top {width:400px; float:left;  height:80px;line-height:80px; font-size:20px; color:#ffffff;background-color:#284174;}
.bottom {width:400px; float:left;  height:60px;line-height:60px; font-size:1.2em;font-weight: 400;color:#ffffff;background-color:#284174; margin-top:70px;}
.bottom	a{color:#ffffff;text-decoration:none;}
.right {width:400px; float:left; margin-top:0px;  height:300px;}
.loginform {margin-top:40px;text-align:left; padding:0 20px; position:relative;}
.loginform p {height:70px;}
.loginform label {width:70px; display:block; float:left; line-height:45px;color: #B8B8B8;font-size:16px; }
.loginform .wbk {	border: 2px ridge rgba(187, 185, 189, 0.11);border-radius: 0.3em;-webkit-border-radius:0.3em;-moz-border-radius:0.3em;
-o-border-radius:0.3em;list-style:none;margin-bottom:12px;background:#F0EEF0;height:40px;line-height:40px;padding-left:20px;color: #B8B8B8;font-size:20px;}
.loginform .anniu {	border: 2px ridge rgba(187, 185, 189, 0.11);border-radius: 0.3em;-webkit-border-radius:0.3em;-moz-border-radius:0.3em;-o-border-radius:0.3em;list-style:none;margin-bottom:12px;background:#F0EEF0;height:40px;line-height:40px;font-weight:bold;background:#284174;width:150px;height:50px;border:none;color:#ffffff;font-size:20px;}
.loginform input.long { width:260px;}
.loginform input.small { width:100px;}
.loginform .error { background-color:#284174; width:210px;padding:3px 0 5px; text-align:center;color:#fff; height:24px; line-height:24px;  display:none; margin:0 auto}
input,img {vertical-align:middle;}/*文本框和验证码对齐*/
</style>
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<script language="javascript">
function RefreshImage(valImageId) {
	var objImage = document.images[valImageId];
	if (objImage == undefined) {
		return;
	}
	var now = new Date();
	objImage.src = objImage.src.split('?')[0] + '?x=' + now.toUTCString();
}
</script>
<script language="javascript" type="text/javascript">
function cklogin(){
	$('.error').html('&nbsp;&nbsp;正在登陆...!');
	$('.error').show(300);	
	if ($('#username').val()==''){
		$('#username').focus();
		$('.error').html('&nbsp;&nbsp;用户名不能为空!');
		return;
	}
	if ($('#password').val()==''){
		$('#password').focus();
		$('.error').html('&nbsp;&nbsp;密码不能为空!');
		return;
	}
	if ($('#captchacode').val()==''){
		$('#captchacode').focus();
		$('.error').html('&nbsp;&nbsp;验证码不能为空!');
		return;
	}
	$.ajax({type:'POST',
		   url:'Admin_ChkLogin.asp',
		   data:'username='+escape($('#username').val())+'&password='+escape($('#password').val())+'&captchacode='+escape($('#captchacode').val()),
		   success:function(msg){
			   if(msg=='0'){window.location='index.asp';}
			   else{$('.error').html('&nbsp;&nbsp;'+msg);$('.error').show();}
	
		   }});

}
function KeyDown()
{
    if (event.keyCode == 13)
    {
        event.returnValue=false;
        event.cancel = true;
        cklogin();
    }
}
</script>

</head>

<body>
<div id="login">
<div class="top">管理员登陆</div>
<div class="right">
<div class="loginform">
<form name="myform">
<p><label>用户名</label><input type="text" name="username" id="username" class="wbk long" /></p>
<p><label>密&nbsp;&nbsp;&nbsp;&nbsp;码</label><input type="password" class="wbk long" name="password" id="password" /></p>
<p><label>验证码</label><input type="text" name="captchacode" class="wbk small" id="captchacode" />
<a href="###" onclick="RefreshImage('imgCaptcha')"><img id="imgCaptcha" src="../inc/captcha.asp" height=40px style="margin-top:-10px"/></a></p>
<p> <a href="javascript:void(0);" onclick="cklogin();"><img src="images/dl.png" alt="登 陆" width="159" height="57" border="0" /></a>  <span style="padding-left:20px;">技术咨询QQ：2685863740</span>  </p>
</form>
 <div class="error">&nbsp;&nbsp;用户名错误!</div>
</div>
</div>

<div class="bottom">Copyright &copy; 版权归 <a href="http://www.zmxt.cn/" target="_blank">追梦工作室</a> 所有</div>
</div>
</body>
</html>
