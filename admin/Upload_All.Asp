<!-- #Include File="conn.asp" -->

<%

		
add=trim(request.QueryString("add"))
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title>zmcms</title>
<script type="text/javascript" src="swfupload/swfupload.js"></script>
<script type="text/javascript" src="swfupload/multiupload.js"></script>
<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<script type="text/javascript">
var Setting={
	debug:false,
	upload_url: "upload.asp",
	file_post_name : "filedata",
	file_types : "*.jpg;*.gif;*.png",//文件格式限制
	file_types_description : "文件类型",//文件格式描述
	file_size_limit : "100 MB",	// 文件大小限制
	file_upload_limit : 20//上传文件总数
};
function CallBack(){
	document.getElementById("fileUploaded").innerHTML+="<li><img src='../"+ this.name+"'/></li>";
	document.getElementById("call").value +=this.name+",";
}
function callbox(){
		 var value = $('#call').val();
	
    value=value.substring(0,value.length-1);
		var reg=new RegExp(",","g"); //创建正则RegExp对象  
		var stringObj=value; 
		var newstr=stringObj.replace(reg,'"/></li><li><img src="../');
		var callimg='<li><img src="../'+newstr +'"/></li>' 
		var url=stringObj.replace(reg,','); //URL值数组符号为,可以换成任意值
		var callurl=url;
	          parent.$('#bigimage').val(callurl);		 
            parent.$("#mypic").html(callimg);		
	parent.$.fn.colorbox.close();return false;
}
<%if add="one" then%>
<%else%>
window.onload=pageInit;

<%end if%>
</script>
<style>
body,div,ul,li{ font-size:12px; font-family:Arial, Helvetica, sans-serif; padding:0; margin:0; list-style:none}
.clearfix:after {content: ".";display: block;height: 0;clear: both;visibility: hidden;}
.clearfix {display: inline-block;} 
* html .clearfix {height: 1%;} 
.clearfix {display: block;} 
.clear{clear:both}
.divAddFiles1{ width:80px; height:25px; line-height:25px;border:#F1AA2D 1px solid; background-color:#FEFFC8; cursor:pointer; float:left}
#buttonArea{
	background:#FFF;
	border-bottom:1px solid #DDD;
	padding:5px; float:left; width:940px
}
#buttonArea li { float:left;display:block; margin-left:15px}
#buttonArea li input { height:27px}
#progressArea{
	background:#FFF;
	padding:3px;
	padding-left:5px;
	height:15px;
	width:850px;
	float:left;
}
.actButton{ float:left; width:90; margin:0 20px}
#upload{ width:950px;border:1px #DDD solid; margin:10px}

#listArea{ padding:0;}
.filelist{padding:5px; margin:5px; border:1px  #739F1D  solid; width:214px; float:left; background:#D4F684}
.filelist .filename{ font-size:13px; font-weight:bold; background: #34820B; padding:2px; color:#FFF}
.filelist .fileprocess{margin:5px 0 0 0;}

#progressArea ul li{ float:left;}
#processer{width:950px; padding-top:3px}
#progressArea ul{ padding-top:3px}
#progressArea ul li.u-text{ width:40px}
#fileUploaded { padding:5px}
#fileUploaded img { width:83px; height:60px;border: #CCC 1px solid; padding:2px; display:block}
#fileUploaded li { float:left; width:87px;margin:3px; }
fieldset{border:1px solid #ddd;padding:0px 15px 15px 15px;}
legend{font-weight:bold;font-size:1.2em; padding:10px 5px; line-height:1.8em}
label{margin:2px; line-height:30px; padding-left:10px;}
</style>
</head>
<body>
<% If add="one" Then %>
 <fieldset>
			<legend>请上传文件</legend>
        <form action="upload.asp?add=one" method="post" enctype="multipart/form-data">
        <div id="sfile">
         <input name="file1" type="file" id="fuImages" size="20" />
          <input type="submit" class="submit4" value="上传文件" /></div>
       
       </form>
       </fieldset>
<% Else %>
<input id="call" type="hidden" value="" />
	<div id="upload">
		<div id="buttonArea">
        	<div class="divAddFiles1">
        	<span id="divAddFiles"></span>
            </div>
             
            <li><input class="divAddFiles1" type="button" id="startU" disabled="disabled" value="开始上传" onclick="startUploadFiles();" /></li>
            <li><input class="divAddFiles1" type="button" value="插入内容" onclick="callbox();" /></li>
            
		</div>
		<div id="listArea"></div>
        <div style="padding:3px" class="clearfix">
        	
            <div id="progressArea">
                <ul>
                    <li id="processer"></li>
                    <li class="all-text"></li>
                    <li class="clear"></li>
                </ul>
        	</div>
            <div class="clear">
       </div>
     <div class="clear" style="border-bottom:#CCC 1px dotted;  padding:5px">已上传文件：</div>
    <div id="fileUploaded"></div>
      <div class="clear"></div>
	</div>
	</div>
  <% End If %>

</body>
</html>