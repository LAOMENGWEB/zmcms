<!--#include file="../base.asp"-->
<!--#include file="../../lib/db.asp"-->
<!--#include file="../../lib/function.asp"-->
<!--#include file="../../lib/weixin_class.asp"-->
<!--#include file="../../lib/md5.asp"-->



<%
call dbopen()
If request.Form("submit") = "确定" Then
	p_Title=request.form("p_Title")
		p_ClassId=request.form("p_ClassId")
		p_Info=request.form("p_Info")
		p_Pic=request.form("p_Pic")
		p_Type=request.form("p_Type")
		p_Content=request.form("p_Content")
		p_UrlType=request.form("p_UrlType")
			If p_UrlType="外链" then
			p_Url=request.form("p_Url")
		else
			p_Url=request.form("p_Plug_Url")
		end if
		ImagePath=request.form("ImagePath")

Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From Plug_PicMsg ",conn2,1,3
rs.addnew
rs("p_Title")=p_Title
rs("p_ClassId")=p_ClassId
rs("p_Info")=p_Info
rs("p_Pic")=p_Pic
rs("p_Type")=p_Type
rs("p_url")=p_url
rs("p_content")=p_content
rs("p_UrlType")=p_UrlType
rs("ImagePath")=ImagePath
rs("p_addtime")=now()
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""admin.asp""</script>"
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="../../Css/wx.css" rel="stylesheet" type="text/css" />

<script language="JavaScript" src="../../../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../../../js/Validform_v5.3.2_min.js"></script>
<script type="text/javascript" src="../../../javascript/colorbox/jquery.colorbox-min.js"></script>
<link rel="stylesheet" href="../../../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../../../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../../../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../../../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../../../zmbjq/plugins/code/prettify.js"></script>
<link href="../../../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">

<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>

 <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>

<script>
 KindEditor.ready(function(K) {
 var imgh="";
 var imgurl="";
 var editor = K.editor({
 allowFileManager : true,
 uploadJson : '../../../zmbjq/asp/upload_json.asp'
 });
 K('#J_selectImage').click(function() {
 editor.loadPlugin('multiimage', function() {
 editor.plugin.multiImageDialog({
 clickFn : function(urlList) {
 var div = K('#J_imageView');
 div.html('');
 K.each(urlList, function(i, data) {
 var j=document.getElementById('ImagePath').value;
 if(j==""){
 imgurl=data.url; 
 }
 else{
 imgurl=j+"|"+data.url;
 }
 $("#ImagePath").val(imgurl);  
 });
 editor.hideDialog();
 var array = imgurl.split("|");
 var ids;
                                 for (var s=0 ; s< array.length ; s++)
                                 {
                                  ids=array[s];
                                  ids=ids.substr(ids.lastIndexOf("/")+1,ids.lastIndexOf("."));
                                  ids=ids.replace(".",""); 
                                     $('#J_imageView').append('<div class="imgDiv" id="'+ids+'"><a href="javascript:setContent(\'' + array[s] + '\')"><img title="点击图片将图片插入到编辑器" src="'+array[s] + '" width="120" height="90" border="0" /></a><br><a href="javascript:dropThisDiv(\''+ids+'\',\''+array[s] +'\')">删除</a></div>');
                                 }
 }
 });
 });
 });
 });

 //设置封面图片
function setIndexImage(Value)
  {
    $("#IcoImage").val(Value);
 //$("#picView").html("<img id='img_view' src='"+t+"' width='160'/>");
  }
 //向编辑器插入图片 
function setContent(Value)    
  {
     if (Value != '' && Value != null) {
   var valueimg='<img src="'+Value+'" />';
   editor.insertHtml(''+valueimg+''); //利用kindeditor里的insertHtml方法向光标所在处添加html内容
    }
 }
 //向编辑器插入分页标签
function SetEditorPage(Value)    
  {
     if (Value != '' && Value != null) {
   editor.insertHtml(''+Value+''); //利用kindeditor里的insertHtml方法向光标所在处添加html内容
    }
 }
     //删除图片
    function dropThisDiv(t,p)
     {
         document.getElementById(t).style.display='none'
         var str =document.getElementById("ImagePath").value;   //取上传了的图片内容
var strs =document.getElementById("IcoImage").value; //取缩略图文本框的内容
//利用kindeditor里的内置API方法，取内容编辑器的内容,开始
        var strcon = editor.html(); //editor为页面添加编辑器里JS里定义的那个值
        editor.sync();//同步到editor这个编辑器
        strcon = document.getElementById('p_Content').value; // 原生API，content为页面上textarea的ID值
//editor.html('HTML内容');设置编辑器的值
        //取内容编辑器的内容,结束
        var arr = str.split("|");
         var nstr="";
         for (var i=0; i<arr.length; i++)
         {
             if(arr[i]!=p)
             {
            if (nstr!="")
            {
            nstr=nstr+"|";
            } 
            nstr=nstr+arr[i]
             }
         }        
         if (nstr=="")
         {
             $("#imgTable").hide();
         }
 //判断要删除的图片是否被设置成了封面，如果是就清空，如果不是就不管,开始
if (strs==p){
 document.getElementById("IcoImage").value=""; //原来为""
 }
else
 {
 document.getElementById("IcoImage").value=strs; //原来为strs
 }
 //判断要删除的图片是否被设置成了封面，如果是就清空，如果不是就不管，结束
        document.getElementById("ImagePath").value=nstr;

        if (confirm("你确定从服务上一起删除这个图片吗？")){ //弹出提示框，点确定后执行删除图片代码
        //利用ajax调用deltu.asp页面进行删除图片
           var myDate = new Date();
             $.ajax({
                 url: "deltu.asp?id="+p+"",
                 data: { datetime: myDate.getTime() },
                 dataType: "text"
             });
        //删除图片结束
  //同时删除编辑器里插入的相应图片，开始
      //正则替换开始
      var regimgb='<img src="'+p+'" />';
       var regconb=new RegExp(""+regimgb+"","g"); //创建正则RegExp对象  
            var stringObjconb=""+strcon+"";  
            var newstrconb=stringObjconb.replace(regconb,"");   
   //正则替换结束
  editor.html(''+newstrconb+'');  //利用kindeditor里的内置API方法替换掉原来的编辑器内的内容
  //document.getElementById("IndexImage").value=stringObjcona; //调试用的
  //同时删除编辑器里插入的相应图片，结束 
        }  
     }















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
 	   <script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="p_Content"]', {
 filterMode : false, 
 uploadJson : '../../../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../../../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
 afterBlur: function(){this.sync();}
 
 });

 });

 </script>

</head>

<body>
   <form method="post" name="myform" class="checkform">
	
	
		<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >添加单图文素材</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 

 <div class="dmeun"><a href="Admin.asp" target="main" style="color:#ffffff">素材管理</a></div>
 </td>
 

 </tr>
</table>	
	
	
	<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >图文标题：</td>
<td width="88%">
 <input type="text" name="p_Title" size="50" datatype="*" nullmsg="图文标题不能为空"/>
</td>
 </tr>
 <tr > 
<td width="10%"   >所属分类：</td>
<td width="88%">
 <select name="p_ClassId" id="p_ClassId" datatype="*" nullmsg="所属分类不能为空">
  <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from Plug_PicMsg_class",conn2,1,1
		  Response.Write("<option value="""">请选择分类</option>") 
    If Rsc.Eof and Rsc.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
    Else
		  while not rsc.eof
			response.Write("<option value=""" & rsc("c_id") & """>" & rsc("c_classname") & "</option>")
			rsc.movenext
		wend
		end if
		rsc.close
		set rsc=nothing
		  %></select> 
  本项只作为后台备注分类，不影响正常使用
</td>
 </tr>
 
 
 <tr > 
<td width="10%"   >描述：</td>
<td width="88%">
<textarea name="p_Info" cols="60" rows="5" datatype="*" nullmsg="描述不能为空"></textarea>
</td>
 </tr>
 
 <tr > 
<td width="10%"   >素材类型：</td>
<td width="88%">
<select name="p_Type" id="p_Type" onChange="setclose(this.value)">
          <option value="图文">图文</option>
          <option value="链接">链接</option>
        </select>
        【链接】点击时直接链接到外部地址，【图文】直接显示内容
</td>
 </tr> 
 
 
    <tr><td>链接类型：</td>
	<td>
        <select name="p_UrlType" id="p_UrlType" >
          <option value="外链">外部链接</option>
          <!--<option value="插件">内部插件</option>-->

        </select>        
        
        <input type="text" name="p_Url" id="p_Url" size="60" value="http://"/>  当素材类型为 链接时  此性内容必须填写
		</td>
        </tr>
		    <td >缩略图：
				   <input name="p_Pic"  type="hidden" id="weepic"  value="" size="50"  />   		        </td>
                <td colspan="2"> <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="dtup.jpg"  border="0"/></a>
		        <div id="mysmaimg"></div></td>
	      </tr>
 
   <tr > 
                    <td> 图文内容: </br></td>
                    <td  colspan="2"><textarea  name="p_Content" id="p_Content" style="width:700px;height:300px;visibility:hidden;" datatype="*" nullmsg="图片内容不能为空"> </textarea>
					
					<br />
					当素材类型为图文时有效
					
					</td>
				  </tr>
	
                 
                  <tr  >
                  <td><input type="hidden" id="IcoImage" name="IcoImage" value="" style="width:600px;" />
        批量上传：         </td>
                  <td  colspan="2"><input name="button" type="button" id="J_selectImage" style=" background:url(allimg.png) no-repeat; width:170px;height:34px;border:none" value="" />
                    <input type="hidden" id="ImagePath" name="ImagePath" value="" style="width:600px;" /> 
                    小提示：批量上传图片以后可点击图片插入到编辑框
                      <div id="J_imageView">  </div></td>
                </tr>
 
 
 
 
 
 
 
    
 <tr > 
            <td colspan="2"> <input class="bnt" name="submit" type="submit" value="确定">
			  <input type="button" value="返回" onClick="location.href='javascript:history.go(-1)'" class="bnt"/>
            　            </td>
          </tr>	 
 
	
	
</table>	
	
	
	
	
	
	


    </form>











</body>
</html>
