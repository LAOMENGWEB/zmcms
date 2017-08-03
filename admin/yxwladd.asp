<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
if Instr(session("AdminPurview"),"|202,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%If request.Form("submit") = "确定" Then%>
<%
province=Trim(request.form("province")) 
Content=Trim(request.form("Content"))
company=Trim(request.form("company"))
address=Trim(request.form("address"))
user=Trim(request.form("user"))
tel=Trim(request.form("tel"))
fax=Trim(request.form("fax"))
zip=Trim(request.form("zip"))
gsurl=Trim(request.form("gsurl"))
SmallImage=trim(request.form("SmallImage"))
ImagePath=trim(request.form("ImagePath"))
IcoImage=trim(request.form("Icoimage"))
dtlx=Trim(request.form("dtlx"))
zdfy=Trim(request.form("zdfy"))
Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From network ",conn,1,3
		
rs.addnew
rs("province")=province
rs("Content")=Content
rs("company")=company
rs("address")=address
rs("SmallImage")=SmallImage
rs("IcoImage")=IcoImage
rs("ImagePath")=ImagePath
rs("user")=user
rs("zdfy")=zdfy
rs("tel")=tel
rs("fax")=fax
rs("zip")=zip
rs("gsurl")=gsurl
rs("dtlx")=dtlx
rs("lang")=lang
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=8 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=8  order by Sequence"
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



rs("adddate")=now()
rs.update
rs.close
call addlog("增加营销网络")
response.Write "<script language=javascript>alert(""恭喜您添加成功"");window.location.href=""yxwl.asp""</script>"
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
 <style type="text/css">
.ys td{ padding:10px}
.imgDiv { float:left; padding-top:5px; padding-left:5px; }
.imgDiv img{ border:0px; width:120px; height:90px}
</style>
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



 KindEditor.ready(function(K) {
 var imgh="";
 var imgurl="";
 var editor = K.editor({
 allowFileManager : true,
 uploadJson : '../zmbjq/asp/upload_json.asp'
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
                                     $('#J_imageView').append('<div class="imgDiv" id="'+ids+'"><a href="javascript:setContent(\'' + array[s] + '\')"><img title="点击图片将图片插入到编辑器" src="'+array[s] + '" width="120" height="90" border="0" /></a><br><input type="radio" name="IndexImageradio" onclick="setIndexImage(\''+array[s] +'\')" value="'+array[s] +'" />设为幻灯图 <a href="javascript:dropThisDiv(\''+ids+'\',\''+array[s] +'\')">删除</a></div>');
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
        strcon = document.getElementById('Content').value; // 原生API，content为页面上textarea的ID值
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
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>
 <form   method="post" name="myform" class="checkform" >


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">营销网络增加  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chajian.asp" target="main">返回菜单</a></div>
 
 <div class="dmeun"><a href="yxwl.asp" target="main">营销网络管理</a></div>
 
 </td>

 </tr>
</table>



                <table width="99%"  border="0" cellpadding="6" cellspacing="0"  align="center" class="dtable4">
                  
				  <tr  >
				    <td  >销售区域:</td><td>
					<select name="Province" datatype="*" nullmsg="所属省份不能为空">
    <option value="">选择省份</option>
<%
dim sqlp,rsp
sqlp="select * from Province"
set rsp=server.createobject("adodb.recordset")
rsp.open sqlp,conn,1,1
while(not rsp.eof)
%>
    <option value="<%= rsp("Province") %>" ><%= rsp("Province") %></option>
<%
rsp.movenext
wend
rsp.close
set rsp=nothing
%>
    </select>					</td>
			      </tr>
				  

			  <tr>
			    <td >公司形象图：</td><td>
				   <input name="SmallImage"  type="text" id="weepic"  value="" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
		        <div id="mysmaimg"></div>		        </td>

	      </tr>
				  <tr  > 
                    <td  > 销售点名称:</td><td>
                   
                      <input name="company"type="text" class="inputtext" id="company"   size="30" maxlength="30" datatype="*" nullmsg="名称不能为空"/>                  </td>
                  </tr><tr  > 
                    <td  > 联系地址: </td><td>
                      <input name="address" type="text" class="inputtext" id="address" size="30" maxlength="30" datatype="*" nullmsg="联系地址不能为空"/>
                    </br></br>                   </td>
                    </tr>
				  <tr  >
				    <td  >联系人: </td><td><input name="user" type="text" class="inputtext" id="user" size="20" maxlength="30" datatype="*" nullmsg="联系人不能为空"></td>
			      </tr>
				  <tr  >
				    <td  >联系电话:</td><td>
			        <input name="tel" type="text" id="tel" size="20" datatype="*" nullmsg="联系电话不能为空"/></td>
			      </tr>
				  <tr  > 
                    <td  > 传真号码:</td><td> <input name="fax"type="text" class="inputtext" id="fax"   size="20" maxlength="30"></td>
                  </tr>
				  <tr  > 
                    <td  > 邮政编码:</td><td>
                    <input name="zip"type="text" class="inputtext" id="zip"  size="20" maxlength="30" /></td>
                  </tr>
				  
				  <tr  >
				    <td  >公司网址:</td><td>
				      <label>
				      <input name="gsurl" type="text" id="gsurl" />
			        </label></td>
			      </tr>
				  
				  	  	<%'自定义字段%>
          <%call showColumn("NetWork",8,id)%>
				  
				  
				  
				  
				  
				  
				  
				  
				  
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
				  <tr  > 
                    <td  > 详细说明:</td><td> <textarea name="content" id="content" style="width:700px;height:300px;visibility:hidden;" datatype="*" nullmsg="内容不能为空"></textarea></td>
                  </tr>
				  
				  
				   <tr >
		     <td >手动分页：</td>
			 <td><INPUT type=button class="bnt" name="appendHtml" value="插入分页符"></td>
			 </tr>
			 <tr >
			<input name="zdfy" type="hidden" id="zdfy"  value="0" size="10">
		 
				  
				   <tr  >
                  <td><input type="hidden" id="IcoImage" name="IcoImage" value="" style="width:600px;" />
        批量上传：         </td>
                  <td  colspan="2"><input name="button" type="button" id="J_selectImage" style=" background:url(images/allimg.png) no-repeat; width:170px;height:34px;border:none" value="" />
                    <input type="hidden" id="ImagePath" name="ImagePath" value="" style="width:600px;" /> 
                    小提示：批量上传图片以后可点击图片插入到编辑框
                      <div id="J_imageView">  </div></td>
                </tr> 
				  
				  
				  
				  
				  
                  
    <tr >
      <td  >地图路线：</td><td>地图代码:<a href="http://api.map.baidu.com/lbsapi/creatmap/index.html"> 百度地图申请</a>
	  <br />  <br />
        <label>
        <textarea name="dtlx" cols="100" rows="30" id="map"></textarea>
        </label>
    </td>
    </tr>
                  <tr  > 
                    <td colspan="2" > 
                        <input class="bnt" name="submit" type="submit" value="确定">
&nbsp;                                      </td>
                  </tr>
            </table>
            
</form>
       

</body>
</html>
