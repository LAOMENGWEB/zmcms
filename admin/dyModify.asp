<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<% 
id=request.QueryString("id")
If request.Form("submit") = "修改" Then

Title=Trim(request.form("Title")) 
Content=Trim(request.form("Content"))
px=Trim(request.form("px"))
gjz=Trim(request.form("gjz"))
ms=Trim(request.form("ms"))
zdfy=Trim(request.form("zdfy"))
zmone=Trim(request.form("zmone"))
wlurl=Trim(request.form("wlurl"))
SmallImage=trim(request.form("SmallImage"))
ImagePath=trim(request.form("ImagePath"))
IcoImage=trim(request.form("Icoimage"))

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from dy where id="&id
rs.open sql,conn,1,3
GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
rs("GroupID")=GroupIdName(0)
rs("Exclusive")=trim(Request.Form("Exclusive"))
rs("Title")=Title
rs("Content")=Content
rs("px")=px
rs("gjz")=gjz
rs("zdfy")=zdfy
rs("lang")=lang
'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=1 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=1  order by Sequence"
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
rs("zmone")=zmone
rs("wlurl")=wlurl
rs("ms")=ms
rs("SmallImage")=SmallImage
rs("IcoImage")=IcoImage
rs("ImagePath")=ImagePath
rs("updatetime")=now()
rs.update
rs.close
set rs=nothing

if yy=2 then
sql="Select * from dy where id="&id&" order by ID desc"
Set Rs=Conn.Execute(sql)
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from mnclass where id = "&pone&" "
rs1.open sql1,conn,1,1

m=rs1("m")


rs1.close
set rs1=Nothing


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&aboutname&""&Separated&""&pone&""&Separated&""&id&""&Separated&"1.html",""&aabout&"/index.asp","pone=",pone,"aboutID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&aboutname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&aabout&"/index.asp","pone=",pone,"aboutID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end if




rs.close
set rs=Nothing
call Htmlabout
call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
end if
call addlog("修改"&danye&"信息")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""dy.asp""</script>"
end if
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From dy where id="&id&" and lang='"&lang&"'", conn,1,1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""dy.asp""</script>"
    response.End
End if
Exclusive=rs("Exclusive")
groupid=rs("groupid")
 Content=rs("Content")
 IcoImage=rs("IcoImage")
 ImagePath=rs("ImagePath")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.ys td{ padding:10px}
 .imgDiv { float:left; padding-top:5px; padding-left:5px; }
 .imgDiv img{ border:0px; width:120px; height:90px}
</style>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">

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
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
<script language=javascript>
		function ResumeError()
			{return true;}
			window.onerror = ResumeError;
			$(document).ready(function(){
				
			  $('#KeyLinkByTitle').click(function(){
			    GetKeyTags();
			  });
			});
			function GetKeyTags()
			{
			  var text=escape($('input[name=Title]').val());
			  if (text!=''){
				  $('#KeyWord').val('稍等,正在根据标题自动获取tags');
				  $("#KeyWord").attr("disabled",true);
				  $.get("ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWord').val(unescape(data));
					$('#KeyWord').attr("disabled",false);
				  });
			  }else{
			   alert('请先输入标题!');
			  }
			}
</script>

<script>
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
</head>

<body>
<form method="post" name="myform"  class="checkform">




<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改<%=danye%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>




<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="danye.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="dy.asp" target="main"><%=danye%>列表管理</a></div><div class="dmeun"><a href="fl_dy.asp" target="main"><%=danye%>分类管理</a></div></td>
 

 </tr>
</table>



                  <table width="99%"  border="0" cellpadding="0" cellspacing="0"align="center" class="dtable4"> 
				    
				<tr  >
                      <td width="17%"  ><%=danye%>名称:                                                                   </td>
                  <td width="18%"  ><input name="Title" type="text" id="Title" value="<%=rs("Title")%>" size="25"  datatype="*" nullmsg="<%=danye%>名称不能为空"></td>
                  <td width="65%"  ><div class="Validform_checktip"><%=danye%>名称不能为空</div></td>
				</tr>
					     
                      <tr  >
                        <td  >所属分类:</td>
                        <td  ><select name="zmone" id="select" datatype="*" nullmsg="所属分类不能为空">
      <%
		  dim rsc
		  set rsc=server.CreateObject("adodb.recordset")
		  rsc.open "select * from mnclass where lang='"&lang&"'",conn,1,1
		  Response.Write("<option value="""">请选择分类</option>") 
    If Rsc.Eof and Rsc.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
    Else
		  while not rsc.eof
		    if rs("zmone")=rsc("id") then
			response.Write("<option value=""" & rsc("id") & """ selected>" & rsc("classname") & "</option>")
			else
			response.Write("<option value=""" & rsc("id") & """>" & rsc("classname") & "</option>")
			end if
			rsc.movenext
		wend
		end if
		rsc.close
		set rsc=nothing
		  %>
                        </select></td>
                        <td  ><div class="Validform_checktip">所属分类不能为空</div></td>
                      </tr>
                      <tr  >
                        <td  >外链地址:                        </td>
                        <td   colspan="2"><input name="wlurl" type="text" id="wlurl" value="<%=rs("wlurl")%>" /></td>
                      </tr>
                      <tr  >
                      <td  >关键字:                                            </td>
                      <td colspan="2"  ><input name="gjz" type="text" id="gjz" value="<%=rs("gjz")%>" size="80" ></td>
                      </tr><tr  >
                      <td  >描述:                        </td>
                      <td colspan="2"  ><input name="ms" type="text" id="ms" value="<%=rs("ms")%>" size="80" ></td>
                      </tr>
					  <%call showColumn("dy",1,id)%>
					 	<tr>
        <td>查看权限：         </td>
        <td ><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select></td>
        <td > <input class="none" name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> 
          隶属
           <input class="none" type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
          专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
				 	</tr>  
				<tr >

			    <td >缩略图：
				  	      			        </td>
<td colspan="2">  <input name="SmallImage"  type="text" id="weepic"  value="<%=rs("SmallImage")%>" size="50"  />	<br/><br/>	<a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a><br />
<div id="mysmaimg"><%if rs("Smallimage")<>"" then%>
<img src="../<%=rs("SmallImage")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("Smallimage")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div></td>
	      </tr><script>
 var editor;
 KindEditor.ready(function(K) {
 editor = K.create('textarea[name="Content"]', {
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
      <td  ><%=danye%>内容：</td>
      <td colspan="2"  ><textarea datatype="*" nullmsg="<%=danye%>内容不能为空"  name="Content" id="Content" style="width:700px;height:300px;visibility:hidden;"><%=rs("content")%></textarea></td>
      </tr>
	
	 <tr  >
	 <td>批量上传图片：</td>
            <td  colspan="2" >
			幻灯图片地址：<input type="text" id="IcoImage" name="IcoImage" value="<%=rs("icoimage")%>" style="width:600px;" /><br/><br/>
			<input name="button" type="button" id="J_selectImage" style=" background:url(images/allimg.png) no-repeat; width:170px;height:34px;border:none" value="" />  
			<br/></br>
        多图地址： <input type="text" id="ImagePath" name="ImagePath" value="<%=rs("ImagePath")%>" style="width:600px;" /><br/></br>
			
			<div id="J_imageView">
 <%
 '输出已有图片信息
'die imgUrls
 Dim img,i,ids
 if ImagePath<>"" then
 img = split(ImagePath,"|")
 For i=0 to ubound(img) 
 ids=replace(replace(img(i),"/",""),".","")
 
 %>
<div class="imgDiv" id="<%=ids%>">
<a href="javascript:setContent('<%=img(i)%>')" ><img title="点击图片将图片插入到编辑器" src="<%=img(i)%>" border="0" /></a><br><input type="radio" name="IndexImageradio" onClick="setIndexImage('<%=img(i)%>')" value="<%=img(i)%>" <%if rs("icoimage")=img(i) then response.write "checked" %>/>设为幻灯图 <a href="javascript:dropThisDiv('<%=ids%>','<%=img(i)%>')"">删除</a></div>
 <%
 next
 end if
 %>    
 </div></td>
            </tr>
	
	
	
	
	
                      <tr  >
		     <td  >手动分页：		      </td>
	         <td colspan="2"><INPUT type=button class="bnt" name="appendHtml" value="插入分页符"></td>
                    </tr>
                 <input name="zdfy" type="hidden" id="zdfy"  value="<%=rs("zdfy")%>" size="10">
                     
                     <tr  >
                      <td  >排序:</td>
                      <td  colspan="2" ><input name="px" type="text" id="px" value="<%=rs("px")%>" size="4"  /></td>
                     </tr>
                      <tr  >
                        <td colspan="3"  style="border-bottom:none">
                          <div align="left">
                            <input class="bnt"name="submit" type="submit" value="修改">
                          &nbsp;</div></td></tr>
              </table>
             
</form>
 

</body>
</html>
<%
rs.close
set rs=nothing
%>

<% 
sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from hyGroup where lang='"&lang&"'"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
    if GroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
end sub
%>