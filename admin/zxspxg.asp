<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->

<%
If request.QueryString("id") = "" Or Not IsNumeric(request.QueryString("id")) Then
    response.Write "<script language=javascript>alert(""非法操作！"");window.history.back();</script>"
    response.End
End If
id=request.QueryString("id")
If request.Form("submit") = "修改" Then
keyword=Trim(request.form("keyword"))
Title=Trim(request.form("Title")) 
User=Trim(request.form("User"))
ly=Trim(request.form("ly"))
Content=Trim(request.form("Content"))
SmallImage=trim(request.form("SmallImage"))
ImagePath=trim(request.form("ImagePath"))
IcoImage=trim(request.form("Icoimage"))
zdfy=trim(request.form("zdfy"))
Hits=trim(request.form("Hits"))
zmone=trim(request.form("zmone"))
zmtwo=trim(request.form("zmtwo"))
zmthree=trim(request.form("zmthree"))
btcolor=trim(request.form("btcolor"))
screencolor=trim(request.form("screencolor"))
stretching=trim(request.form("stretching"))
controlbar=trim(request.form("controlbar"))
spwidth=trim(request.form("spwidth"))
spheight=trim(request.form("spheight"))
skin=trim(request.form("skin"))
spurl=trim(request.form("spurl"))
wapspurl=trim(request.form("wapspurl"))
arrspurl=trim(request.form("arrspurl"))
tvkey=trim(request.form("tvkey"))
tvms=trim(request.form("tvms"))
tj=Trim(request.form("tj"))
adddate=Trim(request.form("adddate"))

ClassId=trim(Request.Form("ClassId"))
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From spClass where ID="&ClassId  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then
	  ParentClassID=rs("ParentClassID")  
	  end if
	  rs.close
	  if ParentClassID<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From spclass where id="&ParentClassID  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID1=rs("ParentClassID")  
	  end if
	  rs.close
	  end if
      if ParentClassID1<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From spclass where id="&ParentClassID1 
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID2=rs("ParentClassID")     
	  end if
	  rs.close
	  end if

Set rs = server.CreateObject("adodb.recordset")
sql="Select * From zxsp Where id = "&request.QueryString("id")&""
rs.open sql,conn,1,3
GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
rs("GroupID")=GroupIdName(0)
rs("Exclusive")=trim(Request.Form("Exclusive"))
rs("Title")=Title
rs("Content")=Content
rs("pic")=pic
rs("Hits")=Hits
rs("spurl")=spurl
rs("wapspurl")=wapspurl
rs("ly")=ly

'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=5 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=5  order by Sequence"
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

rs("lang")=lang
rs("screencolor")=screencolor
rs("stretching")=stretching
rs("controlbar")=controlbar
rs("skin")=skin
rs("User")=User
rs("arrspurl")=arrspurl
rs("IcoImage")=IcoImage
rs("ImagePath")=ImagePath
rs("spwidth")=spwidth
rs("spheight")=spheight
rs("adddate")=date()
	 if ParentClassID=0 then
rs("zmone")=trim(Request.Form("ClassId"))
rs("zmtwo")=0
rs("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1=0 then
rs("zmone")=ParentClassID
rs("zmtwo")=trim(Request.Form("ClassId"))
rs("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1<>0 and ParentClassID2=0 then
rs("zmone")=ParentClassID1
rs("zmtwo")=ParentClassID
rs("zmthree")=trim(Request.Form("ClassId"))
end if
rs("tvkey")=tvkey
rs("btcolor")=btcolor
rs("tvms")=tvms
rs("zdfy")=zdfy
rs("SmallImage") = SmallImage
rs("keyword")=keyword
rs("adddate")=adddate
 if tj="1" then
		rs("tj")=1
	else
		rs("tj")=0
	end if
rs.update
rs.close
set rs=nothing


'获取添加新闻的ID  生成html开始
set rs=conn.execute("select id,zmone,zmtwo,zmthree from [zxsp] where id="&id&"")
tvid=rs(0)
pone=rs(1)
ptwo=rs(2)
pthree=rs(3)
rs.close:set rs=Nothing


if yy=2 then

sql="Select * from zxsp where id="&tvid&" order by ID desc"
Set Rs=Conn.Execute(sql)
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")
ptwo=rs("zmtwo")
pthree=rs("zmthree")


'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取视频分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where id = "&pone&" "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If




'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取视频分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where id = "&ptwo&" "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If

'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取视频分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where id = "&pthree&" "
rs1.open sql1,conn,1,1
m=rs1("m")

if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&tvname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showvideo&"/index.asp","pone=",pone,"videoID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If






























If pone<>0 And ptwo=0 And pthree=0 Then

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where ID ="&pone&""
rs1.open sql1,conn,1,1
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from zxsp where zmone="&pone&"")(0)
totalpage=int(totalrec/tvinfo)       
If (totalpage * tvinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&tvclass&""&Separated&""&pone&""&Separated&"1.html",""&video&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&tvclass&""&Separated&""&pone&""&Separated&""&i&".html",""&video&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",i)
next
End If

rs1.close:set rs1=Nothing
end if





If pone<>0 And ptwo<>0 And pthree=0 Then
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from spclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
m=rs2("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from zxsp where zmone="&pone&" and zmtwo="&ptwo&"")(0)
totalpage=int(totalrec/tvinfo)       
If (totalpage * tvinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&tvclass&""&Separated&""&ptwo&""&Separated&"1.html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&tvclass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",i,"","")
next
End If







rs2.close:set rs2=Nothing
rs1.close:set rs1=Nothing
end if


If pone<>0 And ptwo<>0 And pthree<>0 Then
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from spclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from spclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from spclass where ID ="&pthree&""
rs3.open sql3,conn,1,1
m=rs3("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from zxsp where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&"")(0)
totalpage=int(totalrec/tvinfo)       
If (totalpage * tvinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&tvclass&""&Separated&""&pthree&""&Separated&"1.html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&tvclass&""&Separated&""&pthree&""&Separated&""&i&".html",""&video&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"language=",lang,"m=",m,"Page=",i)
next
End If
rs3.close:set rs3=Nothing
rs2.close:set rs2=Nothing
rs1.close:set rs1=Nothing
end if



rs.close
set rs=Nothing





call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&video&"/","",""&tvclass&".html",""&video&"/index.asp","m=",mtv,"language=",lang,"","","","","","","","")

end if





call addlog("修改"&shipin&"")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""sp_gl.asp""</script>"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="javascript" type="text/javascript" src="js/getdate/WdatePicker.js"></script>
<style type="text/css">
 .imgDiv { float:left; padding-top:5px; padding-left:5px; }
 .imgDiv img{ border:0px; width:120px; height:90px}
</style>

		
	<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
		<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<!--#include file="sc.asp"-->
<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
</script>
<script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
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
                                     $('#J_imageView').append('<div class="imgDiv" id="'+ids+'"><a href="javascript:setContent(\'' + array[s] + '\')"><img title="点击图片将图片插入到编辑器" src="'+array[s] + '" width="120" height="90" border="0" /></a><br><label><input type="radio" name="IndexImageradio" onclick="setIndexImage(\''+array[s] +'\')" value="'+array[s] +'" />设为幻灯图</label> <a href="javascript:dropThisDiv(\''+ids+'\',\''+array[s] +'\')">删除</a></div>');
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

<!--#include file="ys.asp"-->
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->

</head>

<body>
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from zxsp where id = "&request.QueryString("id")&" and lang='"&lang&"'"
rs.Open sql, conn, 1, 1
If rs.EOF Then
    response.Write "<script language=javascript>window.location.href=""sp_gl.asp""</script>"
    response.End
End If
 IcoImage=rs("IcoImage")
 ImagePath=rs("ImagePath")
Exclusive=rs("Exclusive")
groupid=rs("groupid")
if rs("zmone")<>0 and  rs("zmtwo")=0  and rs("zmthree")=0 then
ClassId=rs("zmone")
end if
if rs("zmone")<>0 and  rs("zmtwo")<>0 and rs("zmthree")=0  then
ClassId=rs("zmtwo")
end if
if rs("zmone")<>0 and  rs("zmtwo")<>0 and rs("zmthree")<>0  then
ClassId=rs("zmthree")
end if
%>


     <form method="post" name="myform" class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改<%=shipin%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="shipin.asp" target="main">返回菜单</a></div>
 
  <div class="dmeun"><a href="sp_gl.asp" target="main"><%=shipin%>管理</a></div>
 
 </td>
 

 </tr>
</table>
          
           <table width="99%" border="0" cellpadding="0" cellspacing="0" align="center" class="dtable4">
             
             
<tr>
      <td width="15%" ><%=shipin%>标题：</td>
	  <td width="85%">
                 <input name="Title" type="text" id="Title" value="<%=rs("title")%>" size="40" datatype="*" nullmsg="名称不能为空"/>  
      </td>
</tr>

<tr>
      <td >标题颜色：</td>
	  <td><input name="btcolor"type="text" id="color"   value="<%=rs("btcolor")%>" size="8" />
                 <input name="button"   type="button" id="colorpicker" value="打开取色器" /></td>
</tr>
           <tr>
             <td  >	tags:</td>
  <td>
       <input name="KeyWord" type="text" id="KeyWord" value="<%=rs("keyword")%>" size="40"  datatype="*" nullmsg="Tags不能为空">
     【<a href="#" id="KeyLinkByTitle" style="color:green">根据标题自动获取</a>】(只对中文内容生效，多个tags用|隔开)
     </td>

</tr>
          
         <tr > 
<td>产品类别：</td>
  <td><select ID="ClassID" name="ClassID" datatype="*" nullmsg="所属分类不能为空">
	   <%call mxlb("spClass",ClassId)%>
	  </select><font color=red>（有三级栏目的必须选择三级栏目）</font></td>
</tr>
<tr>
  <td  ><%=shipin%>关键词：</td>
  <td>
    <label>
    <input name="tvkey" type="text" id="tvkey" value="<%=rs("tvkey")%>" size="100" />
    </label></td>
</tr>
<tr>
  <td  ><%=shipin%>描述：</td>
  <td>
    <label>
    <input name="tvms" type="text" id="tvms" value="<%=rs("tvms")%>" size="100" />
    </label></td>
</tr>

				  	<%'自定义字段%>
<%call showColumn("zxsp",5,id)%>

		 	<tr>
        <td >查看权限：</td>
  <td>
          <select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> 
          隶属
           <input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
          专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
      </tr> 

			  <tr>
			    <td ><%=shipin%>缩略图：</td>
  <td>
				   <input name="SmallImage"  type="text" id="weepic"  value="<%=rs("SmallImage")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
<div id="mysmaimg"><%if rs("Smallimage")<>"" then%>
<img src="../<%=rs("SmallImage")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs("Smallimage")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>			        </td>

	      </tr>
             
<tr>
  <td  ><%=shipin%>播放地址：</td>
  <td><input name="spurl" type="text" id="url" value="<%=rs("spurl")%>" style="width:400px" /> 
            <input type="button" id="insertfile" value="选择文件" /> <br />建议大家调用各大网站视频，以节省服务器资源。如果要调用本地视频，建议调用flv,mp4格式
		  (支持swf,flv,mp4) 
		  支持调用各大视频网站的swf 地址（例子：http://www.aipai.com/c13/NDY5ISYnJGgnaiQg/playerOut.swf)</td>
</tr>

	  <tr>
				    <td  ></td>
  <td>wap<%=shipin%>播放地址：<input name="wapspurl" type="text" id="url1" value="<%=rs("wapspurl")%>" style="width:400px" /> 
            <input type="button" id="insertfile1" value="选择文件" /> <font color="#FF0000">*</font>(手机只支持MP4格式的视频)</td>
			      </tr>
<tr>
				      <td >多文件连播地址：</td>
  <td>
				        <label>
				        <textarea name="arrspurl" rows="5" id="arrspurl" style="width:800px;height:300px"><%=rs("arrspurl")%></textarea>
			          </label></br>多文件连播地址：只支持flv,mp4,多个文件请用|隔开.可以是flv或者mp4， 如果调用多文件连播，请把 单视频播放地址清空，同样如果只调用单文件播放，请把多文件连播清空。  只支持调用flv,mp4的格式。</br>					  </td>
			      </tr>	  
				  
		 <tr>
			          <td colspan="2" >播放器设置：只有在播放文件为flv，mp4时有效</td>
		          </tr>
			        <tr>
			          <td >播放屏幕的背景颜色:</td>
  <td>
			            <label>
			            <input type="text" name="screencolor" value="<%=rs("screencolor")%>"/>   颜色不带#号 加入黑色#000000 只需输入#号后面的6位就可以了
		              </label></td>
		          </tr>
			        <tr>
			          <td >控制面板放置位置:</td>
  <td>
			            <label>
			            <select name="controlbar" id="controlbar">
		                  <option value="none" <%if rs("controlbar")="none" then Response.Write("selected")%>>无</option>
		                  <option value="bottom" <%if rs("controlbar")="bottom" then Response.Write("selected")%>>底部</option>
		                  <option value="over" <%if rs("controlbar")="over" then Response.Write("selected")%>>覆盖</option>
	                    </select>
		              </label></td>
		          </tr>
			        <tr>
			          <td>播放器封面图片延展方式:</td>
  <td>
			            <label>
			            <select name="stretching" id="stretching">
		                  <option value="none" <%if rs("stretching")="none" then Response.Write("selected")%>>不延展</option>
		                  <option value="exactfit" <%if rs("stretching")="exactfit" then Response.Write("selected")%>>不锁定宽高填充</option>
		                  <option value="fill" <%if rs("stretching")="fill" then Response.Write("selected")%>>锁定宽高填充</option>
		                  <option value="uniform" <%if rs("stretching")="uniform" then Response.Write("selected")%>>锁定以黑色填充空白</option>
	                    </select>
		              </label></td>
		          </tr>
		            <tr>
		              <td>播放器皮肤:</td>
  <td>
		                <label>
		                <input type="text" name="skin"  value="<%=rs("skin")%>"/>
	                  填写 数字 1-93 </label></td>
	              </tr>		  
				  
				    <tr>
		              <td  >播放器宽度:</td>
  <td>
		                <label>
		                <input type="text" name="spwidth" value="<%=rs("spwidth")%>"/>
	                  </label></td>
	              </tr>
				   <tr>
		              <td  >播放器高度:</td>
  <td>
		                <label>
		                <input type="text" name="spheight" value="<%=rs("spheight")%>"/>
	                  </label></td>
	              </tr>
				  
				  
				  
				  
<script>
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
 
 
 
 
 
 
 
<tr>
               <td  ><%=shipin%>说明:</td><td><textarea datatype="*" nullmsg="内容不能为空" name="Content" id="Content" style="width:700px;height:300px;visibility:hidden;"><%=rs("content")%></textarea></td>
             </tr>
	
	 <tr>
      <td >手动分页：</td>
			   <td><INPUT type=button class="bnt" name="appendHtml" value="插入分页符">
			   </td>
          </tr>
          <tr>
      
        <input name="zdfy" type="hidden" id="zdfy"  value="<%=rs("zdfy")%>" size="10" />
        
	 <tr  >
            <td>批量上传图片</td><td>
			幻灯图片地址：<input type="text" id="IcoImage" name="IcoImage" value="<%=rs("icoimage")%>" style="width:600px;" /><br/><br/>
		<input type="button" id="J_selectImage" value="" style=" background:url(images/allimg.png) no-repeat; width:170px;height:34px;border:none" /><br/><br/>
         多图片地址：<input type="text" id="ImagePath" name="ImagePath" value="<%=rs("ImagePath")%>" style="width:600px;" /><br/><br/>
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
<a href="javascript:setContent('<%=img(i)%>')" ><img title="点击图片将图片插入到编辑器" src="<%=img(i)%>" border="0" /></a><br><label><input type="radio" name="IndexImageradio" onClick="setIndexImage('<%=img(i)%>')" value="<%=img(i)%>" <%if rs("icoimage")=img(i) then response.write "checked" %>/>设为幻灯图</label> <a href="javascript:dropThisDiv('<%=ids%>','<%=img(i)%>')">删除</a></div>
 <%
 next
 end if
 %>    
 </div></td>
          </tr>
	
	
	
	
	 <tr  >
	   <td  >是否推荐:</td><td><input name="tj" type="checkbox" id="tj" value="1" <% if rs("tj")=1 then response.Write("checked") end if%>></td>
	   </tr>
	   
	   		  <tr > 
            <td  >点击次数：</td>
			<td>
              <input  name="hits" type="text" class="input" value="<%=rs("hits")%>"  size="30"> 
			  </td>
          </tr>
	   
	   
	   
	 <tr  >
      <td  >发布人：</td><td>
        <input  name="user" type="text" class="input" size="30" value="<%=rs("user")%>">      </td>
    </tr>
    <tr  >
      <td  >来源：</td><td>
      <input  name="ly" type="text" id="ly" value="<%=rs("ly")%>"></td>
    </tr>
			 
			 

             
  <tr>
    <td >发布时间：</td><td>
      <label>
      <input name="adddate" type="text" id="adddate" value="<%=rs("adddate")%>"  onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})"/>
      </label></td>
  </tr>
  <tr>
      <td colspan="2">
                   
        <div align="left">
          <input class="bnt" name="submit" type="submit" value="修改" />
          </div></td>
             </tr>
       </table>
  
     </form>
 
 <%
rs.close
set rs=nothing
%>
</body>
</html>
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