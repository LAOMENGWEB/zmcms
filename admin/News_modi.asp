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
zmone=trim(request.form("zmone"))
zmtwo=trim(request.form("zmtwo"))
zmthree=trim(request.form("zmthree"))
btcolor=trim(request.form("btcolor"))
Content=trim(request.form("Content"))
keyword=trim(request.form("keyword"))
User=Trim(request.form("User"))
bfnr=Trim(request.form("bfnr"))
newkey=Trim(request.form("newkey"))
newms=Trim(request.form("newms"))
ly=Trim(request.form("ly"))
adddate=trim(request("adddate"))
Hits=trim(request.form("Hits"))
tj=trim(request("tj"))
tt=trim(request("tt"))
ft=trim(request("ft"))
gg=trim(request("gg"))
sytj=trim(request("sytj"))
zdfy=trim(request("zdfy"))
wlurl=trim(request("wlurl"))
SmallImage=trim(request.form("SmallImage"))
ImagePath=trim(request.form("ImagePath"))
IcoImage=trim(request.form("Icoimage"))
ClassId=trim(Request.Form("ClassId"))
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From newsClass where ID="&ClassId  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then
	  ParentClassID=rs("ParentClassID")  
	  end if
	  rs.close
	  if ParentClassID<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From newsclass where id="&ParentClassID  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID1=rs("ParentClassID")  
	  end if
	  rs.close
	  end if
      if ParentClassID1<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From newsclass where id="&ParentClassID1 
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID2=rs("ParentClassID")     
	  end if
	  rs.close
	  end if
Set rs_news = Server.CreateObject("ADODB.Recordset")
sql1="select * from news where id="&id
rs_news.open sql1,conn,1,3
GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
rs_news("GroupID")=GroupIdName(0)
rs_news("Exclusive")=trim(Request.Form("Exclusive"))
rs_news("Title")=Title
rs_news("Content")=Content
rs_news("bfnr")=bfnr
rs_news("User")=User
rs_news("btcolor")=btcolor
if ParentClassID=0 then
rs_news("zmone")=trim(Request.Form("ClassId"))
rs_news("zmtwo")=0
rs_news("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1=0 then
rs_news("zmone")=ParentClassID
rs_news("zmtwo")=trim(Request.Form("ClassId"))
rs_news("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1<>0 and ParentClassID2=0 then
rs_news("zmone")=ParentClassID1
rs_news("zmtwo")=ParentClassID
rs_news("zmthree")=trim(Request.Form("ClassId"))
end if
  
  rs_news("adddate")=adddate
  rs_news("ly")=ly
  rs_news("Hits")=Hits
  rs_news("zdfy")=zdfy
  rs_news("wlurl")=wlurl
  rs_news("newkey")=newkey
  rs_news("newms")=newms
  rs_news("lang")=lang
  
  '自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=2 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=2  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd2=request.form(""&zdyzd1&"")
rs_news(""&zdyzd1&"")=zdyzd2
rsObj.movenext
loop
end if
'自定义字段结束
  
  
  
  
  rs_news("SmallImage")=SmallImage
  rs_news("IcoImage")=IcoImage
  rs_news("ImagePath")=ImagePath
  rs_news("keyword")=keyword
 if tj="1" then
		rs_news("tj")=1
	else
		rs_news("tj")=0
	end if
	 if ft="1" then
		rs_news("ft")=1
	else
		rs_news("ft")=0
	end if

    if tt="1" then
		rs_news("tt")=1
	else
		rs_news("tt")=0
	end if
	if gg="1" then
		rs_news("gg")=1
	else
		rs_news("gg")=0
	end if
	 if sytj="1" then
		rs_news("sytj")=1
	else
		rs_news("sytj")=0
	end if
	rs_news.update
rs_news.close
set rs_news=nothing


'获取添加新闻的ID  生成html开始
set rs=conn.execute("select id,zmone,zmtwo,zmthree from [news] where id="&id&"")
newid=rs(0)
pone=rs(1)
ptwo=rs(2)
pthree=rs(3)
rs.close:set rs=Nothing

if yy=2 then
'生成添加新闻的html
sql="Select * from News where id="&newid&" order by ID desc"
Set Rs=Conn.Execute(sql)
content=rs("content")
zdfy=rs("zdfy")

pone=rs("zmone") '一级分类
ptwo=rs("zmtwo") '二级分类
pthree=rs("zmthree") '三级分类


'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where id = "&pone&" "
rs1.open sql1,conn,1,1
m=rs1("m")

	
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&id&""&Separated&"1.html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If


rs1.close
Set rs1=nothing
End If
'一级分类结束




'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where id = "&ptwo&" "
rs1.open sql1,conn,1,1
m=rs1("m")

	
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&id&""&Separated&"1.html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If


rs1.close
Set rs1=nothing
End If
'二级分类结束


'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where id = "&pthree&" "
rs1.open sql1,conn,1,1
m=rs1("m")

	
if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&id&""&Separated&"1.html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&NewName&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showinfo&"/index.asp","pone=",pone,"newsID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If


rs1.close
Set rs1=nothing
End If
'三级分类结束




'生成所属所属分类的新闻

If pone<>0 And ptwo=0 And pthree=0 Then

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where ID ="&pone&""
rs1.open sql1,conn,1,1
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&"")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
creatwjj("../"&rs1("mulu")&"/")
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&NewClass&""&Separated&""&pone&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&NewClass&""&Separated&""&pone&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",i)
next
End If


rs1.close:set rs1=Nothing

end if



If pone<>0 And  ptwo<>0 And pthree=0  then

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from newsclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
m=rs2("m")

creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and zmtwo="&ptwo&"")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&NewClass&""&Separated&""&ptwo&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&NewClass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",i,"","")
next
End If




rs2.close:set rs2=Nothing

rs1.close:set rs1=Nothing
end if


If pone<>0 And ptwo<>0 and pthree<>0 then

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from newsclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from newsclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from newsclass where ID ="&pthree&""
rs3.open sql3,conn,1,1
m=rs3("m")

creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from news where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&"")(0)
totalpage=int(totalrec/NewInfo)       
If (totalpage * NewInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&NewClass&""&Separated&""&pthree&""&Separated&"1.html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&NewClass&""&Separated&""&pthree&""&Separated&""&i&".html",""&iinfo&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m,"language=",lang,"Page=",i)
next
End If
rs3.close:set rs3=Nothing
rs2.close:set rs2=Nothing
rs1.close:set rs1=Nothing
end If


rs.close
set rs=Nothing





call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&iinfo&"/","",""&Newclass&".html",""&iinfo&"/index.asp","m=",minfo,"language=",lang,"","","","","","","","")
end if






call addlog("修改"&xinwen&"")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""news_manage.asp""</script>"
end if
Set rs_news = server.CreateObject("adodb.recordset")
sql2 = "select * from news where id = "&request.QueryString("id")&" and lang='"&lang&"'"
rs_news.Open sql2, conn, 1, 1
If rs_news.EOF Then
    response.Write "<script language=javascript>window.location.href=""news_manage.asp""</script>"
    response.End
End If
Exclusive=rs_news("Exclusive")
groupid=rs_news("groupid")
if rs_news("zmone")<>0 and  rs_news("zmtwo")=0  and rs_news("zmthree")=0 then
ClassId=rs_news("zmone")
end if
if rs_news("zmone")<>0 and  rs_news("zmtwo")<>0 and rs_news("zmthree")=0  then
ClassId=rs_news("zmtwo")
end if
if rs_news("zmone")<>0 and  rs_news("zmtwo")<>0 and rs_news("zmthree")<>0  then
ClassId=rs_news("zmthree")
end if
 Content=rs_news("Content")
 IcoImage=rs_news("IcoImage")
 ImagePath=rs_news("ImagePath")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<style type="text/css">
.ys td{ padding:10px}
 .imgDiv { float:left; padding-top:5px; padding-left:5px; }
 .imgDiv img{ border:0px; width:120px; height:90px}
</style>
<script language="javascript" type="text/javascript" src="js/getdate/WdatePicker.js"></script>
<script language="javascript" src="js/Admin.js"></script>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
<!--#include file="ys.asp"-->
<script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>
<link href="../javascript/colorbox/colorbox.css" rel="stylesheet" media="screen">

<script language=javascript>
$(document).ready(function(){
$(".allfile").colorbox({iframe: true, innerWidth: 1000,innerHeight: 500});
$(".sfile").colorbox({iframe: true, innerWidth: 400,innerHeight: 160});
});
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

<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->
</head>

<body>

<form method="POST" name="myform"   class="checkform">

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改<%=xinwen%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="xinwen.asp" target="main">返回菜单</a></div><div class="dmeun"><a href="news_manage.asp" target="main"><%=xinwen%>管理</a></div></td>
 

 </tr>
</table>
<table  width="99%"  border="0" align="center" cellpadding="0" cellspacing="0"  class="dtable4"   >
  
    
    <tr  >
      <td width="12%"  ><%=xinwen%>标题：</td>
	  <td width="88%">
        <input  name="Title" type="text" size="50" id="title" value="<%=rs_news("title")%>" datatype="*" nullmsg="标题不能为空"/>
      </td>
    </tr>
	 <tr  >
      <td  >标题颜色：</td>
	  <td><input name="btcolor"type="text" id="color" value="<%=rs_news("btcolor")%>" size="8" /> 
      <input   type="button" id="colorpicker" value="打开取色器" /></td>
	  </tr>
    <tr  >
      <td  >Tags标签:</td>
	  <td>
	  
	  
   
     <input name="KeyWord" type="text" id="KeyWord" value="<%=rs_news("KeyWord")%>" size="40"  datatype="*" nullmsg="Tags标签不能为空">
     【<a href="#" id="KeyLinkByTitle" style="color:green">根据标题自动获取</a>】(只对中文内容生效,多个tags之间用|隔开)
     </td>
    </tr>
    <tr  >
      <td  ><%=xinwen%>类别：</td>
	  <td>
      <select ID="ClassID" name="ClassID" datatype="*" nullmsg="所属类别不能为空">
	   <%call mxlb("NewsClass",ClassId)%> <font color=red>（有三级栏目的必须选择三级栏目）</font>
	  </select></td>
    </tr>
	 <tr  >
      <td  ><%=xinwen%>关键词: </td>
	  <td>
	           <input name="newkey" type="text" id="newkey" value="<%=rs_news("newkey")%>" size="100" />       </td>
    </tr>
	 <tr  >
      <td  ><%=xinwen%>描述:</td>
	  <td>
	           <input name="newms" type="text" id="newms" value="<%=rs_news("newms")%>" size="100" />       </td>
    </tr>
    <tr  >
      <td  >外链<%=xinwen%>:</td>
	  <td>
	           <input name="wlurl" type="text" id="wlurl" value="<%=rs_news("wlurl")%>" />
	           （如果不是外链<%=xinwen%>请留空）</td>
    </tr>
	
	<%'自定义字段%>
<%call showColumn("news",2,id)%>
	
	
	
	
	
   	<tr>
        <td>查看权限：</td>
	  <td>
          <select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> 
          隶属
           <input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
          专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
      </tr>
   <script>
 var editor1;
 KindEditor.ready(function(K) {
 editor1 = K.create('textarea[name="bfnr"]', {
 filterMode : false, 
 uploadJson : '../zmbjq/asp/upload_json.asp',
 fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
 allowFileManager : true,
 width: '700px',
 hight: '300px',
 afterBlur: function(){this.sync();}
 });

 });
 </script>
 
		   <tr > 
            <td   >内容提示：</td>
	  <td>
            <textarea name="bfnr" id="bfnr" style="width:700px;height:300px;visibility:hidden;"><%=rs_news("bfnr")%></textarea></td>
          </tr>
   
			  <tr>
			    <td ><%=xinwen%>缩略图：</td>
	  <td>
				   <input name="SmallImage"  type="text" id="weepic"  value="<%=rs_news("SmallImage")%>" size="50"  /><br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
<div id="mysmaimg"><%if rs_news("Smallimage")<>"" then%>
<img src="../<%=rs_news("SmallImage")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs_news("Smallimage")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>			        </td>

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
      <td  ><%=xinwen%>内容：</td>
	  <td>
     <textarea name="Content" id="Content" style="width:700px;height:300px;visibility:hidden;" datatype="*" nullmsg="内容不能为空"><%=rs_news("content")%></textarea><font color="#FF0000">*</font></td>
    </tr>
	
	   <tr  >
            <td  >批量图片上传：</td>
	  <td>
		幻灯图片地址:	<input type="text" id="IcoImage" name="IcoImage" value="<%=rs_news("icoimage")%>" style="width:600px;" /><br/><br/>
		<input type="button" id="J_selectImage" value="" style=" background:url(images/allimg.png) no-repeat; width:170px;height:34px;border:none" /><br/><br/>
         多图地址：<input type="text" id="ImagePath" name="ImagePath" value="<%=rs_news("ImagePath")%>" style="width:600px;" />
			
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
<a href="javascript:setContent('<%=img(i)%>')" ><img title="点击图片将图片插入到编辑器" src="<%=img(i)%>" border="0" /></a><br><label><input type="radio" name="IndexImageradio" onClick="setIndexImage('<%=img(i)%>')" value="<%=img(i)%>" <%if rs_news("icoimage")=img(i) then response.write "checked" %>/>设为幻灯图</label> <a href="javascript:dropThisDiv('<%=ids%>','<%=img(i)%>')"">删除</a></div>
 <%
 next
 end if
 %>    
 </div></td>
          </tr>
	
	
	
	
    <tr  >
      <td height="25" >手动分页：</td>
	  <td>
	  <INPUT type=button class="bnt" name="appendHtml" value="插入分页符"></td>
	  </tr>
	
        <input name="zdfy" type="hidden" id="zdfy"  value="<%=rs_news("zdfy")%>" size="10" />
   
    <tr  >
      <td  >属性：</td>
	  <td>
      <input name="tj" type="checkbox" id="tj" value="1" <% if rs_news("tj")=1 then response.Write("checked") end if%> class="none">
      <%=xinwen%>幻灯
	  <input name="tt" type="checkbox" id="tt" value="1" <% if rs_news("tt")=1 then response.Write("checked") end if%> class="none">
	  置顶
	  <input name="sytj" type="checkbox" id="sytj" value="1" <% if rs_news("sytj")=1 then response.Write("checked") end if%> class="none">
	  推荐	  
	  <input name="gg" type="checkbox" id="gg" value="1" <% if rs_news("gg")=1 then response.Write("checked") end if%> class="none">
	  头条</td>
    </tr>
	  <tr  >
      <td  >点击次数：</td>
	  <td>
        <input  name="hits" type="text" class="input" size="30" value="<%=rs_news("hits")%>">      </td>
    </tr>
    <tr  >
      <td  >发布人：</td>
	  <td>
        <input  name="user" type="text" class="input" size="30" value="<%=rs_news("user")%>">      </td>
    </tr>
    <tr  >
      <td  >来源：</td>
	  <td>
      <input  name="ly" type="text" id="ly" value="<%=rs_news("ly")%>"></td>
    </tr>
    <tr   >
      <td  >录入时间：</td>
	  <td>
        <input  name="AddDate" type="text" id="AddDate" value="<%=rs_news("AddDate")%>" onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})">      
        </td>
    </tr>
    <tr   >
      <input type="hidden" name="Id" value="<%=Id%>">
      <td colspan=2>
        <div align="left">
          <input class="bnt" type="submit" name="Submit" value="修改" >
      &nbsp;&nbsp; </div></td>
    </tr>
</table>
</form>
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