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
  keyword=Trim(request.form("keyword")) 
zmone=trim(request.form("zmone"))
zmtwo=trim(request.form("zmtwo"))
ly=trim(request.form("ly"))
user=trim(request.form("user"))
zdfy=trim(request.form("zdfy"))
zmthree=trim(request.form("zmtwo"))
Content=trim(request.form("Content"))
System=trim(request.form("System"))
Language=trim(request.form("Language"))
tj=Trim(request.form("tj"))
Softclass=trim(request.form("Softclass"))
PhotoUrl=trim(request.form("PhotoUrl"))
DownloadUrl=trim(request.form("DownloadUrl"))
DownloadUrl1=trim(request.form("DownloadUrl1"))
DownloadUrl2=trim(request.form("DownloadUrl2"))
Hits=trim(request.form("Hits"))
AddDate=trim(request.form("AddDate"))
downkey=trim(request.form("downkey"))
downms=trim(request.form("downms"))
Action=trim(request("Action"))
zdy1=trim(request("zdy1"))
SmallImage=trim(request("SmallImage"))
ImagePath=trim(request.form("ImagePath"))
IcoImage=trim(request.form("Icoimage"))
btcolor=trim(request.form("btcolor"))
if zmone="" then 
 zmone=0
 end if
 if zmtwo="" then 
	zmtwo=0 
	end if
 if zmthree="" then 
	zmthree=0  
end  if

		
		ClassId=trim(Request.Form("ClassId"))
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From xzClass where ID="&ClassId  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then
	  ParentClassID=rs("ParentClassID")  
	  end if
	  rs.close
	  if ParentClassID<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From xzclass where id="&ParentClassID  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID1=rs("ParentClassID")  
	  end if
	  rs.close
	  end if
      if ParentClassID1<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From xzclass where id="&ParentClassID1 
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID2=rs("ParentClassID")     
	  end if
	  rs.close
	  end if
		Set rs = Server.CreateObject("ADODB.Recordset")
sql1="select * from download where id="&id
rs.open sql1,conn,1,3
GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
rs("GroupID")=GroupIdName(0)
rs("Exclusive")=trim(Request.Form("Exclusive"))
rs("Title")=Title
  rs("Content")=Content 
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
  rs("System")=System
rs("Language")=Language
  rs("Softclass")=Softclass
  rs("PhotoUrl")=PhotoUrl
  rs("DownloadUrl")=DownloadUrl 
  rs("DownloadUrl1")=DownloadUrl1
  rs("DownloadUrl2")=DownloadUrl2
  
  '自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=6 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=6  order by Sequence"
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
  rs("AddDate")=AddDate  
  rs("downkey")=downkey
  rs("downms")=downms  
   rs("ly")=ly 
   rs("lang")=lang
  rs("user")=user 
  rs("zdfy")=zdfy
  rs("Hits")=Hits
  rs("SmallImage")=SmallImage
  rs("IcoImage")=IcoImage
  rs("ImagePath")=ImagePath
  rs("keyword")=keyword
  rs("zdy1")=zdy1 
  if tj="1" then
		rs("tj")=1
	else
		rs("tj")=0
	end if
  rs("btcolor")=btcolor
rs.update
rs.close
set rs=nothing




'获取添加新闻的ID  生成html开始
set rs=conn.execute("select id,zmone,zmtwo,zmthree from [Download] where id="&id&"")
xzid=rs(0)
pone=rs(1)
ptwo=rs(2)
pthree=rs(3)
rs.close:set rs=Nothing



if yy=2 then

sql="Select * from Download where id="&xzid&" order by ID desc"
Set Rs=Conn.Execute(sql)
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone")
ptwo=rs("zmtwo")
pthree=rs("zmthree")



'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取下载分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where id = "&pone&" "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If



rs1.close
Set rs1=nothing
End If



'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取下载分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where id = "&ptwo&" "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If



rs1.close
Set rs1=nothing
End If



'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取下载分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where id = "&pthree&" "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&"1.html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&downname&""&Separated&""&pone&""&Separated&""&ID&""&Separated&""&i&".html",""&showdown&"/index.asp","pone=",pone,"downID=",ID,"m=",m,"language=",lang,"page=",i,"","")
next
end If



rs1.close
Set rs1=nothing
End If





If pone<>0 And ptwo=0 And pthree=0 Then
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where ID ="&pone&""
rs1.open sql1,conn,1,1
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from download where zmone="&pone&"")(0)
totalpage=int(totalrec/downinfo)       
If (totalpage * downinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&downclass&""&Separated&""&pone&""&Separated&"1.html",""&down&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&downclass&""&Separated&""&pone&""&Separated&""&i&".html",""&down&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",i)
next
End If


rs1.close:set rs1=Nothing
end if


If pone<>0 And ptwo<>0 And pthree=0 Then
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from xzclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
m=rs2("m")

creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from download where zmone="&pone&" and zmtwo="&ptwo&"")(0)
totalpage=int(totalrec/downinfo)       
If (totalpage * downinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&downclass&""&Separated&""&ptwo&""&Separated&"1.html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&downclass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",i,"","")
next
End If




rs2.close:set rs2=Nothing

rs1.close:set rs1=Nothing
end if

If pone<>0 And ptwo<>0 And pthree<>0 Then
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from xzclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from xzclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from xzclass where ID ="&pthree&""
rs3.open sql3,conn,1,1
m=rs3("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from download where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&"")(0)
totalpage=int(totalrec/downinfo)       
If (totalpage * downinfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&downclass&""&Separated&""&pthree&""&Separated&"1.html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&downclass&""&Separated&""&pthree&""&Separated&""&i&".html",""&down&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m,"language=",lang,"Page=",i)
next
End If
rs3.close:set rs3=Nothing
rs2.close:set rs2=Nothing

rs1.close:set rs1=Nothing
end if
rs.close
set rs=Nothing

call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&down&"/","",""&Downclass&".html",""&down&"/index.asp","m=",mdo,"language=",lang,"","","","","","","","")

end if












call addlog("修改"&xiazai&"")
response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""Down_Manage.asp""</script>"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="javascript" type="text/javascript" src="js/getdate/WdatePicker.js"></script>

		<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
	<!-- #include file="sc.asp" -->
<!--#include file="ys.asp"-->	

		<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
		<script type="text/javascript" src="../javascript/colorbox/jquery.colorbox-min.js"></script>
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
<%
Set rs_down = server.CreateObject("adodb.recordset")
sql = "select * from download where id = "&request.QueryString("id")&" and lang='"&lang&"'"
rs_down.Open sql, conn, 1, 1
If rs_down.EOF Then
    response.Write "<script language=javascript>window.location.href=""Down_Manage.asp""</script>"
    response.End
End If

Exclusive=rs_down("Exclusive")
groupid=rs_down("groupid")
 IcoImage=rs_down("IcoImage")
 ImagePath=rs_down("ImagePath")
if rs_down("zmone")<>0 and  rs_down("zmtwo")=0  and rs_down("zmthree")=0 then
ClassId=rs_down("zmone")
end if
if rs_down("zmone")<>0 and  rs_down("zmtwo")<>0 and rs_down("zmthree")=0  then
ClassId=rs_down("zmtwo")
end if
if rs_down("zmone")<>0 and  rs_down("zmtwo")<>0 and rs_down("zmthree")<>0  then
ClassId=rs_down("zmthree")
end if
%> <form method="POST" name="myform" class="checkform">



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改<%=xiazai%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="xiazai.asp" target="main">返回菜单</a></div>
 
  <div class="dmeun"><a href="down_manage.asp" target="main">下载管理</a></div>
 </td>
 

 </tr>
</table>

      <table width="99%" border="0" cellpadding="6" cellspacing="0" align="center" class="dtable4">
       
     
          <tr> 
            <td width="10%"  ><%=xiazai%>名称：</td>
			<td width="90%"  >
              <input 
name="Title" type="text" class="input" id="Title" value="<%=rs_down("title")%>" size="30" datatype="*" nullmsg="名称不能为空">  </td>
          </tr>
		  
		  <tr >
            <td  >标题颜色</td>
			<td  >
		  <input name="btcolor"type="text" id="color"   value="<%=rs_down("btcolor")%>" size="8" /> 
 <input   type="button" id="colorpicker" value="打开取色器" />
		  </td>
		  
		 </tr>
        <tr >
            <td  >tags:</td>
			<td  >
    <input name="KeyWord" type="text" id="KeyWord" value="<%=rs_down("keyword")%>" size="40"  datatype="*" nullmsg="tags不能为空">
     【<a href="#" id="KeyLinkByTitle" style="color:green">根据标题自动获取</a>】(只对中文有效，多个tags用|隔开)
    </td>
          </tr>
           <tr > 
<td ><%=xiazai%>类别：</td>
			<td  ><select ID="ClassID" name="ClassID" datatype="*" nullmsg="所属分类不能为空">
	   <%call mxlb("xzClass",ClassId)%>
	  </select><font color=red>（有三级栏目的必须选择三级栏目）</font></td>
</tr>
          <tr>
            <td><%=xiazai%>关键词：</td>
			<td  >
              <label>
              <input name="downkey" type="text" id="downkey" value="<%=rs_down("downkey")%>" size="100" />
            </label></td>
          </tr>
          <tr>
            <td  ><%=xiazai%>描述：</td>
			<td  >
              <label>
              <input name="downms" type="text" id="downms" value="<%=rs_down("ms")%>" size="100" />
            </label></td>
          </tr>
		  
		  				  	<%'自定义字段%>
<%call showColumn("Download",6,id)%>
		  
		  
		  
		  

		   <tr>
			    <td >缩略图：</td>
			<td  >
				   <input name="SmallImage"  type="text" id="weepic"  value="<%=rs_down("SmallImage")%>" size="50"  />
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
<div id="mysmaimg"><%if rs_down("Smallimage")<>"" then%>
<img src="../<%=rs_down("SmallImage")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rs_down("Smallimage")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>			        </td>
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
            <td ><%=xiazai%>内容：</td>
			<td  >
              <textarea name="Content" id="Content" style="width:700px;height:300px;visibility:hidden;" datatype="*" nullmsg="内容不能为空"><%=rs_down("content")%></textarea>
            </td>
          </tr>
		   <tr>
      <td >手动分页：</td>
			   <td><INPUT type=button class="bnt" name="appendHtml" value="插入分页符">
			   </td>
          </tr>
          <tr>
 
        <input name="zdfy" type="hidden" id="zdfy"  value="<%=rs_down("zdfy")%>" size="10" />
   
		  
		   <tr  >
	 <td>批量上传图片：</td>
            <td  colspan="2" >
			幻灯图片地址：<input type="text" id="IcoImage" name="IcoImage" value="<%=rs_down("icoimage")%>" style="width:600px;" /><br/><br/>
			<input name="button" type="button" id="J_selectImage" style=" background:url(images/allimg.png) no-repeat; width:170px;height:34px;border:none" value="" />  
			<br/><br/>
         多图片地址：<input type="text" id="ImagePath" name="ImagePath" value="<%=rs_down("ImagePath")%>" style="width:600px;" /><br/><br/>
			
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
<a href="javascript:setContent('<%=img(i)%>')" ><img title="点击图片将图片插入到编辑器" src="<%=img(i)%>" border="0" /></a><br><input type="radio" name="IndexImageradio" onClick="setIndexImage('<%=img(i)%>')" value="<%=img(i)%>" <%if rs_down("icoimage")=img(i) then response.write "checked" %>/>设为幻灯图 <a href="javascript:dropThisDiv('<%=ids%>','<%=img(i)%>')">删除</a></div>
 <%
 next
 end if
 %>    
 </div></td>
            </tr>
		  
		  
		  
		  
		  
		  
          <tr> 
            <td  >适用系统：</td>
			<td  >
            <input 
name="System" type="text" class="input" id="System" value="<%=rs_down("System")%>" size="30"></td>
          </tr>
          <tr>
            <td  >软件语言：</td>
			<td  >
            <input 
name="Language" type="text"  id="Language" value="<%=rs_down("Language")%>" size="30"></td>
          </tr>
          <tr> 
            <td  >软件类型：</td>
			<td  >
              <input 
name="Softclass" type="text" class="input" id="Softclass" value="<%=rs_down("Softclass")%>" size="30">
            </td>
          </tr>

          <tr>
        <td>查看权限：</td>
			<td  >
          <select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> 
          隶属
           <input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
          专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
      </tr>
          <tr > 
            <td  ><%=xiazai%>地址一：</td>
			<td  ><input name="DownloadUrl" type="text" id="url" value="<%=rs_down("DownloadUrl")%>" /> 
            <input type="button" id="insertfile" value="选择文件" /></td>
          </tr>
           <tr > 
            <td  ><%=xiazai%>地址二：</td>
			<td  ><input name="DownloadUrl1" type="text" id="url" value="<%=rs_down("DownloadUrl1")%>" /> 
            </td>
          </tr>
		   <tr > 
            <td  ><%=xiazai%>地址三：</td>
			<td  ><input name="DownloadUrl2" type="text" id="url" value="<%=rs_down("DownloadUrl2")%>" /> 
            </td>
          </tr>
          <tr>
            <td  >文件大小：</td>
			<td  >
              <label>
              <input name="zdy1" type="text" id="zdy1" value="<%=rs_down("zdy1")%>" size="5" />
            </label></td>
          </tr>
		  <tr>
               <td >是否推荐:</td>
			<td  ><input name="tj" type="checkbox" id="tj" value="1" <% if rs_down("tj")=1 then response.Write("checked") end if%> /></td>
             </tr>
			 
			 			  <tr > 
            <td  >点击次数：</td>
			<td>
              <input  name="hits" type="text" class="input" value="<%=rs_down("hits")%>"  size="30"> 
			  </td>
          </tr>
			 
			 
			 
			 
			 
			 
		  <tr  >
      <td  >发布人：</td>
			<td  >
        <input  name="user" type="text" class="input" size="30" value="<%=rs_down("user")%>">      </td>
    </tr>
    <tr  >
      <td  >来源：</td>
			<td  >
      <input  name="ly" type="text" id="ly" value="<%=rs_down("ly")%>"></td>
    </tr>
          <tr> 
            <td  >录入时间：</td>
			<td  >
            <input 
name="AddDate" type="text" id="AddDate" value="<%=rs_down("AddDate")%>"  onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})" ></td>
          </tr>
          <tr > 
            <td colspan="2" > <input 
type="submit" name="Submit" value="修改" class="bnt"> 
              
              　            </td>
          </tr>
</table>
  </form>
      <% 
rs_down.close
set rs_down=nothing
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