<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="htmlconfig.asp"-->
<!--#include file="admin_html_function.asp"-->
<%
id = checknum(request("id"),0)
id1 = request.QueryString("id")
Title=Trim(request.form("Title")) 
keyword=Trim(request.form("keyword")) 
Product_Id=Trim(request.form("Product_Id"))
'会员组级别价格S
users_price = safereplace(trim(request("users_price")))
users_grade = safereplace(trim(request("users_grade")))
'会员组级别价格E
'自定义商品规格开始
sprice=CheckNum(trim(request("sprice")),0)   '市场价
cprice=CheckNum(trim(request("cprice")),0)    '促销价
p_weight=CheckNum(request("p_weight"),0)  '产品重量
		total_kucun=CheckNum(request("total_kucun"& id),0)  '总的库存
	    categoryid = replace(safereplace(request("categoryid"))," ","") '产品类别

        actionkucun=CheckNum(request("actionkucun"),0) '库存预警
	
	
	
		p_SelectValue = replace(Replace(Trim(request("p_SelectValue")),"'","&#180;")," ","")  '筛选属性
		p_SelectType = checknum(request("SelectType"),0)  '附属属性ID
		SpecificationCount = checknum(request("SpecificationCount"),0)  '组合数量
		If SpecificationCount > 0 Then  
				kucun = total_kucun
		End If 
		P_SpecificationName = safereplace(request("P_SpecificationName"))
		P_SpecificationValue = safereplace(request("P_SpecificationValue"))
		P_SpecificationPic = safereplace(request("P_SpecificationPic"))
		If P_SpecificationPic <> "" Then 
				If InStr(P_SpecificationPic,"|") > 0 Then 
						Arr_P_SpecificationPic = Split(P_SpecificationPic,"|")
						For i = 0 To UBound(Arr_P_SpecificationPic)
								If InStr(cpic,Arr_P_SpecificationPic(i)) = 0 Then 
										cpic = cpic &"@@"& Arr_P_SpecificationPic(i)
								End If 
						Next 
				Else
						If InStr(cpic,P_SpecificationPic) = 0 Then 
								cpic = cpic &"@@"& P_SpecificationPic
						End If 
				End If 
		End If 
		
		
		
		'自定义商品规格结束



zmone=trim(request.form("zmone"))
zmtwo=trim(request.form("zmtwo"))
zmthree=trim(request.form("zmthree"))
scprice=trim(request.form("scprice"))
hyprice=trim(request.form("hyprice"))
cbprice=trim(request.form("cbprice"))
 kucun1=request.form("kucun1") 
tj=trim(request.form("tj"))
btcolor=trim(request.form("btcolor"))
Content=Trim(request.form("Content"))
Passed=trim(request.form("Passed"))
Newproduct=trim(request.form("Newproduct"))
Hits=trim(request.form("Hits"))
adddate=trim(request.form("adddate"))
productkey=trim(request.form("productkey"))
productms=trim(request.form("productms"))
zdfy=trim(request("zdfy"))
User=Trim(request.form("User"))
ly=Trim(request.form("ly"))
SmallImage=trim(request.form("SmallImage"))
ImagePath=trim(request.form("ImagePath"))
IcoImage=trim(request.form("Icoimage"))
sellbuys=trim(request.form("sellbuys"))


If request.Form("submit") = "修改" Then


if hyprice="" then
hyprice=0
end if
if scprice="" then
scprice=0
end if

ClassId=trim(Request.Form("ClassId"))
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From cpClass where ID="&ClassId  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then
	  ParentClassID=rs("ParentClassID")  
	  end if
	  rs.close
	  if ParentClassID<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From cpclass where id="&ParentClassID  
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID1=rs("ParentClassID")  
	  end if
	  rs.close
	  end if
      if ParentClassID1<>0 then
	  Set rs=server.CreateObject("adodb.recordset")
	  sql="Select * From cpclass where id="&ParentClassID1 
	  rs.open sql,conn,1,1
	  if Not rs.bof and Not rs.eof then  
	  ParentClassID2=rs("ParentClassID")     
	  end if
	  rs.close
	  end if
Set rsProduct = Server.CreateObject("ADODB.Recordset")
sql1="select * from product where id="&id
rsProduct.open sql1,conn,1,3
GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
rsProduct("GroupID")=GroupIdName(0)
rsProduct("Exclusive")=trim(Request.Form("Exclusive"))
rsProduct("Product_Id")=Product_Id
if ParentClassID=0 then
rsProduct("zmone")=trim(Request.Form("ClassId"))
rsProduct("zmtwo")=0
rsProduct("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1=0 then
rsProduct("zmone")=ParentClassID
rsProduct("zmtwo")=trim(Request.Form("ClassId"))
rsProduct("zmthree")=0
end if
if ParentClassID<>0 and ParentClassID1<>0 and ParentClassID2=0 then
rsProduct("zmone")=ParentClassID1
rsProduct("zmtwo")=ParentClassID
rsProduct("zmthree")=trim(Request.Form("ClassId"))
end if
rsProduct("Title")=Title
rsProduct("hyprice")=hyprice
rsProduct("cbprice")=cbprice
rsProduct("kucun1")=kucun1
rsProduct("lang")=lang
'自定义商品规格开始
                rsProduct("pro_name")=Replace(request("pro_name"),"'","&#180;") '商品规格
				rsProduct("pro_info")=Replace(request("pro_info"),"'","&#180;") '商品属性

rsProduct("sprice")=sprice
		rsProduct("cbprice")=cbprice
        rsProduct("cprice")=cprice


		rsProduct("p_weight")=p_weight
		rsProduct("categoryid") = ","& categoryid &","
		
		rsProduct("p_order") = p_order
		rsProduct("actionkucun") = actionkucun
	

		rsProduct("p_SelectValue") = ","& p_SelectValue &","  '筛选属性值
		rsProduct("p_SelectType") = p_SelectType     '附属属性分类ID
		rsProduct("P_SpecificationName") = P_SpecificationName   '规格ID
		rsProduct("P_SpecificationValue") = P_SpecificationValue  '规格组合
		rsProduct("P_SpecificationPic") = P_SpecificationPic  '颜色尺寸图片
		
		
		'自定义商品规格结束










'自定义字段开始
set rsObj=server.CreateObject("adodb.recordset")
'Sqlp1 ="select * from Mycolumn where TypeID=3 and lang='"&lang&"' order by Sequence"
Sqlp1 ="select * from Mycolumn where TypeID=3  order by Sequence"
rsObj.open sqlp1,conn,1,1
if not rsObj.eof then
do while not rsObj.eof
zdyzd1=rsObj("ColumnName")
zdyzd2=request.form(""&zdyzd1&"")
rsProduct(""&zdyzd1&"")=zdyzd2
rsObj.movenext
loop
end if
'自定义字段结束
rsProduct("scprice")=scprice
rsProduct("keyword")=keyword
rsProduct("content")=content
rsProduct("productkey")=productkey
rsProduct("productms")=productms
rsProduct("adddate")=adddate	
rsProduct("zdfy")=zdfy
rsProduct("btcolor")=btcolor
rsProduct("hits")=hits
rsProduct("sellbuys")=sellbuys
rsProduct("ly")=ly
rsProduct("user")=user
rsProduct("SmallImage")=SmallImage
rsProduct("IcoImage")=IcoImage
rsProduct("ImagePath")=ImagePath
if Passed="1" then
rsProduct("Passed")=1
else
rsProduct("Passed")=0
end if
if tj="1" then
rsProduct("tj")=1
else
rsProduct("tj")=0
end if
if Newproduct="1" then
rsProduct("Newproduct")=1
else
rsProduct("Newproduct")=0
end if
Num_1=CheckStr(Request.Form("Num_1"),1)
if Num_1="" then Num_1=0
if Num_1>0 then
For i=1 to Num_1
If CheckStr(Request.Form("attribute"&i),0)<>"" and  CheckStr(Request.Form("attribute"&i&"_value"),0)<>"" Then
If attribute1="" then
attribute1=CheckStr(Request.Form("attribute"&i),0)
attribute1_value=CheckStr(Request.Form("attribute"&i&"_value"),0)
Else
attribute1=attribute1&"§§§"&CheckStr(Request.Form("attribute"&i),0)
attribute1_value=attribute1_value&"§§§"&CheckStr(Request.Form("attribute"&i&"_value"),0)
End if
End If
Next
end if
rsProduct("attribute1")=attribute1
rsProduct("attribute1_value")=attribute1_value
rsProduct.update
rsProduct.close
set rsProduct=nothing

If users_price = "" Then users_price = checknum(request("u_price"),0)
		If users_grade = "" Then users_grade = 1
		Call updateUsersPrice(id,users_price,users_grade,lang)

		If safereplace(request("SP_no")) = "" Then 
				SpecificationCount = 0
		Else 
				If InStr(request("SP_no"),",") = 0 Then 
						SpecificationCount = 1
				Else
						SpecificationCount = UBound(Split(request("SP_no"),","))+1
				End If 
		End If 
		If SpecificationCount > 0 Then 
		For i = 0 To SpecificationCount-1
				If CLng(SpecificationCount-1) = 0 Then 
						p_no = safereplace(request("SP_no"))
						s_value = request("SP_value")
						p_stock = checknum(request("SP_stock"),0)
						p_price = checknum(request("SP_price"),0)
						p_cbprice = checknum(request("SP_cbprice"),0)
						p_weight = checknum(request("SP_weight"),0)
						s_id = checknum(request("SP_id"),0)
				Else
						p_no = safereplace(Split(request("SP_no"),",")(i))
						s_value = Split(request("SP_value"),",")(i)
						p_stock = checknum(Split(request("SP_stock"),",")(i),0)
						p_price = checknum(Split(request("SP_price"),",")(i),0)
						p_cbprice = checknum(Split(request("SP_cbprice"),",")(i),0)
						p_weight = checknum(Split(request("SP_weight"),",")(i),0)
						s_id = checknum(Split(request("SP_id"),",")(i),0)
				End If 
				sql = "select * from [Product_Stock]"
				If s_id <> 0 Then 
						sql = sql &" where id = "& s_id
				Else 
						If i = 0 Then 
								conn.execute("delete from [Product_Stock] where p_id = "& id &"")
						End If 
				End If 
				set rs = server.CreateObject("adodb.recordset")
				rs.open sql, conn, 1, 3
				If rs.eof Or s_id = 0 Then
						rs.addnew
						rs("p_id") = id
						For LangT = 1 To TotalLang
								If InStr(s_value,"@") = 0 Then 
										If InStr(s_value,"$$$") = 0 Then 
												rs("s_value" & LanguageTxt(LangT))=safereplace(s_value)
										Else
												rs("s_value" & LanguageTxt(LangT))=safereplace(Split(s_value,"$$$")(LangT-1))
										End If 
								Else
										tmp_s_value = ""
										For ii = 0 To UBound(Split(s_value,"@"))
												If InStr(Split(s_value,"@")(ii),"$$$") = 0 Then 
														tmp_s_value = tmp_s_value & safereplace(Split(s_value,"@")(ii))
												Else
														tmp_s_value = tmp_s_value & safereplace(Split(Split(s_value,"@")(ii),"$$$")(LangT-1))
												End If 
												If ii < UBound(Split(s_value,"@")) Then tmp_s_value = tmp_s_value & ","
										Next
										rs("s_value" & LanguageTxt(LangT))=tmp_s_value
								End If 
						Next
				End If 
				rs("p_no") = p_no
				rs("p_stock") = p_stock
				rs("p_price") = p_price
				rs("p_cbprice") = p_cbprice
				rs("p_weight") = p_weight
				 rs("lang")=lang
				rs.update 
				rs.close : Set rs = Nothing
		Next 
		End If 




'获取添加新闻的ID  生成html开始
set rs=conn.execute("select id,zmone,zmtwo,zmthree from [product] where id="&id1&"")
proid=rs(0)
pone=rs(1)
ptwo=rs(2)
pthree=rs(3)
rs.close:set rs=Nothing
if yy=2 then
'生成添加新闻的html
sql="Select * from product where id="&proid&" order by ID desc"
Set Rs=Conn.Execute(sql)
content=rs("content")
zdfy=rs("zdfy")
pone=rs("zmone") '一级分类
ptwo=rs("zmtwo") '二级分类
pthree=rs("zmthree") '三级分类


'一级分类开始
If pone<>0 And ptwo=0 And pthree=0 Then

'获取产品分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where id = "&pone&" "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&id1&""&Separated&"1.html",""&showproduct&"/index.asp","pone=",pone,"proID=",id1,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&id1&""&Separated&""&i&".html",""&showproduct&"/index.asp","pone=",pone,"proID=",id1,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If





'二级分类开始
If pone<>0 And ptwo<>0 And pthree=0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where id = "&ptwo&" "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&id1&""&Separated&"1.html",""&showproduct&"/index.asp","pone=",pone,"proID=",id1,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&id1&""&Separated&""&i&".html",""&showproduct&"/index.asp","pone=",pone,"proID=",id1,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If




'三级分类开始
If pone<>0 And ptwo<>0 And pthree<>0 Then

'获取新闻分类的高亮值 并循环生成html
Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where id = "&pthree&" "
rs1.open sql1,conn,1,1
m=rs1("m")


if Instr(content,"[zmcmsfy]")=0 then
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&id1&""&Separated&"1.html",""&showproduct&"/index.asp","pone=",pone,"proID=",id1,"m=",m,"language=",lang,"page=",1,"","")
else
arrContent=split(content,"[zmcmsfy]")
pages=Ubound(arrContent)+1
for i=1 to pages
call htmll("/html/","",""&ProductName&""&Separated&""&pone&""&Separated&""&id1&""&Separated&""&i&".html",""&showproduct&"/index.asp","pone=",pone,"proID=",id1,"m=",m,"language=",lang,"page=",i,"","")
next
end If

rs1.close
Set rs1=nothing
End If










'生成所属所属分类的新闻

If pone<>0 And ptwo=0 And pthree=0 Then

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where ID ="&pone&""
rs1.open sql1,conn,1,1
m=rs1("m")
creatwjj("../"&rs1("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from product where zmone="&pone&"")(0)
totalpage=int(totalrec/ProductInfo)       
If (totalpage * ProductInfo)<totalrec Then
totalpage=totalpage+1
End If
creatwjj("../"&rs1("mulu")&"/")
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/","",""&Productclass&""&Separated&""&pone&""&Separated&"1.html",""&product&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/","",""&Productclass&""&Separated&""&pone&""&Separated&""&i&".html",""&product&"/index.asp","pone=",pone,"m=",m,"ptwo=","","pthree=","","language=",lang,"Page=",i)
next
End If


rs1.close:set rs1=Nothing

end if



If pone<>0 And ptwo<>0 And pthree=0 Then

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from cpclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
m=rs2("m")

creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/")

totalrec=Conn.Execute("Select count(*) from product where zmone="&pone&" and zmtwo="&ptwo&"")(0)
totalpage=int(totalrec/ProductInfo)       
If (totalpage * ProductInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&Productclass&""&Separated&""&ptwo&""&Separated&"1.html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",1,"","")
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/","",""&Productclass&""&Separated&""&ptwo&""&Separated&""&i&".html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"m=",m,"language=",lang,"Page=",i,"","")
next
End If




rs2.close:set rs2=Nothing

rs1.close:set rs1=Nothing
end if


If pone<>0 And ptwo<>0 And pthree<>0 Then

Set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select * from cpclass where ID ="&pone&""
rs1.open sql1,conn,1,1
Set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select * from cpclass where ID ="&ptwo&""
rs2.open sql2,conn,1,1
Set rs3=Server.CreateObject("ADODB.Recordset")
sql3="select * from cpclass where ID ="&pthree&""
rs3.open sql3,conn,1,1
m=rs3("m")
creatwjj("../"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/")
totalrec=Conn.Execute("Select count(*) from product where zmone="&pone&" and zmtwo="&ptwo&" and zmthree="&pthree&"")(0)
totalpage=int(totalrec/ProductInfo)       
If (totalpage * ProductInfo)<totalrec Then
totalpage=totalpage+1
End If
if totalpage<=1 then
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&Productclass&""&Separated&""&pthree&""&Separated&"1.html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m,"language=",lang,"Page=",1)
else
for i=1 to totalpage
call htmll("/"&rs1("mulu")&"/"&rs2("mulu")&"/"&rs3("mulu")&"/","",""&Productclass&""&Separated&""&pthree&""&Separated&""&i&".html",""&product&"/index.asp","pone=",pone,"ptwo=",ptwo,"pthree=",pthree,"m=",m,"language=",lang,"Page=",i)
next
End If
rs3.close:set rs3=Nothing
rs2.close:set rs2=Nothing
rs1.close:set rs1=Nothing
end If





rs.close
set rs=Nothing








call htmll("","","default"&lang&".html","default.asp","m=",mde,"language=",lang,"","","","","","","","")
call htmll("/"&product&"/","",""&Productclass&".html",""&product&"/index.asp","m=",mpro,"language=",lang,"","","","","","","","")
end if











call addlog("修改"&chanpin&"")

response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""productmanage.asp""</script>"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<script language="javascript" type="text/javascript" src="js/getdate/WdatePicker.js"></script>
<style type="text/css">

.ys td{ padding:10px}
 .imgDiv { float:left; padding-top:5px; padding-left:5px; }
 .imgDiv img{ border:0px; width:120px; height:90px}
</style>
<script language="javascript" type="text/javascript" src="js/getdate/WdatePicker.js"></script>
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>

<script language="javascript" type="text/javascript" src="js/product.js"></script>
<script language="javascript" type="text/javascript" src="js/admin.js"></script>



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
<script>var AdminPath = "/admin/";</script>
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
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
<!--#include file="hqyy.asp"-->

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

<SCRIPT language=javaScript>
function Checkly()
{
getSubmit();  
  

if (document.forms['myform'].elements['u_price'] != undefined)
	{
		var obj = document.forms['myform'].elements['u_price'];
		var up = "";
		var ug = "";
		//alert(obj.length);
		for(var i = 0; i < obj.length; i++)
		{
			
				up = up + obj[i].value +",";
				ug = ug + uGradeArr[i][0] +",";
		}
		document.forms['myform'].elements['users_price'].value = up;
		document.forms['myform'].elements['users_grade'].value = ug;
	}
 
 }
</SCRIPT>




<script language="javascript">
var Lang_Product = "商品编号$库存$加价$成本价$删除";
</script>
<script language="javascript">


function getSubmit()
{

	var p_ProductName = document.getElementsByName("p_ProductName");
	var p_ProductValueType = document.getElementsByName("p_ProductValueType");
	var p_SelectValue = document.getElementsByName("ajax_SelectValue");
	var pro_name_tArray = new Array();
	var pro_value_type_tArray = new Array();
	var p_SelectValue_param = new Array();
	for(i=0;i<p_ProductName.length;i++){pro_name_tArray.push(p_ProductName[i].value);}
	document.getElementById("pro_name").value=pro_name_tArray.join("$");
	for(i=0;i<p_ProductValueType.length;i++){pro_value_type_tArray.push(p_ProductValueType[i].value);}
	var pro_value_type = pro_value_type_tArray.join(",");
	for(i=0;i<p_SelectValue.length;i++){if (p_SelectValue[i].checked){p_SelectValue_param+=p_SelectValue[i].value +",";}}

	if (p_SelectValue_param.length > 0){document.getElementById("p_SelectValue").value=p_SelectValue_param.substring(0,p_SelectValue_param.length-1);}else{document.getElementById("p_SelectValue").value="";}

	var str = "";
	for(i=0;i<p_ProductName.length;i++){
		if (pro_value_type.split(',')[i] == 0){str += document.getElementById("p_ProductValue_"+i+"").value+"$$$";}else{
		var str_list = "",list_list = document.getElementsByName("p_ProductValue_"+i+"");
		for(k=0;k<list_list.length;k++){if(list_list[k].checked){str_list+=list_list[k].value+",";}}
		str += str_list.substring(0,str_list.length-1)+"$$$";
		}
	}
	document.getElementById("pro_info").value=str;
	var P_SpecificationPic = document.getElementsByName("SValuePic");
	var P_SpecificationPic_tArray = new Array();
	for(i=0;i<P_SpecificationPic.length;i++){P_SpecificationPic_tArray.push(P_SpecificationPic[i].value);}
	document.getElementById("P_SpecificationPic").value=P_SpecificationPic_tArray.join("|");
	var Temp_SpecificationName = document.getElementsByName("Hidden_SName");
	var Temp_SpecificationName_tArray = new Array();
	var Temp_SpecificationValue_tArray = new Array();
	var Temp_SpecificationName_value = "";
	var Temp_SpecificationValue_value = "";
	var Temp_SpecificationValue;
	var sname = "";
	var svalue = "";
	for(i=0;i<Temp_SpecificationName.length;i++){if(Temp_SpecificationName[i].checked==true){Temp_SpecificationName_tArray.push(Temp_SpecificationName[i].value);   }}
	Temp_SpecificationName_value=Temp_SpecificationName_tArray.join(",");
	if (Temp_SpecificationName_value != "")
	{
		if (Temp_SpecificationName_value.indexOf(",")==-1)
		{
			sname += Temp_SpecificationName_value.split("|")[0];
			Temp_SpecificationValue = document.getElementsByName("SValue"+Temp_SpecificationName_value.split("|")[0]+"");
			for(v=0;v<Temp_SpecificationValue.length;v++){if(Temp_SpecificationValue[v].checked==true){svalue += Temp_SpecificationValue[v].value.split("|")[0] +"|"; }}
			document.getElementById("P_SpecificationName").value=sname;
			document.getElementById("P_SpecificationValue").value=svalue;
		}else{
			for(i=0;i<Temp_SpecificationName_value.split(",").length;i++)
			{
				sname += Temp_SpecificationName_value.split(",")[i].split("|")[0]; 
				Temp_SpecificationValue = document.getElementsByName("SValue"+Temp_SpecificationName_value.split(",")[i].split("|")[0]+"");
				for(s=0;s<Temp_SpecificationValue.length;s++){if(Temp_SpecificationValue[s].checked==true){Temp_SpecificationValue_tArray.push(Temp_SpecificationValue[s].value);   }}
				Temp_SpecificationValue_value=Temp_SpecificationValue_tArray.join(",");
				if (Temp_SpecificationValue_value.indexOf(",")==-1){
					for(v=0;v<Temp_SpecificationValue.length;v++){if(Temp_SpecificationValue[v].checked==true){svalue += Temp_SpecificationValue[v].value.split("|")[0]; if (v<Temp_SpecificationValue.length-1){svalue += ",";}}}
					if (svalue.charAt(svalue.length - 1) == ","){svalue = svalue.substring(0,svalue.length-1);}
				}else{
					for(v=0;v<Temp_SpecificationValue.length;v++){if(Temp_SpecificationValue[v].checked==true){svalue += Temp_SpecificationValue[v].value.split("|")[0]; if (v<Temp_SpecificationValue.length-1){svalue += ",";}}}
					if (svalue.charAt(svalue.length - 1) == ","){svalue = svalue.substring(0,svalue.length-1);}
				}
				if (i<Temp_SpecificationName_value.split(",").length-1){sname += ",";svalue += "|";}
			}
			document.getElementById("P_SpecificationName").value=sname;
			document.getElementById("P_SpecificationValue").value=svalue;
		}
	}
	var SP_id = document.getElementsByName("ajax_s_id");
	var SP_id_value = "";
	var SP_no = document.getElementsByName("ajax_p_no");
	var SP_no_value = "";
	var SP_value = document.getElementsByName("ajax_s_value");
	var SP_value_value = "";
	var SP_stock = document.getElementsByName("ajax_p_stock");
	var SP_stock_value = "";
	var SP_price = document.getElementsByName("ajax_p_price");
	var SP_price_value = "";
	var SP_cbprice = document.getElementsByName("ajax_p_cbprice");
	var SP_cbprice_value = "";
	var SP_weight = document.getElementsByName("ajax_p_weight");
	var SP_weight_value = "";
	for(i=0;i<SP_id.length;i++){SP_id_value += SP_id[i].value; if (i<SP_id.length-1){SP_id_value += ",";}}
	for(i=0;i<SP_no.length;i++){SP_no_value += SP_no[i].value; if (i<SP_no.length-1){SP_no_value += ",";}}
	for(i=0;i<SP_value.length;i++){SP_value_value += SP_value[i].value; if (i<SP_value.length-1){SP_value_value += ",";}}
	for(i=0;i<SP_stock.length;i++){SP_stock_value += SP_stock[i].value; if (i<SP_stock.length-1){SP_stock_value += ",";}}
	for(i=0;i<SP_price.length;i++){SP_price_value += SP_price[i].value; if (i<SP_price.length-1){SP_price_value += ",";}}
	for(i=0;i<SP_cbprice.length;i++){SP_cbprice_value += SP_cbprice[i].value; if (i<SP_cbprice.length-1){SP_cbprice_value += ",";}}
	for(i=0;i<SP_weight.length;i++){SP_weight_value += SP_weight[i].value; if (i<SP_weight.length-1){SP_weight_value += ",";}}
	document.getElementById("SP_id").value = SP_id_value;
	document.getElementById("SP_no").value = SP_no_value;
	document.getElementById("SP_value").value = SP_value_value;
	document.getElementById("SP_stock").value = SP_stock_value;
	document.getElementById("SP_price").value = SP_price_value;
	document.getElementById("SP_cbprice").value = SP_cbprice_value;
	document.getElementById("SP_weight").value = SP_weight_value;
}
</script>

</head>

<body>
<%
SpecificationCount = 0
IsSelectValue = 0
selecttype = 0
sprice = "0.1"
price = "0.1"
cbprice = "0"
p_weight = "0"
DataCount = 0
kucun="1000"
call setUsersPrice(id)
Set rs = conn.execute("select s_value from [Product_Stock] where p_id = "& id &"")
If Not rs.eof Then 
DataCount = 1 
DataValue = rs.getrows()
End If 
rs.close : Set rs = Nothing
If DataCount = 1 Then 
		For i = 0 To UBound(DataValue,2)
				s_value = s_value & DataValue(0,i) &","
		Next 
End If 




Set rsProduct = server.CreateObject("adodb.recordset")
sql2 = "select * from product where id = "&request.QueryString("id")&" and lang='"&lang&"'"
rsProduct.Open sql2, conn, 1, 1
If rsProduct.EOF Then
response.Write "<script language=javascript>window.location.href=""productmanage.asp""</script>"
response.End
End If
Exclusive=rsProduct("Exclusive")
groupid=rsProduct("groupid")
if rsProduct("attribute1")<>"" and rsProduct("attribute1_value")<>"" then
attribute1_1=Split(rsProduct("attribute1"),"§§§")
attribute1_value_1=Split(rsProduct("attribute1_value"),"§§§")
Num_1=ubound(attribute1_1)+1
Else
Num_1=0
End If
if rsProduct("zmone")<>0 and  rsProduct("zmtwo")=0  and rsProduct("zmthree")=0 then
ClassId=rsProduct("zmone")
end if
if rsProduct("zmone")<>0 and  rsProduct("zmtwo")<>0 and rsProduct("zmthree")=0  then
ClassId=rsProduct("zmtwo")
end if
if rsProduct("zmone")<>0 and  rsProduct("zmtwo")<>0 and rsProduct("zmthree")<>0  then
ClassId=rsProduct("zmthree")
end if

 Content=rsProduct("Content")
 IcoImage=rsProduct("IcoImage")
 ImagePath=rsProduct("ImagePath")
 
 SelectType = checknum(rsProduct("p_SelectType"),0)
%>
<form method="POST" name="myform"  onSubmit="return Checkly()" class="checkform">


<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">修改<%=chanpin%>  <span id="yuyan" style="margin-left:20px"></span></div></td>
</tr>
</table>



<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="chanpin.asp" target="main">返回菜单</a></div>
 
  <div class="dmeun"><a href="productmanage.asp" target="main"><%=chanpin%>管理</a></div>
 
 </td>
 

 </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="0"  align="center" class="dtable4">
         
		   <tr > 
            <td width="12%"  ><%=chanpin%>名称：</td>
			<td width="88%" >
              <input name="Title" type="text" id="Title" value="<%=rsProduct("title")%>" size="50"  datatype="*" nullmsg="名称不能为空"/>
             <font color="#FF0000">*</font>
			 </td></tr>
			 <tr > 
            <td  >标题颜色：</td>
			<td > 
			 
			 <input name="btcolor" type="text" id="color"   value="<%=rsProduct("btcolor")%>" size="8" /> 
 <input   type="button" id="colorpicker" value="打开取色器" /> </td>
          </tr>
		  
          <tr >
            <td >tags:</td>
			
			<td>
    <input name="KeyWord" datatype="*" nullmsg="tags不能为空" type="text" id="KeyWord" value="<%=rsProduct("keyword")%>" size="40" > 【<a href="#" id="KeyLinkByTitle" style="color:green">根据标题自动获取</a>】(只对中文内容生效,多个tags用|隔开)

   
    </td>
          </tr>
                     <tr > 
<td><%=chanpin%>类别：</td>
<td>
<select ID="ClassID" name="ClassID" datatype="*" nullmsg="所属分类不能为空">
	   <%call mxlb("cpClass",ClassId)%>
	  </select> <font color=red>（有三级栏目的必须选择三级栏目）</font></td>
</tr>
          <tr > 
            <td ><%=chanpin%>编号：</td> 
			<td><input name="Product_Id" type="text" id="Product_Id" value="<%=rsProduct("Product_Id")%>" size="9" > 
              <font color="#FF0000">*</font> </td>
          </tr>
		 <tr >
  <td   ><%=chanpin%>关键词：</td> 
			<td>
    <input name="productkey" type="text" id="productkey" value="<%=rsProduct("productkey")%>" size="100" /></td>
</tr>
<tr >
  <td   ><%=chanpin%>描述：</td> 
			<td>
    <input name="productms" type="text" id="productms" value="<%=rsProduct("productms")%>" size="100" /></td>
</tr>
          <tr > 
            <td   >市场价格：</td> 
			<td>
              <input name="scprice" type="text" id="scprice" value="<%=rsProduct("scprice")%>"></td>
          </tr>
		  
		  <tr class="classrow">

<td class="classrow" ><input type="hidden" id="sprice" name="sprice" size="6" value="<%= sprice %>"  class="form" onKeyUp="value=value.replace(/[^\d\.]/g,'');"> 
&nbsp;&nbsp;成本价：</td><td><input type="text" id="cbprice" name="cbprice"   value="<%=rsProduct("cbprice")%>" class="form" /> </td></tr>
		  
          <tr >
            <td   >购买价格：</td> 
			<td>
              <input name="hyprice" type="text" id="hyprice" value="<%=rsProduct("hyprice")%>" /></td>
          </tr>
		  
		  
<tr >
<td   >会员组价格设置：</td> 
			<td>(如果需要购物这块功能建议 会员组价格这块必须设置！不需要购物功能，可以不设置)</br></br>

<SCRIPT LANGUAGE="JavaScript">setPrice();</SCRIPT><INPUT TYPE="hidden" NAME="users_price"><INPUT TYPE="HIDDEN" NAME="users_grade">

</td>
</tr> 
		  
		  
		  
		<tr>
<td >已销售数量:</td>
<td>
  <input type="text" name="sellbuys"  size="5"   value="<%= rsproduct("sellbuys") %>"/></td>
</tr>  
		  
		  
		  
		  <tr class="classrow" >
<td  class="classrow">库存数量:</td> 
			<td>
  <input type="text" name="kucun1" id="kucun1" size="5"  value="<%= rsproduct("kucun1") %>" ></td>
</tr>	  
		  
		  
		<tr class="classrow"  style="display:none">
<td  class="classrow">库存数量:</td> 
			<td>
  <input type="text" name="kucun" id="kucun<%=id%>" size="5"  value="1000" class="form" onKeyUp="value=value.replace(/[^\d\.]/g,'');">  <input type="hidden" name="total_kucun<%=id%>" id="total_kucun<%=id%>" value="1000" /></td>
</tr>	  
		  

		  
		  
  
		  
		  
		  
		  
		  
	<!-- 属性类别开始-->
<tr class="classrow">
<td  class="classrow">请选择属性类别:</td> 
			<td>
  <select name="SelectType" id="SelectType" onChange="getSelectValue(0,'',<%=id%>,0,'');getSelectValue(0,'',<%=id%>,1,'');getSelectValue(0,'',<%=id%>,2,'');">
  <%
sql = "select s_name,s_id from [SelectType] where lang='"&lang&"' order by s_order desc"
Set rs = conn.execute(sql)
If Not rs.eof Then 
	IsSelectValue = 1
	DataSelectValue = rs.getrows()
End If 
rs.close : Set rs = Nothing 
Response.Write("<option value="""">┌ 属性类别 </option>")
If IsSelectValue = 1 Then 
For i = 0 To UBound(DataSelectValue,2)
Arr_SelectValue = Split(DataSelectValue(1,i),Chr(13)&Chr(10))
For ii = 0 To UBound(Arr_SelectValue)
	Response.Write("<option value="""& DataSelectValue(1,i) &""" "& judge2(DataSelectValue(1,i),SelectType) &">├ "& DataSelectValue(0,i) &"</option>")
Next 
Next 
End If 
%>
  </select></td>
</tr>

<tr id="SpecificationValueBox" class="classrow" style="display:none">
<td height="30" colspan="2"  class="classrow">
<input type="hidden" name="P_SpecificationName" id="P_SpecificationName" value="<%=P_SpecificationName%>" />
<input type="hidden" name="P_SpecificationValue" id="P_SpecificationValue" value="<%=P_SpecificationValue%>" />
<input type="hidden" name="P_SpecificationPic" id="P_SpecificationPic" value="<%=P_SpecificationPic%>" />
<input type="hidden" name="SpecificationCount" id="SpecificationCount" value="<%=checknum(conn.execute("select count(id) from [Product_Stock] where p_id = "& id &"")(0),0)%>" />
<input type="hidden" name="SP_id" id="SP_id" />
<input type="hidden" name="SP_no" id="SP_no" />
<input type="hidden" name="SP_value" id="SP_value" />
<input type="hidden" name="SP_stock" id="SP_stock" />
<input type="hidden" name="SP_price" id="SP_price" />
<input type="hidden" name="SP_cbprice" id="SP_cbprice" />
<input type="hidden" name="SP_weight" id="SP_weight" />
<span id="SpecificationValue"><script type="text/javascript">getSelectValue(<%=SelectType%>,'<%=rsProduct("P_SpecificationName")%>$$$<%=rsProduct("P_SpecificationValue")%>',<%=id%>,1,'');</script>
</span>
<span id="Get_SpecificationPic<%=id%>"></span>
<span id="Get_Specification<%=id%>"></span>
<table id="Table_Specification" cellpadding="0" cellspacing="0" width="100%" class="tablelist" style="display:none">
</table><script type="text/javascript">getspecificationpic(<%=id%>,'<%=rsProduct("P_SpecificationValue")%>','<%=rsProduct("P_SpecificationPic")%>');getspecification(<%=id%>,'<%=rsProduct("P_SpecificationName")%>',0);</script></td>
</tr>

<!-- 属性类别结束-->
<tr id="ProductValueBox" style="display:none">
<td   class="classrow">商品属性：</td> 
			<td>
  <input type="hidden" name="pro_name" id="pro_name">
  <input type="hidden" name="pro_info" id="pro_info">
  <span id="ProductValue"><script type="text/javascript">getSelectValue(<%=SelectType%>,'<%=rsProduct("pro_name")%>@@<%=rsProduct("pro_info")%>',0,2,'');</script>
  </span></td>
</tr>

<tr id="SelectValueBox" class="classrow" style="display:none">
<td class="classrow" >筛选属性:</td> 
			<td>
  <input type="hidden" name="p_SelectValue" id="p_SelectValue" value="<%=rsProduct("p_SelectValue")%>" />
  <br /><br />
  <span id="SelectValue"><script type="text/javascript">getSelectValue(<%=SelectType%>,'<%=rsProduct("p_SelectValue")%>',<%=id%>,0,'');</script>
  </span></td>
</tr>

<tr class="classrow" style="display:none">
<td  class="classrow"><input type="hidden" name="actionkucun" size="5"  value="<%= actionkucun %>" class="form" onKeyUp="value=value.replace(/[^\d\.]/g,'');">
 
  <input type="hidden" name="p_weight" id="p_weight" size="5"  value="<%= p_weight %>" class="form" onKeyUp="value=value.replace(/[^\d\.]/g,'');"></td>
</tr>
	  
		  
		  
		  
		  
		  
		  
		  
		  
<%'自定义字段%>
<%call showColumn("product",3,id1)%>
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
			    <td ><%=chanpin%>缩略图：</td> 
			<td>
				   <input name="SmallImage"  type="text" id="weepic"  value="<%=rsProduct("SmallImage")%>" size="50"  /> <br/><br/>
			      <a class='sfile' href="Upload_All.Asp?add=one" title="支持上传格式(*.swf;*.gif;*.bmp)"><img src="Images/dtup.jpg"  border="0"/></a>
			      <div id="mysmaimg">

<%if rsproduct("Smallimage")<>"" then%>

<img src="../<%=rsProduct("SmallImage")%>" id="sc"/><br />
<a href="delslt.asp?sltdz=<%=rsproduct("Smallimage")%>" target="_blank" onClick="deltp()" id="lj">删除</a>
<%end if%>
</div>		        </td>
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
			  	K('input[name=btfy]').click(function(e) {
					editor.insertHtml('{选项卡标记}{标题:输入选项卡标题}');
				});
			prettyPrint();
 });
 </script>
          <tr > 
            <td ><%=chanpin%>说明：</td>
			<td>
			 	   
 <INPUT type=button class="bnt" name="btfy" value="插入选性卡">
 </br></br>
             <textarea name="Content" id="Content" style="width:700px;height:300px;visibility:hidden;" datatype="*" nullmsg="<%=chanpin%>内容不能为空"><%=rsProduct("content")%></textarea></td>
          </tr>
          <tr  >
            <td >批量上传图片：</td>
			<td>
		幻灯图片地址：	<input type="text" id="IcoImage" name="IcoImage" value="<%=rsProduct("icoimage")%>" style="width:600px;" /><br/><br/>
		<input type="button" id="J_selectImage" value="" style=" background:url(images/allimg.png) no-repeat; width:170px;height:34px;border:none" /><br/><br/>
        多图地址： <input type="text" id="ImagePath" name="ImagePath" value="<%=rsProduct("ImagePath")%>" style="width:600px;" /><br/><br/>
			
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
<a href="javascript:setContent('<%=img(i)%>')" ><img title="点击图片将图片插入到编辑器" src="<%=img(i)%>" border="0" /></a><br><label><input type="radio" name="IndexImageradio" onClick="setIndexImage('<%=img(i)%>')" value="<%=img(i)%>" <%if rsProduct("icoimage")=img(i) then response.write "checked" %>/>设为幻灯图</label> <a href="javascript:dropThisDiv('<%=ids%>','<%=img(i)%>')">删除</a></div>
 <%
 next
 end if
 %>    
 </div>


 
 </td>
          </tr>
          <tr  >
	
      <td  >手动分页：</td>
			<td><INPUT type=button class="bnt" name="appendHtml" value="插入分页符">
			 </td>
          </tr>
			
		  <tr>
	

        <input name="zdfy" type="hidden" id="zdfy"  value="<%=rsProduct("zdfy")%>" size="10" />
       
        <!--
          <tr > 
            <td   >已通过审核：</br></br> <input name="Passed" type="checkbox" id="Passed" value="1" <% if rsProduct("Passed")=1 then response.Write("checked") end if%>>
            是</td>
          </tr>
		  -->
          <tr >
            <td   >推荐<%=chanpin%>：</td>
		<td><input name="tj" type="checkbox" id="tj" value="1" <% if rsProduct("tj")=1 then response.Write("checked") end if%>></td>
          </tr>
          <tr > 
            <td   >首页新品显示：</td>
		<td>
              <input name="newproduct" type="checkbox" id="newproduct" value="1" <% if rsProduct("newproduct")=1 then response.Write("checked") end if%>>
            是</td>
          </tr>
		  	 <tr > 
            <td  >点击次数：</td>
			<td>
              <input  name="hits" type="text" class="input" value="<%=rsProduct("hits")%>" size="30"></td>
          </tr>
	 <tr >
		  <tr  >
      <td  >发布人：</td>
		<td>
        <input  name="user" type="text" class="input" size="30" value="<%=rsProduct("user")%>">      </td>
    </tr>
    <tr  >
      <td  >来源：</td>
		<td>
      <input  name="ly" type="text" id="ly" value="<%=rsProduct("ly")%>"></td>
    </tr>
          <tr > 
            <td   >录入时间：</td>
		<td> <input name="adddate" type="text" id="adddate" value="<%=rsProduct("adddate")%>"  onClick="WdatePicker({dateFmt:'yyyy/MM/dd HH:mm:ss'})" ></td>
          </tr>
          <tr > 
            <td   colspan="2"> 
              <div align="left">
                <input  type="submit" name="Submit" value="修改" class="bnt">
                </div></td></tr>
  </table>
</form>

<%

rsProduct.close
set rsProduct=nothing
conn.close
	set conn=nothing
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

<%

Sub setUsersPrice(pid)	
%>
<script language="javascript">

var uGradeArr = new Array();
<%
sql="select g.grouplevel,g.groupname,p.price from [hygroup] as g Left Join (Select price,grade_id from [product_price] where product_id = "&pid&") as p on p.grade_id=g.grouplevel where grouplevel <> 0 and lang='"&lang&"' order by g.grouplevel"
set rs_up = conn.execute(sql)
j = 0
do while not rs_up.eof
%>
	var tArr = new Array();
	tArr[0] = "<%=rs_up(0)%>";
	tArr[1] = "<%=rs_up(2)%>";
	tArr[2] = "<%=rs_up(1)%>";

	uGradeArr[<%=j%>] = tArr;
<%
	j = j+1
	rs_up.movenext
Loop
rs_up.close:Set rs_up = Nothing
%>
function setPrice()
{
	var s='';
	for(var i = 0; i < uGradeArr.length; i++)
	{
		if (i==5) s = s + "<br />&nbsp;&nbsp;";
		s = s + uGradeArr[i][2]+'价格'+'：';
		s = s + '<br /><br /><input type="text" name="u_price" id="u_price" size="6"  value="'+uGradeArr[i][1]+'"  "><br /><br />';
	}
	document.write(s);
}
</script>
<%
End Sub
Sub updateUsersPrice(pid,price_str,grade_str,lang)	'更新会员组价格
	conn.execute("delete from [product_price] where product_id="&pid&" ")
	pArr = Split(price_str,",")
	gArr = Split(grade_str,",")
	If ubound(pArr) = ubound(gArr) Then
		for j=0 to ubound(pArr)
			If Trim(gArr(j)) <> "" And Trim(pArr(j)) <> "" Then
				conn.execute("insert into [product_price] (product_id,grade_id,price,lang)values("&pid&","&gArr(j)&","&pArr(j)&",'"&lang&"')")
			End If
		Next
	End If
End Sub
%>
