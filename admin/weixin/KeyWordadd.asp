<!--#include file="../../weixin/Lib/base.asp"-->
<!--#include file="admin.asp"-->

<%
if Instr(session("AdminPurview"),"|9999,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""../right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>


<%

		
		''查询数据库，t0为top,t1为字段,t2为表名,t3为where条件，t4为排序
	function dbload(byval t0,byval t1,byval t2,byval t3,byval t4)
		if len(t0)>0 then t0=" top "&t0
		if len(t1)=0 then t1=" * "
		if len(t3)>0 then t3=" where "&t3
		if left(lcase(t4),3)="rnd" then
			if datatype then
				randomize
				dim orderid
				if len(t4)>4 then orderid=right(t4,len(t4)-4)
				if len(orderid)=0 then orderid="id"
				t4="rnd(-("&orderid&" +"&rnd()&"))"
			else
				t4="newid()"
			end if
		end if
		if len(t4)>0 then t4=" order by "&t4
		if t1<>" * " then
			dim str:str=""
			dim arr:arr=split(t1,",")
			dim i
			for i=0 to ubound(arr)
				if i>0 then str=str&","
				if instr(arr(i),"(")>0 or instr(arr(i),")")>0 or instr(arr(i)," ")>0 or instr(arr(i),".")>0 then
					str=str&arr(i)
				else
					str=str&"["&arr(i)&"]"
				end if
			next
			t1=str
		end if
		set rs=exedb("select "&t0&" "&t1&" from "&t2&" "&t3&" "&t4)
		if rs.eof then
			dbload=array()
		else
			dbload=rs.getrows()
		end if
		rs.close
		set rs=nothing
	end function

If request.Form("submit") = "确定" Then
	   k_classId=request.Form("k_classId")
		k_keyWord=request.Form("k_keyWord")
		k_keyWord_beizhu=request.Form("k_keyWord_beizhu")
		k_keyType=request.Form("k_keyType")
		k_type=request.Form("k_type")
		k_text=request.Form("k_text")
		k_plugName=request.Form("k_plugName")
		k_plugParam=request.Form("k_plugParam")
		k_num=request.Form("k_num")
		k_addtime=now()

Set rs = server.CreateObject("adodb.recordset")
rs.open "Select * From Plug_PicMsg ",conn,1,3
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
<link href="../../weixin/Css/wx.css" rel="stylesheet" type="text/css" />

<script language="JavaScript" src="../../../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../../../js/Validform_v5.3.2_min.js"></script>



 <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})

function getdata(t0,t1){
	var id = t1;
	var plug= t0;
	if(!t0==""){
		$.ajax({
			   type: "post",
			   url:"<%=webroot%>plug/"+plug+"/Admin_Api.asp?id="+id,
			   cache:false,
			   timeout:10000,
			   success: function(data)
			   {   
				   $("#k_plugParam").empty();
				   $("#k_plugParam").append(data);
			   }
		 });
	}
}
</script>


</head>

<body>
   <form method="post" name="myform" class="checkform">
	

	
	
	<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >添加关键词</td>
</tr>
</table>
	
	
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 
 <div class="dmeun"><a href="keyword.asp" target="main" style="color:#ffffff">关键字管理</a></div>
 
 </td>
 

 </tr>
</table>	
	
	
		   
<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">

<tr > 
<td width="10%"   >接收类型：</td>
<td width="88%">
   <select name="k_msgtype" id="k_msgtype" onChange="setBeizhu(this.value)">
          <option value="text" selected>文本</option>
          <!--<option value="voice">语音</option>
          <option value="image">图片</option>
          <option value="video">视频</option>-->
        </select>
         
</td>
 </tr>	
 

 
 <tr > 
<td width="12%"   >关键字优先级别：</td>
<td width="88%">
 <input type="text" name="k_num" size="4" value="0"/>
        关键字匹配的优先级别，数字越小的级别越高。
</td>
 </tr>	
 
 
  <tr > 
<td width="12%"   >关键字：</td>
<td width="88%">
  <input type="text" name="k_keyWord" size="50" value="<%=keyword%>"/>
        设定关键字，当用户发送此关键字时将会自动回复消息
</td>
 </tr>	
 
   <tr > 
<td width="12%"   >关键字备注：</td>
<td width="88%">
  <input type="text" id="beizhu" name="k_keyWord_beizhu" size="50"/>
        设定关键字备注
</td>
 </tr>	
 
    <tr > 
<td width="12%"   >所属分类：</td>
<td width="88%">
 <select name="k_classId" id="k_classId">
 
      	<%datadb=dbload("","c_id,c_ClassName","[sys_KeyWord_class]","c_id<>1","c_id asc")
		if ubound(datadb)>=0 then
		for i=0 to ubound(datadb,2)%>
          <option value="<%=datadb(0,i)%>"><%=datadb(1,i)%></option>
      <%next
	  end if%>
        </select>  本项只作为后台备注分类，不影响正常使用
</td>
 </tr>	
 
 
    <tr > 
<td width="12%"   >匹配类型：</td>
<td width="88%">
 <select name="k_keyType" id="k_keyType">
          <option value="完全匹配">完全匹配</option>
          <option value="左匹配">左匹配</option>
          <option value="右匹配">右匹配</option>
          <option value="模糊匹配">模糊匹配</option>
        </select>
</td>
 </tr>	
 
 
 
     <tr > 
<td width="12%"   >回复类型：</td>
<td width="88%">
 <select name="k_type" id="k_type" onChange="setclose(this.value)">
          <option value="文本">文本</option>
          <option value="插件">插件</option>
        </select>
        【文本】为直接回复文本信息，【插件】为调用系统中的插件进行处理关键字
</td>
 </tr>	
 
 
 
      <tr > 
<td width="12%"   >指定调用插件：</td>
<td width="88%">
 <select name="k_plugName" id="k_plugName"  onChange="getdata(this.value,0)">
          <option value="">请选择调用插件</option>
		  
		  
		  
		  
		  
		  
      	<%datadb=dbload("","p_Name,p_Title","[Plug]","(p_Type=1 or p_Type=2) and p_IsLock='开启'","p_num asc")
		if ubound(datadb)>=0 then
		for i=0 to ubound(datadb,2)%>
          <option value="<%=datadb(0,i)%>"><%=datadb(1,i)%></option>
      <%next
	  end if%>
          </select> 请选择数据：
        <select name="k_plugParam" id="k_plugParam">
	  	<option value="">请选择数据</option>
        </select>
		
		<font color="#ff0000">当回复类型为插件时有效</font>
</td>
 </tr>	
 
 
      <tr > 
<td width="12%"   >回复内容：</td>
<td width="88%">
<textarea name="k_text" cols="60" rows="20"></textarea> <font color="#ff0000">当回复类型为文本时有效</font></td>
 </tr>	
 
 
 
 
 
 
 
 
 
 
 	   
	   <tr > 
            <td colspan="2"> <input class="bnt" type="submit" name="send" value="保存">
            　            </td>
          </tr>	   
</table>	
	
	
	
	
	
	
	
	
	
	

    </form>

	
</body>
</html>
