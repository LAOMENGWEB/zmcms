<!--#include file="../../weixin/Lib/base.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../../weixin/Lib/Page.asp"-->
<%
if Instr(session("AdminPurview"),"|9999,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""../right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%

	dim datadb,datadb2
	dim i
	dim act:act=lcase(Cls.fget("act",0))
	select case act
		case "adddb":adddb
		case "editdb":editdb
		case "del":del
	end select
	
	sub adddb()
		Dim c_ClassName:c_ClassName=Cls.enhtml(Cls.fpost("c_ClassName",0))		
		Dim data:data=array(array("c_ClassName",c_ClassName,50,1))
		if Cls.db.dbnew("[sys_KeyWord_class]",data,"c_ClassName='"&c_ClassName&"'")=1 then
			Cls.echo "1"
		else
			Cls.echo "0该分类已存在"
		end if	
		Cls.die
	end sub
	
	
	sub editdb()
		Dim id:id=Cls.getint(Cls.fget("id",0),0)		
		Dim c_ClassName:c_ClassName=Cls.enhtml(Cls.fpost("c_ClassName",0))		
		Dim data:data=array(array("c_ClassName",c_ClassName,50,1))
		if Cls.db.dbupdate("[sys_KeyWord_class]","c_id="&id,data)=1 then
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if	
		Cls.die	
	end sub
	
	
	sub del()
		dim id:id=Cls.getint(Cls.fget("id",0),0)
		if id>0 then
			Cls.db.dbdel "[sys_KeyWord_class]","c_id="&id&""
		end if
		Cls.echo "1"
		Cls.die
	end sub
	
%>
<link href="../../weixin/css/wx.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../../weixin/js/jquery.js"></script>
<script language="javascript" src="../../weixin/js/jquery.artDialog.js?skin=default"></script>
<script language="javascript" src="../../weixin/js/artDialog.iframeTools.js"></script>
<script language="javascript" src="../../weixin/js/base.js"></script>
<script language="javascript" src="../../weixin/js/admin.js"></script>
<script>
function checkadd(the)
{
	if($.trim(the.c_ClassName.value)=="")
	{
		$.message({content:"分类名称不能为空"});
		the.c_ClassName.focus();
		return false;
	}
	var url,data;
	url="?act=adddb";
	data="c_ClassName="+encodeURIComponent($.trim(the.c_ClassName.value));
	
	$.ajax({
	type:"post",
	cache:false,
	url:url,
	data:data,
	error:function(){$.message({type:"error",content:"服务器错误，操作失败！"});},
	success:function(_)
	{
		var act=_.substring(0,1);
		var info=_.substring(1);
		switch(act)
		{
			case "0":
				$.message({type:"error",content:info});
				break;
			case "1":
				$.message({type:"ok",content:"保存成功"});
				setTimeout("location.href='?'",1000);
				break;
			default:
				alert(_)
				break;
		}
	}
	});
	return false

}

function checkedit(the,id)
{
	if($.trim(the.c_ClassName.value)=="")
	{
		$.message({content:"分类名称不能为空"});
		the.c_ClassName.focus();
		return false;
	}
	var url,data;
	url="?act=editdb&id="+id+"";
	data="c_ClassName="+encodeURIComponent($.trim(the.c_ClassName.value));
	
	$.ajax({
	type:"post",
	cache:false,
	url:url,
	data:data,
	error:function(){$.message({type:"error",content:"服务器错误，操作失败！"});},
	success:function(_)
	{
		var act=_.substring(0,1);
		var info=_.substring(1);
		switch(act)
		{
			case "0":
				$.message({type:"error",content:info});
				break;
			case "1":
				$.message({type:"ok",content:"保存成功"});
				setTimeout("location.href='?'",1000);
				break;
			default:
				alert(_)
				break;
		}
	}
	});
	return false

}
$(function(){
	$(".del").click(function(){
		var id=this.getAttribute('rel');
		var throughBox=$.dialog.through;
			throughBox({
				icon:'question',
				content:'确定要删除？不可恢复！',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='?act=del&id='+id;
					$.ajax({
					type:"get",
					cache:false,
					url:url,
					error:function(){$.message({type:"error",content:"服务器错误，操作失败！"});},
					success:function(_)
					{
						if(_.substring(0,1)==0)
						{
							$.message({type:"error",content:_.substring(1)});
						}
						else
						{
							$.message({type:"ok",content:"删除成功"});
							$("#list_"+id).fadeOut('slow');
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			})
		
		})
 })
</script>
</head>
<body>

	<%
	if act="edit" then	
	dim id:id=Cls.getint(Cls.fget("id",0),0)
	datadb=Cls.db.dbload("","c_ClassName","[sys_KeyWord_class]","c_id="&id,"")
	%>

    <form onSubmit="return checkedit(this,<%=id%>)">
	<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >修改关键字分类</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 
 <div class="dmeun"><a href="?act=add" target="main" style="color:#ffffff">添加关键字分类</a></div>
 
 </td>
 

 </tr>
</table>		   
	<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >分类名称：</td>
<td width="88%">
 <input type="text" name="c_ClassName" size="50" value="<%=datadb(0,0)%>"/>
</td>
 </tr>	
 <tr > 
            <td colspan="2">  <input type="submit" name="send" value="保存" class="bnt"/>
            　            </td>
          </tr>	  
 </table>
	
    </form>

<%elseif act="add" then%>

    <form onSubmit="return checkadd(this)">
	
	
	
	<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >添加关键字分类</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 
 <div class="dmeun"><a href="keyword_class.asp" target="main" style="color:#ffffff">关键字分类管理</a></div>
 
 </td>
 

 </tr>
</table>		   
	<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >分类名称：</td>
<td width="88%">
 <input type="text" name="c_ClassName" size="50" />
</td>
 </tr>	
 <tr > 
            <td colspan="2">  <input type="submit" name="send" value="保存" class="bnt"/>
            　            </td>
          </tr>	  
 </table>
	
	
	
	
    </form>

<%else%>

  <form>
  
  <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >关键字分类管理</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 
 <div class="dmeun"><a href="?act=add" target="main" style="color:#ffffff">添加关键字分类</a></div>
 <div class="dmeun"><a href="keyword.asp" target="main" style="color:#ffffff">关键字管理</a></div>
 
 </td>
 

 </tr>
</table>	
  
  
  
  
  
  
   <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
      <tr class="wz">
        <td><input type="checkbox" name="chkall" style="border:0;" onClick="checkall(this.form)" title="全选/取消" /></td>
        <td>ID</th>
        <td>分类名称</td>
        <td>管理</td>
      </tr>        
      	<%datadb=Cls.db.dbload("","c_id,c_ClassName","[sys_KeyWord_class]","","c_id asc")
		if ubound(datadb)>=0 then
		for i=0 to ubound(datadb,2)%>
      <tr id="list_<%=datadb(0,i)%>">
        <td align="center"><input name="id" type="checkbox" style="border:0;" value="<%=datadb(0,i)%>" /></td>
        <td align="center" ><%=datadb(0,i)%></td>
        <td align="center" ><%=datadb(1,i)%></td>
        <td align="center" ><a href="?act=edit&id=<%=datadb(0,i)%>">编辑</a>　<a href="javascript:;" class="del" rel="<%=datadb(0,i)%>">删除</a></td>
      </tr>
      <%next
	  end if%>
    </table>
  </form>
</div>
<%end if%>
</body>
</html>
