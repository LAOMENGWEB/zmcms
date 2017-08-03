<!--#include file="../../Lib/base.asp"-->
<!--#include file="../../Lib/Page.asp"-->
<%

	
	dim datadb,datadb2
	dim i
	dim act:act=lcase(Cls.fget("act",0))
	select case act
		case "adddb":adddb
		case "editdb":editdb
		case "del":del
		case "delsome":delsome
	end select
	
	sub adddb()
		dim p_Title,p_IdList
		
		p_Title=Cls.enhtml(Cls.fpost("p_Title",0))
		p_IdList=Cls.enhtml(Cls.fpost("p_IdList",0))
		p_IdList=left(p_IdList,len(p_IdList)-1)		'去掉最后一个,
		p_Addtime=now()
		
		Dim data
		data=array(array("p_Title",p_Title,255,1),array("p_IdList",p_IdList,0,1),array("p_Addtime",p_Addtime,50,0))
		if Cls.db.dbnew("[Plug_ListProduct]",data,"")=1 then
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if	
		Cls.die
	end sub
	
	
	sub editdb()
		dim p_Title,p_IdList
		dim id:id=Cls.getint(Cls.fget("id",0),0)
		
		p_Title=Cls.enhtml(Cls.fpost("p_Title",0))
		p_IdList=Cls.enhtml(Cls.fpost("p_IdList",0))
		p_IdList=left(p_IdList,len(p_IdList)-1)		'去掉最后一个,
		
		Dim data
		data=array(array("p_Title",p_Title,255,1),array("p_IdList",p_IdList,0,1))
		if Cls.db.dbupdate("[Plug_ListProduct]","p_Id="&id,data)=1 then
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if	
		Cls.die	
	end sub
	
	
	sub del()
		dim id:id=Cls.getint(Cls.fget("id",0),0)
		if id>0 then
			Cls.db.dbdel "[Plug_ListProduct]","p_id="&id&""			
			Cls.db.dbdel "[sys_KeyWord]","k_plugName='ListProduct' and k_plugParam="&id&""	'删除关键字已指定的素材
		end if
		Cls.echo "1"
		Cls.die
	end sub
	
	sub delsome()
		dim id:id=Cls.enhtml(Cls.fget("id",0))
		dim idarr:idarr=split(id,",")
		if ubound(idarr)<0 then
			Cls.echo "0至少选择一条信息"
		else
			dim i
			for i=0 to ubound(idarr)
				if not isnumeric(idarr(i)) then
					Cls.echo "0参数："&id(i)&"不正确，请确认后再操作"
					exit sub
				else
					Cls.db.dbdel "[sys_KeyWord]","k_plugName='ListProduct' and k_plugParam="&id&""	'删除多图文素材关联的关键字
				end if
			next	
			Cls.db.dbdel "[Plug_ListProduct]","p_id in("&id&")"
			Cls.echo "1"
		end if
		Cls.die
	end sub
	
%>
<link href="../../Css/wx.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../../js/jquery.js"></script>
<script language="javascript" src="../../js/jquery.artDialog.js?skin=default"></script>
<script language="javascript" src="../../js/artDialog.iframeTools.js"></script>
<script language="javascript" src="../../js/base.js"></script>
<script language="javascript" src="../../js/admin.js"></script>
<script>
function checkadd(the)
{
	var ListProduct=document.getElementById('ListProduct');
	var ItemList
	ItemList="";
	for(var i=0;i<ListProduct.options.length;i++){
		 ItemList +=ListProduct.options[i].value + ",";
	}
	if($.trim(the.p_Title.value)==""){
		$.message({content:"标题不能为空"});
		the.p_Title.focus();
		return false;
	}
	if (ListProduct.options.length<1){
		$.message({content:"请选择图文素材！"});
		return false;
		
	}
	if (ListProduct.options.length>10){
		$.message({content:"图文素材只允许10篇以下！"});
		return false;
	}
	var url,data;
	url="?act=adddb";
	data="p_Title="+encodeURIComponent($.trim(the.p_Title.value));
	data+="&p_IdList="+encodeURIComponent($.trim(ItemList));
	
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
	var ListProduct=document.getElementById('ListProduct');
	var ItemList
	ItemList="";
	for(var i=0;i<ListProduct.options.length;i++){
		 ItemList +=ListProduct.options[i].value + ",";
	}
	if($.trim(the.p_Title.value)==""){
		$.message({content:"标题不能为空"});
		the.p_Title.focus();
		return false;
	}
	if (ListProduct.options.length<1){
		$.message({content:"请选择图文素材！"});
		return false;
		
	}
	if (ListProduct.options.length>10){
		$.message({content:"图文素材只允许10篇以下！"});
		return false;
	}
	var url,data;
	url="?act=editdb&id="+id+"";
	data="p_Title="+encodeURIComponent($.trim(the.p_Title.value));
	data+="&p_IdList="+encodeURIComponent($.trim(ItemList));
	
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
				content:'确定要删除？不可恢复！删除图文素材将会连带删除与之关联的关键字！',
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
	
	$(".delsome").click(function(){
		    var arrchk=$("input[name='id']:checked");
			var idarr="";
			$(arrchk).each(function(){
				if(idarr==""){idarr+=this.value}else{idarr+=","+this.value}                   
			}); 
			if(idarr=="")
			{
				$.message({content:"至少选择一条信息"});
			}
			else
			{
				var throughBox=$.dialog.through;
				throughBox({
				icon:'question',
				content:'确定要删除？不可恢复！删除图文素材将会连带删除与之关联的关键字！',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='?act=delsome&id='+idarr;
					$.ajax({
					type:"get",
					cache:false,
					url:url,
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
								$.message({type:"ok",content:"删除成功"});
								var idnum;
								idnum=idarr.split(",")
								for(i=0;i<=idnum.length;i++)
								{
									$("#list_"+idnum[i]).fadeOut('slow');
								}
								break;
							default:
								alert(_);
								break;
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			})
			}
		})
	
 })

function removeItem(){
   var DataList=document.getElementById('DataList');
   var ListProduct=document.getElementById('ListProduct');
   for(var i=0;i<DataList.options.length;i++)
   {
	   var tempOption=DataList.options[i];
	   if(tempOption.selected){
		   DataList.removeChild(tempOption);
		   ListProduct.appendChild(tempOption);
	   }    
   }
}

function addItem(){
   var DataList=document.getElementById('DataList');
   var ListProduct=document.getElementById('ListProduct');
   for(var i=0;i<ListProduct.options.length;i++)
   {
	   var tempOption=ListProduct.options[i];
	   if(tempOption.selected){
		   ListProduct.removeChild(tempOption);
		   DataList.appendChild(tempOption);
	   }    
   }
}
function getdata(t0,t1,t2){
	var id = t1;
	var plug= t0;
	if(!t0==""){
		$.ajax({
			   type: "post",
			   url:"<%=webroot%>plug/"+plug+"/Admin_Api.asp?act=getlist&cmstype="+plug+"&id="+id,
			   cache:false,
			   timeout:10000,
			   success: function(data)
			   {   
				   $("#"+t2).empty();
				   $("#"+t2).append(data);
			   }
		 });
	}
}
</script>
</head>
<body>

	<%
	if act="edit" then	
	dim id:id=Cls.getint(Cls.fget("id",0),0)
	datadb=Cls.db.dbload("","p_Title,p_IdList","[Plug_ListProduct]","p_id="&id,"")
	%>
    <script>	
	$(function(){
		getdata("ListProduct","","DataList");
		getdata("ListProduct","<%=datadb(1,0)%>","ListProduct");
	})
	</script>

    <form onSubmit="return checkedit(this,<%=id%>)">
		<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >编辑产品模块素材</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
  <div class="dmeun"><a href="admin.asp" target="main" style="color:#ffffff">产品素材管理</a></div>
  
 
 
 
 </td>
 

 </tr>
</table>	
	
	
			<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >图文标题：</td>
<td width="88%">
<input type="text" name="p_Title" size="50" value="<%=datadb(0,0)%>"/>
</td>
 </tr>	
 
 <tr > 
<td width="10%"   >数据集：</td>
<td width="88%">
  <select ondblclick="removeItem();" id="DataList" multiple="true" style="height:150px;width:300px">
     </select> ==>
     <select ondblclick="addItem();" id="ListProduct"  multiple="true" style="height:150px;width:300px">
     </select> 直接双击选择图文素材到右边，最多只支持10篇信息
</td>
 </tr>
 
    <tr > 
            <td colspan="2"> <input class="bnt" type="submit" name="send" value="保存"><input type="button" value="返回" onClick="location.href='javascript:history.go(-1)'" class="bnt"/>
            　            </td>
          </tr>	  
 </table>
	
	
	
    </form>

<%elseif act="add" then%>
    <script>	
	$(function(){
		getdata("ListProduct","","DataList");
	})
	</script>

    <form onSubmit="return checkadd(this)">
		<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >添加产品模块素材</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="admin.asp" target="main" style="color:#ffffff">产品素材管理</a></div>

 
 
 
 </td>
 

 </tr>
</table>
	
	
	

			<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >图文标题：</td>
<td width="88%">
<input type="text" name="p_Title" size="50" />
</td>
 </tr>	
 
 <tr > 
<td width="10%"   >数据集：</td>
<td width="88%">
  <select ondblclick="removeItem();" id="DataList" multiple="true" style="height:150px;width:300px">
     </select> ==>
     <select ondblclick="addItem();" id="ListProduct"  multiple="true" style="height:150px;width:300px">
     </select> 直接双击选择图文素材到右边，最多只支持10篇信息
</td>
 </tr>
 
 
     <tr > 
            <td colspan="2"> <input class="bnt" type="submit" name="send" value="保存">
			<input type="button" value="返回" onClick="location.href='javascript:history.go(-1)'" class="bnt"/>
            　            </td>
          </tr>	  
 
 </table>
    </form>

<%else%>

  <form>
  <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >产品素材管理</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="?act=add" target="main" style="color:#ffffff">添加产品素材</a></div>
 <div class="dmeun"><a href="javascript:;" class="delsome" target="main" style="color:#ffffff">批量删除</a></div>
 </td>
 

 </tr>
</table>

   <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
      <tr class="wz">
        <td><input type="checkbox" name="chkall" style="border:0;" onClick="checkall(this.form)" title="全选/取消" /></td>
        <td>ID</td>
        <td>标题</td>
        <td>数据集</td>
        <td>添加时间</td>
        <td>管理</td>
      </tr>
        <%
		Set Page = new Page_List
		Page.Con = Cls.db.conn		
		dim SqlStr:SqlStr="Select p_Id,p_Title,p_idList,p_Addtime from [Plug_ListProduct] order by p_Id desc"
		Page.Sql = SqlStr
		Page.PageSize = 20
		Page.Style = 5
		Set Rs = Page.Rs
		
		If not Rs.Bof or Rs.Eof Then
			For i = 1 To Page.PageSize%>
      <tr id="list_<%=rs("p_id")%>">
        <td align="center"><input name="id" type="checkbox" style="border:0;" value="<%=rs("p_id")%>" /></td>
        <td align="center" ><%=rs("p_id")%></td>
        <td align="center"><%=rs("p_Title")%></td>
        <td align="center" ><%=rs("p_idList")%></td>
        <td align="center" ><%=cls.format_time(rs("p_Addtime"),1)%></td>
        <td align="center" ><a href="?act=edit&id=<%=rs("p_id")%>">编辑</a>　<a href="javascript:;" class="del" rel="<%=rs("p_id")%>">删除</a></td>
      </tr>
      <%
				Rs.MoveNext
			Next
		End If
		
		Call Page.ShowPage
		
		Rs.Close
		Set Rs = Nothing
		Set Page = Nothing%>
    </table>
  </form>
</div>
<%end if%>
</body>
</html>