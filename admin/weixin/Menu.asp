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

	dim datadb,datadb2,id,errcode,diyMenu
	dim i,n
	dim ClassName,Fid,Types,Url,Keyword
	
	dim act:act=lcase(Cls.fget("act",0))
	select case act
		case "adddb":adddb
		case "editdb":editdb
		case "del":del
		case "postmenu":postmenu
		case "delmenu":delmenu
		case "getmenu":getmenu
	end select
	
	sub adddb()
		ClassName=Cls.enhtml(Cls.fpost("ClassName",0))
		Fid=Cls.getint(Cls.fpost("Fid",0),0)
		Types=Cls.enhtml(Cls.fpost("Types",0))
		Url=Cls.enhtml(Cls.fpost("Url",0))
		Keyword=Cls.enhtml(Cls.fpost("Keyword",0))
		
		'判断字符长度
		if Fid=0 then
			If Cls.strlen(ClassName)>5 then Cls.echo "0一级栏目名称不能超过5个中文字符":exit sub
		else
			If Cls.strlen(ClassName)>12 then Cls.echo "0二级栏目名称不能超过12个中文字符":exit sub
		end if
		
		'判断栏目个数
		If Fid=0 then
			if Cls.db.dbcount("[sys_Menu]","Fid=0")>2 then Cls.echo "0一级栏目不能超过3个":exit sub
		Else
			if Cls.db.dbcount("[sys_Menu]","Fid="&Fid)>4 then Cls.echo "0二级栏目不能超过5个":exit sub
		end if
		
		Dim data
		data=array(array("ClassName",ClassName,50,1),array("Fid",Fid,0,0),array("Types",Types,50,1),array("Url",Url,0,1),array("Keyword",Keyword,0,1))
		if Cls.db.dbnew("[sys_Menu]",data,"")=1 then
			Call inputMenu()
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if
		cls.die
	end sub
	
	
	sub editdb()		
		id=Cls.getint(Cls.fget("id",0),0)		
		ClassName=Cls.enhtml(Cls.fpost("ClassName",0))
		Fid=Cls.getint(Cls.fpost("Fid",0),0)
		Types=Cls.enhtml(Cls.fpost("Types",0))
		Url=Cls.enhtml(Cls.fpost("Url",0))
		Keyword=Cls.enhtml(Cls.fpost("Keyword",0))
		
		'判断字符长度
		if Fid=0 then
			If Cls.strlen(ClassName)>5 then Cls.echo "0一级栏目名称不能超过5个中文字符":exit sub
		else
			If Cls.strlen(ClassName)>12 then Cls.echo "0二级栏目名称不能超过12个中文字符":exit sub
		end if
		
		dim OFid:OFid=Cls.db.dbloadone("Fid","[sys_Menu]","id="&id)
		If Fid<>OFid then
			If Fid=0 then
				if Cls.db.dbcount("[sys_Menu]","Fid=0")>2 then Cls.echo "0一级栏目不能超过3个":exit sub
			Else
				if Cls.db.dbcount("[sys_Menu]","Fid="&Fid)>4 then Cls.echo "0二级栏目不能超过5个":exit sub
			end if
		end if
		
		Dim data
		data=array(array("ClassName",ClassName,50,1),array("Fid",Fid,0,0),array("Types",Types,50,1),array("Url",Url,0,1),array("Keyword",Keyword,0,1))
		if Cls.db.dbupdate("[sys_Menu]","id="&id,data)=1 then
			Call inputMenu()
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if	
		cls.die
	end sub
		
	sub del()
		id=Cls.getint(Cls.fget("id",0),0)
		if id>0 then
			Cls.db.dbdel "[sys_Menu]","id="&id&""
			Cls.db.dbdel "[sys_Menu]","Fid="&id&""
			Call inputMenu()
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if
		cls.die
	end sub
	
	'获取远程自定义菜单
	sub getmenu()
		Access_token=GetToken()
		if Cls.strlen(Access_token)=0 then Cls.echo "0接口出错，无法获取Access_token":Cls.die:exit sub
		strJson=GetURL(SendMenuUrl&"get?access_token="&Access_token)
		if InStr(strJson,"errcode")>0 then Cls.echo "0远程菜单不存在":Cls.die:exit sub
		
		'转换成标准菜单格式
		strJson=left(strJson,len(strJson)-1)
		strJson=Mid(strJson,9)
		strJson=replace(strJson,",""sub_button"":[]","")
		'将标准菜单存入diyMenu字段
		datadb=array(array("diyMenu",strJson,0,1))
		if Cls.db.dbupdate("[sys_Config]","id=1",datadb)=1 then
			Dim jsonObject
			Dim buttonnum, i, n
			dim ClassName,Fid,Types,Keyword,Url,ClassName_sub
			Set jsonObject = parseJSON(strJson)
			buttonnum = jsonObject.button.length
			'清空菜单数据库
			Cls.db.dbdel "[sys_Menu]",""
			'入库一级栏目
			for i = 0 to buttonnum - 1
				ClassName=jsonObject.button.Get(i).name
				Fid=0
				If isEmpty(scriptCtrl.eval("result.button["& i &"].sub_button")) Then
					Types=jsonObject.button.Get(i).type
					select case jsonObject.button.Get(i).type
					case "click"
						Keyword=jsonObject.button.Get(i).key
					case "view"
						Url=jsonObject.button.Get(i).url
					End Select
				End If
				call addmenudb(ClassName,Fid,Types,Keyword,Url)
				ClassName="":Types="":Keyword="":Url=""
			next
			'入库二级栏目
			for i = 0 to buttonnum - 1
				ClassName=jsonObject.button.Get(i).name
				If not isEmpty(scriptCtrl.eval("result.button["& i &"].sub_button")) Then
					for n = 0 to jsonObject.button.Get(i).sub_button.length - 1
						Fid=Cls.db.dbload("","id","[sys_Menu]","ClassName='"&ClassName&"' and Fid=0","")(0,0)
						ClassName_sub=jsonObject.button.Get(i).sub_button.Get(n).name
						Types=jsonObject.button.Get(i).sub_button.Get(n).type
						select case jsonObject.button.Get(i).sub_button.Get(n).type
						case "click"
							Keyword=jsonObject.button.Get(i).sub_button.Get(n).key
						case "view"
							Url=jsonObject.button.Get(i).sub_button.Get(n).url
						End Select
						call addmenudb(ClassName_sub,Fid,Types,Keyword,Url)
						ClassName_sub="":Types="":Keyword="":Url=""
					Next
				End If
			next
			
			Set jsonObject = Nothing 
			If IsObject(scriptCtrl) Then Set scriptCtrl = Nothing
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if	
		cls.die
	end sub
	
	'清除微信平台上的自定义菜单
	sub delmenu()
		Access_token=GetToken()
		if Cls.strlen(Access_token)=0 then Cls.echo "0接口出错，无法获取Access_token":Cls.die:exit sub
		strJson=GetURL(SendMenuUrl&"delete?access_token="&Access_token)
		Call InitScriptControl:Set objTest = getJSONObject(strJson)
		errcode=objTest.errcode
		If errcode="0" then
			Cls.echo "1"
		else
			Cls.echo "0意外出错，自定义菜单删除不成功，代码："&errcode
		end if
		cls.die
	end sub
		
	'发布自定义菜单
	sub postmenu()
		dim Menutype
		Access_token=GetToken()
		if Cls.strlen(Access_token)=0 then Cls.echo "0接口出错，无法获取Access_token":Cls.die:exit sub
		datadb=Cls.db.dbload("","diyMenu","[sys_Config]","id=1","")
			strJson=PostURL(SendMenuUrl&"create?access_token="&Access_token,datadb(0,0))
			Call InitScriptControl:Set objTest = getJSONObject(strJson)
			errcode=objTest.errcode
			If errcode="0" then	
				Cls.echo "1"
			else
				Cls.echo "0意外出错，自定义菜单添加不成功，代码："&errcode
			end if
		cls.die
	end sub
	
	'提前将栏目菜单转换成微信接口格式Jason
	sub inputMenu()
		diyMenu=""
		datadb=Cls.db.dbload("","id,ClassName,Types,Keyword,Url","[sys_Menu]","fid=0","id asc")
		if ubound(datadb)>0 then
			diyMenu="{""button"":["
				for i=0 to ubound(datadb,2)
					if Cls.db.dbcount("[sys_Menu]","Fid="&datadb(0,i))=0 then	'无子菜单
						if datadb(2,i)="view" then
							Menutype="""url"":"""&datadb(4,i)&""""
						else
							Menutype="""key"":"""&datadb(3,i)&""""
						end if
						diyMenu=diyMenu&"{""type"":"""&datadb(2,i)&""",""name"":"""&datadb(1,i)&""","&Menutype&"},"		
					else
						diyMenu=diyMenu&"{""name"":"""&datadb(1,i)&""",""sub_button"":["
						datadb2=Cls.db.dbload("","id,ClassName,Types,Keyword,Url","[sys_Menu]","fid="&datadb(0,i),"id asc")
						if ubound(datadb2)>0 then
							for n=0 to ubound(datadb2,2)						
								if datadb2(2,n)="view" then
									Menutype="""url"":"""&datadb2(4,n)&""""
								else
									Menutype="""key"":"""&datadb2(3,n)&""""
								end if
								diyMenu=diyMenu&"{""type"":"""&datadb2(2,n)&""",""name"":"""&datadb2(1,n)&""","&Menutype&"},"		
							next
						end if					
						diyMenu=left(diyMenu,len(diyMenu)-1)
						diyMenu=diyMenu&"]},"
					end if	
				next
			diyMenu=left(diyMenu,len(diyMenu)-1)
			diyMenu=diyMenu&"]}"
		else
			diyMenu=""
		end if
		Call Cls.db.dbupdate("[sys_Config]","id=1",array(array("diyMenu",diyMenu,0,1)))
	end sub
	
	Dim scriptCtrl
	Function parseJSON(str)
		If Not IsObject(scriptCtrl) Then
		   Set scriptCtrl = Server.CreateObject("MSScriptControl.ScriptControl")
		   scriptCtrl.Language = "JavaScript"
		   scriptCtrl.AddCode "function ActiveXObject() {}" ' 覆盖 ActiveXObject
		   scriptCtrl.AddCode "function GetObject() {}" ' 覆盖 ActiveXObject
		   scriptCtrl.AddCode "Array.prototype.get = function(x) { return this[x]; }; var result = null;"
		End If
		On Error Resume Next
		scriptCtrl.ExecuteStatement "var result = " & str & ";"
		Set parseJSON = scriptCtrl.CodeObject.result
		If Err Then
		   Err.Clear
		   Set parseJSON = Nothing
		End If
	End Function
	sub addmenudb(ClassName,Fid,Types,Keyword,Url)
		Dim data:data=array(array("ClassName",ClassName,50,1),array("Fid",Fid,0,0),array("Types",Types,50,1),array("Url",Url,0,1),array("Keyword",Keyword,0,1))
		if Cls.db.dbnew("[sys_Menu]",data,"")=1 then
		end if
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
	if($.trim(the.ClassName.value)=="")
	{
		$.message({content:"栏目名称不能为空"});
		the.ClassName.focus();
		return false;
	}
	if($('#Types option:selected').val()=="view")
	{
		if($.trim(the.Url.value)=="")
		{
			$.message({content:"链接地址不能为空"});
			the.Url.focus();
			return false;
		}
	}else{
		if($.trim(the.Keyword.value)=="")
		{
			$.message({content:"关键字不能为空"});
			the.Keyword.focus();
			return false;
		}
	}
	var url,data;
	url="menu.asp?act=adddb";
	data="ClassName="+encodeURIComponent($.trim(the.ClassName.value));
	data+="&Fid="+encodeURIComponent($.trim(the.Fid.value));
	data+="&Types="+encodeURIComponent($.trim(the.Types.value));	
	data+="&Url="+encodeURIComponent($.trim(the.Url.value));
	data+="&Keyword="+encodeURIComponent($.trim(the.Keyword.value));
	
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
	if($.trim(the.ClassName.value)=="")
	{
		$.message({content:"栏目名称不能为空"});
		the.ClassName.focus();
		return false;
	}
	if($('#Types option:selected').val()=="view")
	{
		if($.trim(the.Url.value)=="")
		{
			$.message({content:"链接地址不能为空"});
			the.Url.focus();
			return false;
		}
	}else{
		if($.trim(the.Keyword.value)=="")
		{
			$.message({content:"关键字不能为空"});
			the.Keyword.focus();
			return false;
		}
	}
	var url,data;
	url="menu.asp?act=editdb&id="+id+"";
	data="ClassName="+encodeURIComponent($.trim(the.ClassName.value));
	data+="&Fid="+encodeURIComponent($.trim(the.Fid.value));
	data+="&Types="+encodeURIComponent($.trim(the.Types.value));	
	data+="&Url="+encodeURIComponent($.trim(the.Url.value));
	data+="&Keyword="+encodeURIComponent($.trim(the.Keyword.value));
	
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
	<%dim appDb:appDb=Cls.db.dbload("","AppId,Appsecret","[sys_Config]","id=1","")
	if Cls.strlen(appDb(0,0))=0 or Cls.strlen(appDb(1,0))=0 then%>
	 var AppBox=$.dialog.through;
	  AppBox({
		  content:'请先填写 AppID 和 AppSecret ！',
		  lock:false,
		  opacity:'0.5',
		  ok:function(){
				setTimeout("location.href='set2.asp'",0);		
			  }
	  })
	<%end if%>
	$(".del").click(function(){
		var id=this.getAttribute('rel');
		var throughBox=$.dialog.through;
			throughBox({
				icon:'question',
				content:'确定要删除？不可恢复！',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='menu.asp?act=del&id='+id;
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
	$(".getmenu").click(function(){
		var throughBox=$.dialog.through;
			throughBox({
				icon:'question',
				content:'确定要获取远程菜单？将覆盖本地菜单数据！',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='menu.asp?act=getmenu';
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
							$.message({type:"ok",content:"远程自定义菜单获取成功"});
							setTimeout("location.href='?'",1500);
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			})
		
		})
	$(".delmenu").click(function(){
		var throughBox=$.dialog.through;
			throughBox({
				icon:'question',
				content:'确定要清除微信平台上的自定义菜单？本地菜单不受影响！',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='menu.asp?act=delmenu';
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
							$.message({type:"ok",content:"远程清除自定义菜单删除成功"});
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			})
		
		})
	$(".postmenu").click(function(){
		var throughBox=$.dialog.through;
			throughBox({
				icon:'question',
				content:'确定栏目已经设置好并直接发布到微信平台？',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='menu.asp?act=postmenu';
					$.ajax({
					type:"get",
					cache:false,
					url:url,
					beforeSend:loading(),
					error:function(){closeDialog();$.message({type:"error",content:"服务器错误，操作失败！"});},
					success:function(_)
					{
						closeDialog();
						if(_.substring(0,1)==0)
						{
							$.message({type:"error",content:_.substring(1)});
						}
						else
						{
							$.message({type:"ok",content:"自定义菜单已发布成功，需要24小时后看到效果，或者重新关注！"});
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			})
		
		})
 })
function loading(){
	var AppBox=$.dialog.through;
	  AppBox({
		  content:'<div align="center"><img src="../images/loading.gif"></div> <br>正在发布菜单，网络不给力，请稍等...',
		  title:false,
		  lock:false
	  })
}
function closeDialog(){
	var list = art.dialog.list;
	for (var i in list) {
		list[i].close();
	};
}
function setclose(t0)
{
	if(t0=="view"){$("#keywords").css("display","none");$("#links").css("display","block");}else{$("#keywords").css("display","block");$("#links").css("display","none");}
}
  </script>
</head>
<body>

	<%
	if act="edit" then	
	id=Cls.getint(Cls.fget("id",0),0)
	dim datadb3:datadb3=Cls.db.dbload("","ClassName,Fid,Types,Url,Keyword","[sys_Menu]","id="&id,"")
	%>
    <script>	
	$(function(){
		$("#<%if datadb3(2,0)="view" then%>keywords<%else%>links<%end if%>").css("display","none");
	})
	</script>

    <form onSubmit="return checkedit(this,<%=id%>)">
	
	<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >修改自定义菜单</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 <div class="dmeun"><a href="?act=add" target="main" style="color:#ffffff">添加自定义菜单</a></div>
 </td>
 

 </tr>
</table>	
<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >栏目名称：</td>
<td width="88%">
<input type="text" name="ClassName" size="50" value="<%=datadb3(0,0)%>"/>
</td>
 </tr>	
 
 <tr > 
<td width="12%"   >所属栏目：</td>
<td width="88%">
    <select name="Fid">
          <option value="0">作为一级栏目</option>
          <%datadb=Cls.db.dbload("","id,ClassName","[sys_Menu]","fid=0","id asc")
			if ubound(datadb)>0 then
				for i=0 to ubound(datadb,2)%>
          <option value="<%=datadb(0,i)%>" <%if datadb3(1,0)=datadb(0,i) then Cls.echo " selected"%>><%=datadb(1,i)%></option>
          <%next
		  end if%>
        </select>
</td>
 </tr>	
 
 <tr > 
<td width="10%"   >链接类型：</td>
<td width="88%">
  <select name="Types" id="Types" >
          <option value="view" <%if datadb3(2,0)="view" then Cls.echo " selected"%>>外部链接</option>
          <option value="click" <%if datadb3(2,0)="click" then Cls.echo " selected"%>>关键字</option>
        </select>
</td>
 </tr>	
 
 
  <tr> 
<td width="10%"   >链接地址：</td>
<td width="88%">
<input type="text" name="Url" size="50" value="<%=datadb3(3,0)%>" /> 请直接输入链接地址   <font color="#ff0000">当栏目类型为外部链接时有效</font>
</td>
 </tr>	
 
 
 
   <tr> 
<td width="10%"   >调用关键字：</td>
<td width="88%">
 <input type="text" name="Keyword" size="50"  value="<%=datadb3(4,0)%>" /> 请输入系统中的关键字 <font color="#ff0000">当栏目类型为关键字时有效</font>
</td>
 </tr>	
 
 
 
 
 
 
 
 
 
 	   
	   <tr > 
            <td colspan="2">   <input type="submit" name="send" value="保存" class="bnt"/>
            　            </td>
          </tr>	   
</table>		
	
	
	
	
	
	
	
	
	
	
	
	
	
    
   
    
   
    </form>
  </dl>
</div>
<%elseif act="add" then%>

    <form onSubmit="return checkadd(this)">
	
		<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >增加自定义菜单</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 <div class="dmeun"><a href="menu.asp" target="main" style="color:#ffffff">自定义菜单管理</a></div>
 </td>
 

 </tr>
</table>
	
	
	
	
	<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 
<td width="10%"   >栏目名称:</td>
<td width="88%">
<input type="text" name="ClassName" size="50" />
</td>
 </tr>	
 
 <tr > 
<td width="12%"   >所属栏目：</td>
<td width="88%">
   <select name="Fid">
          <option value="0">作为一级栏目</option>
          <%datadb=Cls.db.dbload("","id,ClassName","[sys_Menu]","fid=0","id asc")
			if ubound(datadb)>0 then
			for i=0 to ubound(datadb,2)%>
          <option value="<%=datadb(0,i)%>"><%=datadb(1,i)%></option>
          <%next
		  end if%>
        </select>
</td>
 </tr>	
 
  <tr > 
<td width="12%"   >栏目类型：</td>
<td width="88%">
  <select name="Types" id="Types" >
          <option value="click">关键字</option>
          <option value="view">外部链接</option>
        </select>
</td>
 </tr>	
 
 
   <tr  > 
<td width="12%"   >链接地址：</td>
<td width="88%">
 <input type="text" name="Url" size="50" value="http://"/> 请直接输入链接地址  <font color="#ff0000">当栏目类型为外部链接时有效</font>
</td>
 </tr>	
  
   <tr> 
<td width="12%"   >调用关键字：</td>
<td width="88%">
 <input type="text" name="Keyword" size="50" /> 请输入系统中的关键字  <font color="#ff0000">当栏目类型为关键字时有效</font>
</td>
 </tr>	
 
 
 
 
 
 
 
 
 	   
	   <tr > 
            <td colspan="2"> <input type="submit" name="send" value="保存" class="bnt"/>
            　            </td>
          </tr>	   
</table>	
	
	
	
	
	
    </form>

<%else%>

  <form>
  
  <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >自定义菜单管理</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 <div class="dmeun"><a href="Menu.asp?act=add" target="main" style="color:#ffffff">添加自定义菜单</a></div>
 <div class="dmeun"><a href="javascript:;" class="postmenu" target="main" style="color:#ffffff">发布自定义菜单</a></div>
 <div class="dmeun"><a href="javascript:;" class="getmenu" target="main" style="color:#ffffff">加载远程菜单</a></div>
 <div class="dmeun"><a href="javascript:;" class="delmenu" target="main" style="color:#ffffff">删除自定义菜单</a></div>
 
 </td>
 

 </tr>
</table>	
  
  
  
  
  
      <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
      <tr CLASS="WZ">
        <td>ID</td>
        <td>菜单名称</td>
        <td>类型</td>
        <td>关键字/链接</td>
        <td>管理</td>
      </tr>
      <%datadb=Cls.db.dbload("","id,ClassName,Types,Keyword,Url","[sys_Menu]","fid=0","id asc")
		if ubound(datadb)>0 then
		for i=0 to ubound(datadb,2)%>
      <tr id="list_<%=datadb(0,i)%>">
        <td align="center" ><%=datadb(0,i)%></td>
        <td align="center" ><%=datadb(1,i)%></td>
        <td align="center" ><%if Cls.db.dbcount("[sys_Menu]","Fid="&datadb(0,i))=0 then%><%if datadb(2,i)="view" then%>链接<%else%>关键字<%end if%><%end if%></td>
        <td align="center" ><%if Cls.db.dbcount("[sys_Menu]","Fid="&datadb(0,i))=0 then%><%if datadb(2,i)="view" then%><a href="<%=datadb(4,i)%>" target="_blank"><%=cls.cutstr(datadb(4,i),50,1)%></a><%else%><%=datadb(3,i)%><%end if%><%end if%></td>
        <td align="center" ><a href="?act=edit&id=<%=datadb(0,i)%>">编辑</a>　<a href="javascript:;" class="del" rel="<%=datadb(0,i)%>">删除</a></td>
      </tr>
      <%datadb2=Cls.db.dbload("","id,ClassName,Types,Keyword,Url","[sys_Menu]","fid="&datadb(0,i),"id asc")
		if ubound(datadb2)>0 then
		for n=0 to ubound(datadb2,2)%>
      <tr id="list_<%=datadb2(0,n)%>">
        <td align="center" ><%=datadb2(0,n)%></td>
        <td align="center"><img src="../images/line.gif"><%=datadb2(1,n)%></td>
        <td align="center" ><%if datadb2(2,n)="view" then%>链接<%else%>关键字<%end if%></td>
        <td align="center" ><%if datadb2(2,n)="view" then%><a href="<%=datadb2(4,n)%>" target="_blank"><%=cls.cutstr(datadb2(4,n),50,1)%></a><%else%><%=datadb2(3,n)%><%end if%></td>
        <td align="center" ><a href="?act=edit&id=<%=datadb2(0,n)%>">编辑</a>　<a href="javascript:;" class="del" rel="<%=datadb2(0,n)%>">删除</a></td>
      </tr>
		  <%	next
            end if
            next
		end if%>
    </table>
  </form>
</div>
<%end if%>
</body>
</html>
