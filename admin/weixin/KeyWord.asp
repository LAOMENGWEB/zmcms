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
		case "delsome":delsome
	end select
	
	sub adddb()
		dim k_classId,k_keyWord,k_keyWord_beizhu,k_keyType,k_type,k_text,k_plugName,k_plugParam,k_addtime,k_num
		
		k_classId=Cls.getint(Cls.fpost("k_classId",0),0)
		k_keyWord=Cls.enhtml(Cls.fpost("k_keyWord",0))
		k_keyWord_beizhu=Cls.enhtml(Cls.fpost("k_keyWord_beizhu",0))
		k_keyType=Cls.enhtml(Cls.fpost("k_keyType",0))
		k_type=Cls.enhtml(Cls.fpost("k_type",0))
		k_text=Cls.enhtml(Cls.fpost("k_text",0))
		k_plugName=Cls.enhtml(Cls.fpost("k_plugName",0))
		k_plugParam=Cls.enhtml(Cls.fpost("k_plugParam",0))
		k_num=Cls.getint(Cls.fpost("k_num",0),0)
		k_addtime=now()
		
		Dim data
		data=array(array("k_classId",k_classId,0,0),array("k_num",k_num,0,0),array("k_keyWord",k_keyWord,255,1),array("k_keyWord_beizhu",k_keyWord_beizhu,255,1),array("k_keyType",k_keyType,50,1),array("k_type",k_type,50,1),array("k_text",k_text,0,1),array("k_plugName",k_plugName,50,1),array("k_plugParam",k_plugParam,0,1),array("k_hits","0",0,0),array("k_addtime",k_addtime,50,0))
		if Cls.db.dbnew("[sys_KeyWord]",data,"k_keyWord='"&k_keyWord&"'")=1 then
			Cls.echo "1"
		else
			Cls.echo "0该关键字已存在"
		end if	
		Cls.die
	end sub
	
	
	sub editdb()
		dim id:id=Cls.getint(Cls.fget("id",0),0)		
		dim k_classId,k_keyWord,k_keyType,k_keyWord_beizhu,k_type,k_text,k_plugName,k_plugParam,k_addtime,k_num
		
		k_classId=Cls.getint(Cls.fpost("k_classId",0),0)
		k_keyWord=Cls.enhtml(Cls.fpost("k_keyWord",0))
		k_keyWord_beizhu=Cls.enhtml(Cls.fpost("k_keyWord_beizhu",0))
		k_keyType=Cls.enhtml(Cls.fpost("k_keyType",0))
		k_type=Cls.enhtml(Cls.fpost("k_type",0))
		k_text=Cls.enhtml(Cls.fpost("k_text",0))
		k_plugName=Cls.enhtml(Cls.fpost("k_plugName",0))
		k_plugParam=Cls.enhtml(Cls.fpost("k_plugParam",0))
		k_num=Cls.getint(Cls.fpost("k_num",0),0)
		
		Dim data
		data=array(array("k_classId",k_classId,0,0),array("k_num",k_num,0,0),array("k_keyWord",k_keyWord,255,1),array("k_keyWord_beizhu",k_keyWord_beizhu,255,1),array("k_keyType",k_keyType,50,1),array("k_type",k_type,50,1),array("k_text",k_text,0,1),array("k_plugName",k_plugName,50,1),array("k_plugParam",k_plugParam,0,1))
		if Cls.db.dbupdate("[sys_KeyWord]","k_id="&id,data)=1 then
			Cls.echo "1"
		else
			Cls.echo "0数据出错"
		end if	
		Cls.die	
	end sub
	
	
	sub del()
		dim id:id=Cls.getint(Cls.fget("id",0),0)
		if id>0 then
			Cls.db.dbdel "[sys_KeyWord]","k_id="&id&""
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
				end if
			next
			Cls.db.dbdel "[sys_KeyWord]","k_id in("&id&")"
			Cls.echo "1"
		end if
		Cls.die
	end sub
	
%>
<link href="../../weixin/css/wx.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../../weixin/js/jquery.js"></script>
<script language="javascript" src="../../weixin/js/jquery.artDialog.js?skin=default"></script>
<script language="javascript" src="../../weixin/js/artDialog.iframeTools.js"></script>
<script language="javascript" src="../../weixin/js/base.js"></script>
<script language="javascript" src="../../weixin/js/admin.js"></script>
<script type="text/javascript" src="../../js/jquery.SuperSlide.2.1.1.js"></script>
<script>




function checkadd(the)
{
	if($.trim(the.k_keyWord.value)=="")
	{
		if(the.k_msgtype.value=="text"){
			$.message({content:"关键字不能为空"});
			the.k_keyWord.focus();
			return false;
		}
	}
	if($.trim(the.k_keyWord_beizhu.value)=="")
	{
		$.message({content:"关键字备注不能为空"});
		the.k_keyWord_beizhu.focus();
		return false;
	}
	if($('#k_type option:selected').val()=="文本")
	{
		if($.trim(the.k_text.value)=="")
		{
			$.message({content:"回复内容不能为空"});
			the.k_text.focus();
			return false;
		}
	}else{		
		if($.trim(the.k_plugName.value)=="")
		{
			$.message({content:"请选择插件"});
			return false;
		}		
		if($.trim(the.k_plugParam.value)=="")
		{
			$.message({content:"请选择数据"});
			return false;
		}
	}
	var url,data;
	url="?act=adddb";
	data="k_keyWord="+encodeURIComponent($.trim(the.k_keyWord.value));
	data+="&k_keyWord_beizhu="+encodeURIComponent($.trim(the.k_keyWord_beizhu.value));
	data+="&k_classId="+encodeURIComponent($.trim(the.k_classId.value));
	data+="&k_keyType="+encodeURIComponent($.trim(the.k_keyType.value));
	data+="&k_type="+encodeURIComponent($.trim(the.k_type.value));
	data+="&k_text="+encodeURIComponent($.trim(the.k_text.value));
	data+="&k_plugName="+encodeURIComponent($.trim(the.k_plugName.value));
	data+="&k_plugParam="+encodeURIComponent($.trim(the.k_plugParam.value));
	data+="&k_num="+encodeURIComponent($.trim(the.k_num.value));
	
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
	if($.trim(the.k_keyWord.value)=="")
	{
		if(the.k_msgtype.value=="text"){
			$.message({content:"关键字不能为空"});
			the.k_keyWord.focus();
			return false;
		}
	}
	if($.trim(the.k_keyWord_beizhu.value)=="")
	{
		$.message({content:"关键字备注不能为空"});
		the.k_keyWord_beizhu.focus();
		return false;
	}
	if($('#k_type option:selected').val()=="文本")
	{
		if($.trim(the.k_text.value)=="")
		{
			$.message({content:"回复内容不能为空"});
			the.k_text.focus();
			return false;
		}
	}else{		
		if($.trim(the.k_plugName.value)=="")
		{
			$.message({content:"请选择插件"});
			return false;
		}	
		if($.trim(the.k_plugParam.value)=="")
		{
			$.message({content:"请选择数据"});
			return false;
		}
	}
	var url,data;
	url="?act=editdb&id="+id+"";
	data="k_keyWord="+encodeURIComponent($.trim(the.k_keyWord.value));
	data+="&k_keyWord_beizhu="+encodeURIComponent($.trim(the.k_keyWord_beizhu.value));
	data+="&k_classId="+encodeURIComponent($.trim(the.k_classId.value));
	data+="&k_keyType="+encodeURIComponent($.trim(the.k_keyType.value));
	data+="&k_type="+encodeURIComponent($.trim(the.k_type.value));
	data+="&k_text="+encodeURIComponent($.trim(the.k_text.value));
	data+="&k_plugName="+encodeURIComponent($.trim(the.k_plugName.value));
	data+="&k_plugParam="+encodeURIComponent($.trim(the.k_plugParam.value));
	data+="&k_num="+encodeURIComponent($.trim(the.k_num.value));
	
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
				content:'确定要删除？不可恢复！',
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

var oldBeizhu;
function setBeizhu(t){
	if(t!='text'){
		oldBeizhu=$("#beizhu").val();
		if(t=='image'){
			$("#beizhu").val('接收图片信息');
		}else if(t=='voice'){
			$("#beizhu").val('接收语音信息');
		}else if(t=='video'){
			$("#beizhu").val('接收视频信息');
		}
		$("#beizhu").attr("disabled","disabled");
	}else{
		$("#beizhu").val(oldBeizhu);
		$("#beizhu").attr("disabled","");
	}
}

function setclose(t0)
{
	if(t0=="文本"){$("#plug").css("display","none");$("#textmsg").css("display","block");}else{$("#plug").css("display","block");$("#textmsg").css("display","none");}
}
function checksearch(a) {
    return "" == a.key.value || "请输入关键字" == a.key.value ? ($.message({
        content: "请输入关键字"
    }), a.key.focus(), !1) : void 0
}
</script>
</head>
<body>

	<%
	if act="edit" then	
	dim id:id=Cls.getint(Cls.fget("id",0),0)
	key=Cls.enhtml(Cls.fget("key",0))
	If cls.strlen(key)>0 then
		SqlStr="k_keyWord='"&key&"'"
	else
		SqlStr="k_id="&id
	end if
	datadb=Cls.db.dbload("","k_classId,k_keyWord,k_keyType,k_type,k_text,k_plugName,k_plugParam,k_id,k_num,k_keyWord_beizhu","[sys_KeyWord]",SqlStr,"")
	if not ubound(datadb)>=0 then
	cls.echo "数据出错"
	cls.die
	end if
	id=datadb(7,0)
	%>
    <script>	
	$(function(){
		$("#<%if datadb(3,0)="文本" then%>plug<%else%>textmsg<%end if%>").css("display","none");
	})
	getdata('<%=datadb(5,0)%>','<%=datadb(6,0)%>');
	</script>

    <form onSubmit="return checkedit(this,<%=id%>)">
	
<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >编辑关键词</td>
</tr>
</table>
	
	
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 
 
  <div class="dmeun"><a href="keyword.asp?act=add" target="main" style="color:#ffffff">添加关键字</a></div>
 
 </td>
 

 </tr>
</table>	
	
	
<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">

<tr > 
<td width="10%"   >接收类型：</td>
<td width="88%">
   <select name="k_msgtype" id="k_msgtype" onChange="setBeizhu(this.value)">
          <option value="text">文本</option>
		  <!--
          <option value="image" <%if datadb(9,0)="image" then cls.echo " selected"%>>图片</option>
          <option value="voice" <%if datadb(9,0)="voice" then cls.echo " selected"%>>语音</option>
          <option value="video" <%if datadb(9,0)="video" then cls.echo " selected"%>>视频</option>
        </select> -->
</td>
 </tr>	

 
 <tr > 
<td width="12%"   >关键字优先级别：</td>
<td width="88%">
<input type="text" name="k_num" size="4" value="<%=datadb(8,0)%>"/>
        关键字匹配的优先级别，数字越小的级别越高。</td>
 </tr>	
 
  <tr > 
<td width="12%"   >关键字：</td>
<td width="88%">
 <input type="text" name="k_keyWord" size="50" value="<%=datadb(1,0)%>"/>
        设定关键字，当用户发送此关键字时将会自动回复消息。</td>
 </tr>	
 
   <tr > 
<td width="12%"   >关键字备注：</td>
<td width="88%">
 <input type="text" id="beizhu" name="k_keyWord_beizhu" size="50" value="<%=datadb(9,0)%>"/>
        设定关键字备注</td>
 </tr>	
 
    <tr > 
<td width="12%"   >所属分类：</td>
<td width="88%">
 <select name="k_classId" id="k_classId">
      	<%datadb2=Cls.db.dbload("","c_id,c_ClassName","[sys_KeyWord_class]","","c_id asc")
		if ubound(datadb2)>=0 then
		for i=0 to ubound(datadb2,2)%>
          <option value="<%=datadb2(0,i)%>" <%if datadb(0,0)=datadb2(0,i) then cls.echo " selected"%>><%=datadb2(1,i)%></option>
      <%next
	  end if%>
        </select>  本项只作为后台备注分类，不影响正常使用</td>
 </tr>	
 
 
     <tr > 
<td width="12%"   >匹配类型：</td>
<td width="88%">
  <select name="k_keyType" id="k_keyType">
          <option value="完全匹配" <%if datadb(2,0)="完全匹配" then cls.echo " selected"%>>完全匹配</option>
          <option value="左匹配" <%if datadb(2,0)="左匹配" then cls.echo " selected"%>>左匹配</option>
          <option value="右匹配" <%if datadb(2,0)="右匹配" then cls.echo " selected"%>>右匹配</option>
          <option value="模糊匹配" <%if datadb(2,0)="模糊匹配" then cls.echo " selected"%>>模糊匹配</option>
        </select></td>
 </tr>
 
      <tr > 
<td width="12%"   >回复类型：</td>
<td width="88%">
  <select name="k_type" id="k_type" onChange="setclose(this.value)">
          <option value="文本" <%if datadb(3,0)="文本" then Cls.echo "selected"%>>文本</option>
          <option value="插件" <%if datadb(3,0)="插件" then Cls.echo "selected"%>>插件</option>
        </select>
        【文本】为直接回复文本信息，【插件】为调用系统中的插件进行处理关键字</td>
 </tr>
 
 
       <tr> 
<td>指定调用插件：
 </td>
 <td><select name="k_plugName" id="k_plugName" onChange="getdata(this.value,<%if cls.strlen(datadb(6,0))>0 then%><%=datadb(6,0)%><%else%>0<%end if%>)">
          <option value="">请选择调用插件</option>
      	<%datadb2=Cls.db.dbload("","p_Name,p_Title,p_Admin_Api","[Plug]","(p_Type=1 or p_Type=2) and p_IsLock='开启'","p_num asc")
		if ubound(datadb2)>=0 then
		for i=0 to ubound(datadb2,2)%>
          <option value="<%=datadb2(0,i)%>" <%if datadb(5,0)=datadb2(0,i) then Cls.echo "selected"%>><%=datadb2(1,i)%></option>
      <%next
	  end if%>
          </select> 请选择数据：
        <select name="k_plugParam" id="k_plugParam"></select>  <font color="#ff0000">当回复类型为插件时有效</font></td>
    </tr>
 
 
       <tr > 
<td>回复内容：</td>
 <td><textarea name="k_text" cols="60" rows="20"><%=datadb(4,0)%></textarea> <font color="#ff0000">当回复类型为文本时有效</font></td>
    </tr>
 
 
 
 
 	   
	   <tr > 
            <td colspan="2"> <input type="submit" name="send" value="保存" class="bnt"/>
            　            </td>
          </tr>	   
</table>		
	
	
	
    </form>

<%elseif act="add" then%>
<%keyword=Cls.enhtml(Cls.fget("keyword",0))%>

    <form onSubmit="return checkadd(this)">
	
	
	
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
      	<%datadb=Cls.db.dbload("","c_id,c_ClassName","[sys_KeyWord_class]","c_id<>1","c_id asc")
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
      	<%datadb=Cls.db.dbload("","p_Name,p_Title","[Plug]","(p_Type=1 or p_Type=2) and p_IsLock='开启'","p_num asc")
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

<%else

classid=Cls.getint(Cls.fget("classid",0),0)
key=Cls.enhtml(Cls.fget("key",0))
k_type=Cls.enhtml(Cls.fget("k_type",0))
k_keyType=Cls.enhtml(Cls.fget("k_keyType",0))
p_Title=Cls.enhtml(Cls.fget("p_Title",0))
%>


<table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >关键字管理</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div></td>
 

 </tr>
</table>
<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
 <div >
<ul class="nav1">
<li class="m" >
<a href="javascript:;" class="delsome">批量删除</a>
</li>
<li class="m" >
<a href="keyword.asp">关键字管理</a>
</li>
<li class="m" >
<a href="?act=add">添加关键字</a>
</li>


     <li class="m" >
	 <a href="javascript:;">所属分类</a>
	 <ul class="sub">
       <%datadb2=Cls.db.dbload("","c_id,c_ClassName","[sys_KeyWord_class]","","c_id asc")
		if ubound(datadb2)>=0 then
		for i=0 to ubound(datadb2,2)%>
                <li><a href="?classid=<%=datadb2(0,i)%>"><%=datadb2(1,i)%></a></li>
  <%next
	  end if%>
            </ul>
     </li>



   <li class="m" >
   <a href="javascript:;">回复类型</a>
    <ul class="sub">
                <li><a href="?k_type=文本">文本</a></li>
                <li><a href="?k_type=插件">插件</a></li>
            </ul>
      </li>


   <li class="m" >
   <a href="javascript:;">匹配类型</a>
   <ul class="sub">
                <li><a href="?k_keyType=完全匹配">完全匹配</a></li>
                <li><a href="?k_keyType=左匹配">左匹配</a></li>
                <li><a href="?k_keyType=右匹配">右匹配</a></li>
                <li><a href="?k_keyType=模糊匹配">模糊匹配</a></li>
            </ul>
      </li>


    <li class="m" ><a href="KeyWord_Hits.asp" style="width:150px">查看关键字匹配记录</a></li>
	  
	<li class="m" ><a href="keyword_class.asp?act=add" target="main">添加关键字分类</a></li>

	
</ul>
</div>
 

<script type="text/javascript"> jQuery(".nav1").slide({ type:"menu",  titCell:".m", targetCell:".sub", effect:"slideDown", delayTime:300, triggerTime:100,returnDefault:true  });</script>



  <div id="menu">
  <div class="search"><form action="?" method="get" onSubmit="return checksearch(this)"><input type="text" name="key"  value="<%if key<>"" then%><%=key%><%else%>请输入关键字<%end if%>" onFocus="if(this.value=='请输入关键字')this.value=''" onBlur="if(this.value=='')this.value='请输入关键字'" /><input type="submit" class="bnt" value="查询" /></form></div>
  </td>
 

 </tr>
</table>	

  <form>
  
  
    <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
               
                <tr class="wz">
                <td ><input type="checkbox" name="chkall" style="border:0;" onClick="checkall(this.form)" title="全选/取消" /></th>
        <td >ID</th>
        <td >优先级别</td>
        <td >所属分类</td>
        <td >回复类型</td>
        <td>匹配类型</td>
        <td >关键字</td>
        <td>关键字备注</td>
        <td>调用次数</td>
        <td >添加时间</td>
        <td >管理</td>
                </tr>
  
  
  
  
  
  
        
        <%
		dim whereStr
		whereStr="where 1=1 "
		If key<>"" then
			whereStr=whereStr&"and a.k_keyWord like '%"&key&"%'"
		end if
		If classid<>0 then
			whereStr=whereStr&"and a.k_classId="&classid&""
		end if
		If k_type<>"" then
			if p_Title<>"" then
				whereStr=whereStr&"and c.p_Title ='"&p_Title&"'"
			else
				whereStr=whereStr&"and a.k_type ='"&k_type&"'"
			end if
		end if
		If k_keyType<>"" then
			whereStr=whereStr&"and a.k_keyType ='"&k_keyType&"'"
		end if
		Set Page = new Page_List
		Page.Con = Cls.db.conn		
		dim SqlStr:SqlStr="Select a.k_id,a.k_keyWord,a.k_keyType,a.k_type,a.k_text,a.k_hits,a.k_addtime,a.k_num,a.k_keyWord_beizhu,b.c_ClassName,b.c_id,c.p_Title from ([sys_KeyWord] a left join [sys_KeyWord_Class] b ON a.k_classId=b.c_id) left join [Plug] c ON a.k_plugName=c.p_Name "&whereStr&" order by a.k_id desc"
		Page.Sql = SqlStr
		Page.PageSize = 20
		Page.Style = 5
		Set Rs = Page.Rs
		
		If not Rs.Bof or Rs.Eof Then
			For i = 1 To Page.PageSize
				%>
                
      <tr id="list_<%=rs("k_id")%>">
        <td align="center"><input name="id" type="checkbox" style="border:0;" value="<%=rs("k_id")%>" /></td>
        <td align="center" ><%=rs("k_id")%></td>
        <td align="center" ><%=rs("k_num")%></td>
        <td align="center" ><a href="?classid=<%=rs("c_id")%>"><%=rs("c_ClassName")%></a></td>
        <td align="center" ><%if rs("k_type")="插件" then%><a href="?k_type=插件&p_Title=<%=rs("p_Title")%>"><%=rs("p_Title")%>插件</a><%else%><a href="?k_type=文本">文本</a><%end if%></td>
        <td align="center"><a href="?k_keyType=<%=rs("k_keyType")%>"><%=rs("k_keyType")%></a></td>
        <td align="center"><%=rs("k_keyWord")%></td>
        <td align="center"><%=rs("k_keyWord_beizhu")%></td>
        <td align="center"><a href="KeyWord_Hits.asp"><%=rs("k_hits")%></a></td>
        <td align="center" ><%=cls.format_time(rs("k_addtime"),1)%></td>
        <td align="center" ><a href="?act=edit&id=<%=rs("k_id")%>">编辑</a>　<a href="javascript:;" class="del" rel="<%=rs("k_id")%>">删除</a></td>
      </tr><%
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
