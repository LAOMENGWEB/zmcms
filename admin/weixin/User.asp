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

	dim datadb,datadb2,openidlist
	dim i
	dim nickname,sex,language,city,province,country,headimgurl,subscribe_time
	dim act:act=lcase(Cls.fget("act",0))
	select case act
		case "del":del
		case "delsome":delsome
		case "getlist":getlist
		case "getinfo":getinfo
		case "getuserinfo":getuserinfo
	end select
	
	sub del()
		dim id:id=Cls.getint(Cls.fget("id",0),0)
		if id>0 then
			Cls.db.dbdel "[sys_User]","id="&id&""
			Cls.echo "1"
		else			
			Cls.echo "0数据出错"
		end if
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
			Cls.db.dbdel "[sys_User]","id in("&id&")"
			Cls.echo "1"
		end if
		Cls.die
	end sub
	
	'获取粉丝列表
	sub getlist()
		dim result
		Access_token=GetToken()
		if Cls.strlen(Access_token)=0 then Cls.echo "0接口出错，无法获取Access_token":exit sub
		strJson=GetURL(GetUserListUrl&"access_token="&Access_token)
		Call InitScriptControl:Set objTest = getJSONObject(strJson)
		if objTest.total>0 then
			openidlist=objTest.data.openid
			openidlist=split(openidlist,",") 
			for i=0 to ubound(openidlist) 
				result=addUser(openidlist(i))
			next
			Cls.echo "1"
		else
			Cls.echo "0暂无粉丝"
		end if
		cls.die
	end sub
	
	'获取粉丝信息
	sub getinfo()
		'nickname,sex,language,city,province,country,headimgurl,subscribe_time
		openid=Cls.enhtml(Cls.fget("openid",0))
		Access_token=GetToken()
		if Cls.strlen(Access_token)=0 then Cls.echo "0接口出错，无法获取Access_token":exit sub
		strJson=GetURL(GetUserInfoUrl&"access_token="&Access_token&"&openid="&openid)
		Call InitScriptControl:Set objTest = getJSONObject(strJson)
		if objTest.subscribe>0 then
			nickname=objTest.nickname
			sex=objTest.sex
			language=objTest.language
			city=objTest.city
			province=objTest.province
			country=objTest.country
			headimgurl=objTest.headimgurl
			subscribe_time=FromUnixTime(objTest.subscribe_time)
			data=array(array("nickname",nickname,200,1),array("sex",sex,50,1),array("language",language,50,1),array("city",city,50,1),array("province",province,50,1),array("country",country,50,1),array("headimgurl",headimgurl,0,1),array("subscribe_time",subscribe_time,0,1))
			if Cls.db.dbupdate("[sys_User]","openid='"&openid&"'",data)=1 then
				cls.echo "1"
			else
				cls.echo "0数据出错"
			end if
		else
			cls.echo "0获取用户信息出错"
		end if		
		cls.die
	end sub
	
	Sub getuserinfo()
		openids=Cls.enhtml(Cls.fget("openids",0))
		If Cls.strlen(openids)=0 Then cls.echo "0该页无需更新粉丝信息":Exit Sub
		Access_token=GetToken()
		if Cls.strlen(Access_token)=0 then Cls.echo "0接口出错，无法获取Access_token":exit sub
		
		openids = Split(openids, ",", -1, 1)
		For i = 0 To UBound(openids)-1
			strJson=GetURL(GetUserInfoUrl&"access_token="&Access_token&"&openid="&openids(i))
			Call InitScriptControl:Set objTest = getJSONObject(strJson)
			if objTest.subscribe>0 then
				nickname=objTest.nickname
				sex=objTest.sex
				language=objTest.language
				city=objTest.city
				province=objTest.province
				country=objTest.country
				headimgurl=objTest.headimgurl
				subscribe_time=FromUnixTime(objTest.subscribe_time)
				data=array(array("nickname",nickname,200,1),array("sex",sex,50,1),array("language",language,50,1),array("city",city,50,1),array("province",province,50,1),array("country",country,50,1),array("headimgurl",headimgurl,0,1),array("subscribe_time",subscribe_time,0,1))
				if not Cls.db.dbupdate("[sys_User]","openid='"&openids(i)&"'",data)=1 then
					cls.echo "0数据出错" :exit sub
				end if
			else
				cls.echo "0获取用户信息出错" :exit sub
			end if	
		Next		
		cls.echo "1"
		cls.die
	end Sub
	
	function addUser(t0)
		data=array(array("openid",t0,255,1),array("headimgurl","../images/user.jpg",0,1),array("nickname","请手动刷新用户信息",0,1))
		if Cls.db.dbnew("[sys_User]",data,"openid='"&t0&"'")=1 then
			addUser=1
		else
			addUser=0
		end if	
	end function
	
%>

<link href="../../weixin/css/wx.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../../weixin/js/jquery.js"></script>
<script language="javascript" src="../../weixin/js/jquery.artDialog.js?skin=default"></script>
<script language="javascript" src="../../weixin/js/artDialog.iframeTools.js"></script>
<script language="javascript" src="../../weixin/js/base.js"></script>
<script language="javascript" src="../../weixin/js/admin.js"></script>

<script>
$(function(){
	<%dim appDb:appDb=Cls.db.dbload("","AppId,Appsecret,WeiXinType,Verify","[sys_Config]","id=1","")
	if Cls.strlen(appDb(0,0))=0 or Cls.strlen(appDb(1,0))=0 then%>
	 var AppBox=$.dialog.through;
	  AppBox({
		  content:'请先填写 AppID 和 AppSecret ！',
		  ok:function(){setTimeout("location.href='set2.asp'",0);},
		  lock:false,
		  opacity:'0.5'
	  })
	<%elseif appDb(2,0)="订阅号" or appDb(3,0)="未认证" then%>
	 var AppBox=$.dialog.through;
	  AppBox({
		  content:'对不起，只有<span style="color:#F00">认证过</span>的服务号才可以使用该接口 ！',
		  ok:function(){setTimeout("location.href='set.asp'",0);},
		  lock:false,
		  opacity:'0.5'
	  })
	<%end if%>
	<%dim wxtype:wxtype=Cls.db.dbload("","WeiXinType","[sys_Config]","id=1","")(0,0)%>
	$(".sendmsg").click(function(){
		var openid=this.getAttribute('rel');
		var throughBox=$.dialog.through;
			throughBox({
				id:openid,
				padding:0,
				title: "回复消息",	
			<%if wxtype="订阅号" then%>
		  		content:'对不起，只有<span style="color:#F00">认证过</span>的服务号才可以使用该接口 ！',
				okVal:'确定',
				ok:true,
			<%else%>
				content: "<iframe src='replyMsg.asp?openid="+openid+"' width='500' height='200' frameborder='0' id='replyMsg'></iframe>",
			<%end if%>
				padding: "10px",
				lock:false,
				opacity: "0.5"
			})
		
		})
	
	$(".getUserList").click(function(){
		var throughBox=$.dialog.through;
			throughBox({
					   
				lock:false,
				opacity:'0.5',
			<%if wxtype="订阅号" then%>
		  		content:'对不起，只有<span style="color:#F00">认证过</span>的服务号才可以使用该接口 ！',
				okVal:'确定',
				ok:true
			<%else%>				
				icon:'question',
				lock:false,
				content:'确定需要刷新粉丝列表？需要一定的时间！',
				ok:function(){
					var url='?act=getlist';
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
							$.message({type:"ok",content:"粉丝刷新成功"});
							setTimeout("location.href='?'",1000);
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			<%end if%>
			})
		
		})
	
	$(".getinfo").click(function(){
		var openid=this.getAttribute('rel');
		var throughBox=$.dialog.through;
			throughBox({			
				lock:false,
				opacity:'0.5',
			<%if wxtype="订阅号" then%>
		  		content:'对不起，只有<span style="color:#F00">认证过</span>的服务号才可以使用该接口 ！',
				okVal:'确定',
				ok:true
			<%else%>
				icon:'question',
				content:'确定需要更新粉丝信息？需要一定的时间！',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='?act=getinfo&openid='+openid;
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
							$.message({type:"ok",content:"粉丝信息更新成功"});
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			<%end if%>
			})
		
		})
	
	$(".getuserinfo").click(function(){
		var openids = $("#noInfos").val();
		var throughBox=$.dialog.through;
			throughBox({
				icon:'question',
				content:'确定需要批量更新粉丝信息？需要一定的时间！',
				lock:false,
				opacity:'0.5',
				ok:function(){
					var url='?act=getuserinfo&openids='+openids;
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
							$.message({type:"ok",content:"更新成功"});
						}
					}
					});
									
					},
				cancelVal:'取消',
				cancel:true 
			})
		
		})
	
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
</script>

</head>
<body>

  <form>
  
  
  
    <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >粉丝列表</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td >
  <div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 <div class="dmeun"><a href="javascript:;" class="delsome" target="main" style="color:#ffffff">批量删除</a></div>
 <div class="dmeun"><a href="javascript:;" class="getUserList" target="main" style="color:#ffffff">刷新粉丝列表</a></div>
 <div class="dmeun"><a href="javascript:;" class="getuserinfo" target="main" style="color:#ffffff;width:200px">批量获取粉丝头像</a></div>

  <dt>　说明:由于粉丝数量过大，刷新粉丝列表后，需要手动更新用户头像信息</dt>
 
 </td>
 

 </tr>
</table>
  
  
  
  
  
  
  
  
 <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
      <tr CLASS="WZ">
        <tD><input type="checkbox" name="chkall" style="border:0;" onClick="checkall(this.form)" title="全选/取消"/></td>
        <td>编号</td>
        <td>微信头像</td>
        <td>微信名称</td>
        <td>性别</td>
        <td>关注时间</td>
        <td>地理位置信息</td>
        <td>管理</td>
      </tr>      
      <%
	  	dim noInfos
		Set Page = new Page_List
		Page.Con = Cls.db.conn
		Page.Sql = "Select * from [sys_User] order by id desc"
		Page.PageSize = 20
		Page.Style = 5
		Set Rs = Page.Rs
		
		If not Rs.Bof or Rs.Eof Then
			For i = 1 To Page.PageSize
				if IsNull(Rs("Subscribe_Time")) then
					noInfos=noInfos&Rs("openid")&","
				end if
				%>
      <tr id="list_<%=rs("id")%>">
        <td align="center"><input name="id" type="checkbox" style="border:0;" value="<%=rs("id")%>" /></td>
        <td align="center" ><%=rs("id")%></td>
        <td align="center" ><img src="<%=rs("headimgurl")%>" width="64"></td>
        <td align="center"><%=rs("nickname")%></td>
        <td align="center"><%if rs("sex")="1" then%>男生<%elseif rs("sex")="2" then%>女生<%else%>保密<%end if%></td>
        <td align="center" ><%=rs("subscribe_time")%></td>
        <td align="center" ><%=rs("country")%> / <%=rs("province")%> / <%=rs("city")%> </td>
        <td align="center" ><a href="javascript:;" class="getinfo" rel="<%=rs("openid")%>">更新用户信息</a>　<a href="javascript:;" class="sendmsg" rel="<%=rs("openid")%>">发送消息</a>　<a href="javascript:;" class="del" rel="<%=rs("id")%>">删除</a></td>
      </tr>
                <%
				Rs.MoveNext
			Next
		End If
		
		Call Page.ShowPage
		
		Rs.Close
		Set Rs = Nothing
		Set Page = Nothing%> <input name="noInfos" id="noInfos" type="hidden" value="<%=noInfos%>">
    </table>
   
  </form>
</div>
</body>
</html>
