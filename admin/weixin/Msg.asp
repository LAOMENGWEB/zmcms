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
		case "del":del
		case "delsome":delsome
	end select
	
	sub del()
		dim id:id=Cls.getint(Cls.fget("id",0),0)
		if id>0 then
			Cls.db.dbdel "[sys_Msg]","id="&id&""
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
			Cls.db.dbdel "[sys_Msg]","id in("&id&")"
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
<script>
$(function(){
	$(".replyMsg").click(function(){
		var openid=this.getAttribute('rel');
		var throughBox=$.dialog.through;
			throughBox({
				id:openid,
				padding:0,
				title: "回复消息",				
			<%dim wxtype:wxtype=Cls.db.dbload("","WeiXinType,AppId,Verify","[sys_Config]","id=1","")
			if Cls.strlen(wxtype(1,0))=0 then%>
				content:'请先填写 AppID 和 AppSecret ！',
				ok:function(){
						setTimeout("location.href='set2.asp'",0);		
					  },
			<%elseif wxtype(0,0)="订阅号" or wxtype(2,0)="未认证" then%>
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
</script>

</head>
<body>

  <form>
  
  <table width="100%" border="0" cellpadding="0" cellspacing="0"   class="dtable">
<tr>
 <td >消息管理</td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="../weixin.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 
 
 <div class="dmeun"><a href="javascript:;" class="delsome" target="main" style="color:#ffffff">批量删除</a></div>
 </td>
 



 </tr>
</table>	
  
  
  
  
  
  
  
  
  
       <table width="99%"   cellpadding="0" cellspacing="0"  class="dtable2" align="center">
      <tr>
        <td><input type="checkbox" name="chkall" style="border:0;" onClick="checkall(this.form)" title="全选/取消" /></th>
        <td>ID</td>
        <td>微信头像</td>
        <td>微信用户</td>
        <td>消息类型</td>
        <td>消息内容</td>
        <td>时间戳</td>
        <td>回复状态</td>
        <td>管理</td>
      </tr>      
      <%
	  	dim WeixinAvatar:WeixinAvatar=Cls.db.dbload("","WeixinAvatar","[sys_Config]","id=1","")(0,0)
		if WeixinAvatar="" then
			WeixinAvatar="../images/user.jpg"
		end if
		Set Page = new Page_List
		Page.Con = Cls.db.conn
		Page.Sql = "Select a.id,a.FromUserName,a.BeiZhu,a.MsgType,a.Content,a.Addtime,a.reply,a.autoReply,b.nickname,b.headimgurl from [sys_Msg] a left join [sys_User] b ON a.FromUserName=b.openid order by a.id desc"
		Page.PageSize = 20
		Page.Style = 5
		Set Rs = Page.Rs
		
		If not Rs.Bof or Rs.Eof Then
			For i = 1 To Page.PageSize
			if Cls.strlen(rs("headimgurl"))=0 then
				UserAvatar="../images/user.jpg"
			else
				if rs("reply")=1 then
					UserAvatar=WeixinAvatar
				else
					UserAvatar=rs("headimgurl")
				end if
			end if
			if Cls.strlen(rs("nickname"))=0 then
				WeixinName=rs("FromUserName")
			else
				WeixinName=rs("nickname")
			end if
			%>
      <tr id="list_<%=rs("id")%>">
        <td align="center"><input name="id" type="checkbox" style="border:0;" value="<%=rs("id")%>" /></td>
        <td align="center" ><%=rs("id")%></td>
        <td align="center" ><img src="<%=UserAvatar%>" width="48"></td>
        <td align="center" ><%=WeixinName%></td>
        <td align="center"><%=rs("BeiZhu")%></td>
        <td align="center" ><%if rs("MsgType")="image" then%><a href="<%=rs("Content")%>" target="_blank"><img src="<%=rs("Content")%>" width="64" height="64"></a><%else%><%=rs("Content")%><%end if%></td>
        <td align="center"><%=cls.format_time(rs("Addtime"),1)%></td>
        <td align="center" ><%if rs("reply")=1 then%><span style="color:#f00"">已主动回复</span><%elseif rs("autoReply")=1 then%><span style="color:#00f"">已自动回复</span><%elseif rs("MsgType")="text" then%><a href="KeyWord.asp?act=add&keyword=<%=rs("Content")%>">加入智能回复</a>　<%if dateDiff("h",rs("Addtime"),now)<48 then%><a href="javascript:;" class="replyMsg" rel="<%=rs("FromUserName")%>">主动回复</a><%end if%><%end if%></td>
        <td align="center" ><a href="javascript:;" class="del" rel="<%=rs("id")%>">删除</a></td>
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
</body>
</html>
