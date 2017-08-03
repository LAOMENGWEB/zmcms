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
		case "getlocation":getlocation
	end select	
	
	sub getlocation()
		dim id,lat,lng,LocationXml
		lat=Cls.getint(Cls.fget("lat",0),0)
		lng=Cls.getint(Cls.fget("lng",0),0)
		If lat>0 and lng>0 Then
			LocationXml = GetURL("http://api.map.baidu.com/geocoder/v2/?ak=65f4b9318e96c85bdabcbe884c943ab7&callback=renderReverse&location="& lat &","& lng &"&output=xml&pois=0")
			Cls.Echo LocationXml
		end if
		Cls.die
	End sub

	sub del()
		dim id:id=Cls.getint(Cls.fget("id",0),0)
		if id>0 then
			Cls.db.dbdel "[sys_User_Location]","id="&id&""
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
			Cls.db.dbdel "[sys_User_Location]","id in("&id&")"
			Cls.echo "1"
		end if
		Cls.die
	end sub
	
	
%>

<script>
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

function getlocation(id,lat,lng){
	if(!id==""){
		$.ajax({
			   type: "get",
			   url:"?act=getlocation&lat="+lat+"&lng="+lng,
			   cache:false,
			   timeout:10000,
			   error:function(){$.message({type:"error",content:"服务器错误，操作失败！"});},
			   success: function(data)
			   {   
			   		if ($(data).find("status").text()=="0"){
					   $("#location_"+id).empty();
					   $("#location_"+id).append($(data).find("formatted_address").text());
					}else{
					   $.message({type:"error",content:"数据接口查询出错！"});
					}
			   }
		 });
	}
}
</script>

</head>
<body>
<div id="notice"><span>当前位置：</span>位置 > <a href="?">用户实时地理位置</a></div>
<div class="clear_fixed">
  <div id="menu">
    <dl>
      <dt class="dropdown"><span><a href="javascript:;">批量操作↓</a></span>
        <ul>
          <li><a href="javascript:;" class="delsome">批量删除</a></li>
        </ul>
      </dt>
    </dl>
  </div>
  <form>
    <table cellpadding="5" id="table">
      <tr>
        <th width="20"><input type="checkbox" name="chkall" style="border:0;" onClick="checkall(this.form)" title="全选/取消"/></th>
        <th width="50">编号</th>
        <th width="100">微信头像</th>
        <th width="200">微信名称</th>
        <th width="200">更新时间</th>
        <th>实时地理位置信息</th>
        <th width="50">管理</th>
      </tr>      
      <%
	  	OpenId=Cls.enhtml(Cls.fget("OpenId",0))
	  	if cls.strlen(OpenId)>0 then
			SqlStr="where a.OpenId='"&OpenId&"'"
		end if
		Set Page = new Page_List
		Page.Con = Cls.db.conn
		Page.Sql = "Select a.Id,a.OpenId,a.Latitude,a.Longitude,a.Precision,a.AddTime,b.nickname,b.headimgurl from [sys_User_Location] a left join [sys_User] b on a.OpenId=b.openid "&SqlStr&" order by a.id desc"
		Page.PageSize = 15
		Set Rs = Page.Rs
		
		If not Rs.Bof or Rs.Eof Then
			For i = 1 To Page.PageSize
			if Cls.strlen(rs("headimgurl"))=0 then
				UserAvatar="../images/user.jpg"
			else
				UserAvatar=rs("headimgurl")
			end if
			if Cls.strlen(rs("nickname"))=0 then
				WeixinName=rs("OpenId")
			else
				WeixinName=rs("nickname")
			end if
				%>
      <tr id="list_<%=rs("id")%>">
        <td align="center"><input name="id" type="checkbox" style="border:0;" value="<%=rs("Id")%>" /></td>
        <td align="center" ><%=rs("id")%></td>
        <td align="center" ><img src="<%=UserAvatar%>" width="64"></td>
        <td align="center"><%=WeixinName%></td>
        <td align="center" ><%=rs("AddTime")%></td>
        <td align="center" id="location_<%=rs("id")%>"><%=rs("Latitude")%> / <%=rs("Longitude")%> / <%=rs("Precision")%> <a href="javascript:getlocation('<%=rs("id")%>','<%=rs("Latitude")%>','<%=rs("Longitude")%>')">【切换为具体地址】</a></td>
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
