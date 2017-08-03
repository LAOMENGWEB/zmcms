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
	dim datadb,datadb2,openid,result
	dim i
	dim act:act=lcase(Cls.fget("act",0))	
	select case act
		case "send":send
	end select
	
	sub send()
		dim Content
		openid=Cls.enhtml(Cls.fget("openid",0))
		Content=Cls.enhtml(Cls.fpost("Content",0))
		datadb=Cls.db.dbload("","WeiXinOpenID","[sys_Config]","id=1","")
		result=PostMsg(datadb(0,0),openid,Content)
		Cls.echo result
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
function checkedit(the,openid)
{
	if($.trim(the.Content.value)=="")
		{
			$.message({content:"回复内容不能为空"});
			the.Content.focus();
			return false;
	}
	var url,data;
	url="?act=send&openid="+openid+"";
	data="Content="+encodeURIComponent($.trim(the.Content.value));
	
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
				$.message({type:"ok",content:"发送成功"});				
				setTimeout("closeDialog()",1500);
				break;
			default:
				alert(_)
				break;
		}
	}
	});
	return false
}
function closeDialog(){
	var list = parent.art.dialog.list;
	for (var i in list) {
		list[i].close();
	};
}

</script>
</head>
<body >

  <form onSubmit="return checkedit(this,'<%=openid%>')">


<table  width="99%"  border="0" cellpadding="0" cellspacing="0"  class="dtable4" align="center">
<tr > 

<td colspan="2">
 <textarea name="Content" cols="50" rows="7"></textarea>
</td>
 </tr>	
 
		   
	   <tr > 
            <td colspan="2">     <input type="submit" name="send" value="回复" class="bnt"/>
        <input type="button" value="返回" onClick="closeDialog()" class="bnt"/>
            　            </td>
          </tr>	   
</table>







       
   
   
   
  </form>
</div>
</body>
</html>
