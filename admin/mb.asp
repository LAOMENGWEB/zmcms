<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->

<%
t=request.QueryString("t")
if Instr(session("AdminPurview"),"|104,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<script language="JavaScript" src="../js/jquery-1.7.2.min.js"></script>
 <script type="text/javascript" src="../js/Validform_v5.3.2_min.js"></script>
 <link href="css/style.css" rel="stylesheet" type="text/css">
	 <style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
  <script type="text/javascript">
$(function(){
	
$(".checkform").Validform();  //就这一行代码！;
	
})
</script>

<form name="form1" method="post" action="mbedit.asp">



<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center"><%if t=1 then%> PC端模板管理 <%end if%>
 <%if t=2 then%> WAP端模板管理 <%end if%></div></td>
</tr>
</table>


<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="template.asp" target="main" style="color:#ffffff">返回菜单</a></div>
 <%if t=1 then%>
 <div class="dmeun"><a href="mb.asp?t=1" target="main" style="color:#ffffff">PC端模板管理</a></div>
 
  <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="mb.asp?t=2" target="main" style=" color:#999999">WAP端模板管理</a></div>
  
  <%end if%>
   <%if t=2 then%>
 <div class="dmeun" style=" background:#eeeeee; color:#999999"><a href="mb.asp?t=1" target="main" style=" color:#999999">PC端模板管理</a></div>
 
  <div class="dmeun"><a href="mb.asp?t=2" target="main" style="color:#ffffff">WAP端模板管理</a></div>
  
  <%end if%>
 </td>
 

 </tr>
</table>






            <%if t=1 then%>
              <table width="99%" border="0" cellpadding="0" cellspacing="0"  class="dtable2" align="center">
        


                <tr  class="wz">
                  <td  >文件</td>
                  <td  >文件名称</td>
                  <td >操作</td>
                </tr>
                
                <tr  >
                  <td align="center" >公用顶部</td>
                  <td  align="center" ><%=template%>top.asp</td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>top.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >公用底部</td>
                  <td  align="center" ><%=template%>bottom.asp</td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>bottom.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >首页</td>
                  <td  align="center" ><%=template%>default.asp</td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>default.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >关于我们</td>
                  <td  align="center" ><label><%=template%>about.asp</label></td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>about.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >新闻列表页</td>
                  <td  align="center" ><%=template%>info.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>info.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >新闻内容页</td>
                  <td  align="center" ><%=template%>showinfo.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>showinfo.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >产品列表页</td>
                  <td  align="center" ><%=template%>product.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>product.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >产品内容页</td>
                  <td  align="center" ><%=template%>showproduct.asp</td>
                  <td colspan="-2" align="center"style="border-right:none" ><a href="mbedit.asp?file=<%=template%>showproduct.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >图片集列表页</td>
                  <td  align="center" ><%=template%>photo.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>photo.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >图片集内容页</td>
                  <td  align="center" ><%=showphoto%>/index.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=../<%=showphoto%>/index.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >视频列表页</td>
                  <td  align="center" ><%=template%>video.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>video.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >视频内容页</td>
                  <td  align="center" ><%=template%>showvideo.asp</td>
                  <td colspan="-2" align="center"style="border-right:none" ><a href="mbedit.asp?file=<%=template%>showvideo.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >下载列表页</td>
                  <td  align="center" ><%=template%>down.asp</td>
                  <td colspan="-2" align="center"style="border-right:none" ><a href="mbedit.asp?file=<%=template%>down.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >下载内容页</td>
                  <td  align="center" ><%=template%>showdown.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>showdown.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >人才列表页</td>
                  <td  align="center" ><%=template%>job.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>job.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >人才内容页</td>
                  <td  align="center" ><%=template%>showjob.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=../<%=showjob%>/index.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >提交简历页</td>
                  <td  align="center" ><%=template%>tjjl.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>tjjl.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >发表留言页</td>
                  <td  align="center" ><%=template%>guest.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>guest.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >显示留言页</td>
                  <td  align="center" ><%=template%>showguest.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>showguest.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >营销网络地图</td>
                  <td  align="center" ><%=template%>yxwlmap.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>yxwlmap.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >营销网络列表页</td>
                  <td  align="center" ><%=template%>yxwl.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>yxwl.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >营销网络详细信息页面</td>
                  <td  align="center" ><%=template%>yxwlxx.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>yxwlxx.asp">修改</a></td>
                </tr>
				
			
				 <tr   >
				   <td align="center" >支付汇款页面</td>
				   <td  align="center" ><%=template%>member/blank.asp</td>
				   <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/blank.asp">修改</a></td>
			    </tr>
				 <tr   >
                  <td align="center" >会员注册页面</td>
                  <td  align="center" ><%=template%>member/memberregister.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/memberregister.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >会员登录页面</td>
                  <td  align="center" ><%=template%>member/memberuserlogin.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/memberuserlogin.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >会员取回密码页面</td>
                  <td  align="center" ><%=template%>member/membergetpass.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/membergetpass.asp">修改</a></td>
                </tr>
                	 <tr   >
                  <td align="center" >会员中心</td>
                  <td  align="center" ><%=template%>member/membercenter.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/membercenter.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >会员中心个人资料</td>
                  <td  align="center" ><%=template%>member/memberinfo.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/memberinfo.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >会员中心我的留言</td>
                  <td  align="center" ><%=template%>member/memberguest.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/memberguest.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >会员中心我的应聘</td>
                  <td  align="center" ><%=template%>member/memberjob.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/memberjob.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >会员中心我的订单</td>
                  <td  align="center" ><%=template%>member/memberorder.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>member/memberorder.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >站内搜索结果页面</td>
                  <td  align="center" ><%=template%>search.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>search.asp">修改</a></td>
                </tr>
					 <tr   >
                  <td align="center" >订单搜索结果列表页面</td>
                  <td  align="center" ><%=template%>orderso.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>orderso.asp">修改</a></td>
                </tr>
			 
					 <tr   >
                  <td align="center" >css样式</td>
                  <td  align="center" ><%=template%>css/style.css</td>
                  <td colspan="-2" align="center" style="border-right:none;"><a href="mbedit.asp?file=<%=template%>css/style.css">修改</a></td>
                </tr>
				
					 <tr   >
					   <td align="center" >调查结果显示页面</td>
					   <td  align="center" ><%=template%>seevote.asp</td>
					   <td colspan="-2" align="center" style="border-right:none;"><a href="mbedit.asp?file=<%=template%>seevote.asp">修改</a></td>
			    </tr>
					 <tr   >
                  <td align="center" style="border-bottom:none">关闭跳转页面</td>
                  <td  align="center" style="border-bottom:none"><%=template%>error.asp</td>
                  <td colspan="-2" align="center" style="border-right:none;border-bottom:none"><a href="mbedit.asp?file=<%=template%>error.asp">修改</a></td>
                </tr>
  </table>
<%end if%>


            <%if t=2 then%>
              <table width="99%" border="0" cellpadding="0" cellspacing="0" align="center" class="dtable2">
        


                <tr class="wz" >
                  <td  >文件</td>
                  <td >文件名称</td>
                  <td >操作</td>
                </tr>
                
                <tr  >
                  <td align="center" >公用顶部</td>
                  <td  align="center" ><%=template%>wap/top.asp</td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/top.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >公用底部</td>
                  <td  align="center" ><%=template%>wap/bottom.asp</td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/bottom.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >首页</td>
                  <td  align="center" ><%=template%>wap/default.asp</td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/default.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >关于我们</td>
                  <td  align="center" ><%=template%>wap/about.asp</td>
                  <td  align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/about.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >新闻列表页</td>
                  <td  align="center" ><%=template%>wap/info.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/info.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >新闻内容页</td>
                  <td  align="center" ><%=template%>wap/showinfo.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/showinfo.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >产品列表页</td>
                  <td  align="center" ><%=template%>wap/product.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/product.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >产品内容页</td>
                  <td  align="center" ><%=template%>wap/showproduct.asp</td>
                  <td colspan="-2" align="center"style="border-right:none" ><a href="mbedit.asp?file=<%=template%>wap/showproduct.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >图片集列表页</td>
                  <td  align="center" ><%=template%>wap/photo.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/photo.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >图片集内容页</td>
                  <td  align="center" ><%=template%>wap/showphoto.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/showphoto.asp">修改</a></td>
                </tr>
                    <tr   >
                  <td align="center" >视频列表页</td>
                  <td  align="center" ><%=template%>wap/video.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/video.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >视频内容页</td>
                  <td  align="center" ><%=template%>wap/showvideo.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/showvideo.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >下载列表页</td>
                  <td  align="center" ><%=template%>wap/down.asp</td>
                  <td colspan="-2" align="center"style="border-right:none" ><a href="mbedit.asp?file=<%=template%>wap/down.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >下载内容页</td>
                  <td  align="center" ><%=template%>wap/showdown.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/showdown.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >人才列表页</td>
                  <td  align="center" ><%=template%>wap/job.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/job.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >人才内容页</td>
                  <td  align="center" ><%=template%>wap/showjob.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/showjob.asp">修改</a></td>
                </tr>
				 <tr   >
                  <td align="center" >提交简历页</td>
                  <td  align="center" ><%=template%>wap/tjjl.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/tjjl.asp">修改</a></td>
                </tr>
                <tr   >
                  <td align="center" >发表留言页</td>
                  <td  align="center" ><%=template%>wap/guest.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/guest.asp">修改</a></td>
                </tr>
               
				 <tr   >
                  <td align="center" >站内搜索结果页面</td>
                  <td  align="center" ><%=template%>wap/search.asp</td>
                  <td colspan="-2" align="center" style="border-right:none"><a href="mbedit.asp?file=<%=template%>wap/search.asp">修改</a></td>
                </tr>
					
  </table>
<%end if%>
</form>