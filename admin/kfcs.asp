<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
if Instr(session("AdminPurview"),"|110,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%
If request.Form("submit") = "确认修改" Then
id=Trim(request.form("id")) 
kfstyle=Trim(request.form("kfstyle")) 
qqstyle=Trim(request.form("qqstyle"))
wwys=Trim(request.form("wwys"))
top=Trim(request.form("top"))
wz=Trim(request.form("wz"))
zk=Trim(request.form("zk"))
fgx=Trim(request.form("fgx"))
db=Trim(request.form("db"))
gzks=Trim(request.form("gzks"))
gzjs=Trim(request.form("gzjs"))
zxsjks=Trim(request.form("zxsjks"))
zxsjjs=Trim(request.form("zxsjjs"))
fwdh=Trim(request.form("fwdh"))
sfzk=Trim(request.form("sfzk"))
jl=Trim(request.form("jl"))
ifewm=Trim(request.form("ifewm"))
ewmsm=Trim(request.form("ewmsm"))
xzewm=Trim(request.form("xzewm"))
Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from kfcs "
    rs.Open sql, conn, 1, 3
rs("kfstyle") = kfstyle
rs("qqstyle") = qqstyle
rs("wwys") = wwys
rs("top") = top
rs("wz") = wz
rs("zk") = zk
rs("fgx") = fgx
rs("db") = db
rs("gzks") = gzks
rs("gzjs") = gzjs
rs("zxsjks") = zxsjks
rs("zxsjjs") = zxsjjs
rs("fwdh") = fwdh
rs("sfzk") = sfzk
rs("jl") = jl
rs("ifewm") = ifewm
rs("ewmsm") = ewmsm
rs("xzewm") = xzewm
	rs.update
    rs.Close
    Set rs = Nothing

call addlog("修改客服参数")
  
    response.Write "<script language=javascript>alert(""恭喜您修改成功！"");window.location.href=""kfcs.asp""</script>"
    response.End
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>
<link href="css/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="js/jquery-1.7.2.min.js"></script>
<style type="text/css">
table{ font-size: 14px;  border-right:none}
table td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
table td a{color:#ffffff}
</style>
</head>

<body>
<form  name="mx" method="post" >
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from kfcs"
rs.Open sql, conn, 1, 1
If rs.EOF Then
    response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
    response.End
End If
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" style=" background: #F9f9f9;border:none; ">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">客服参数设置 </div></td>
</tr>
</table>

<table border="0" align="center" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>

<table width="99%"  border="0" cellpadding="0" cellspacing="0"   align="center" class="dtable4">
                                      
                                      

<tr >
  <td width="19%"  >客服风格选择：</td>
  
  <td width="81%">
                    
                      <select   name="kfstyle" id="kfstyle">
                        <option value="gray" <%if rs("kfstyle")="gray" then Response.Write("selected")%>>灰色</option>
                        <option value="yellow" <%if rs("kfstyle")="yellow" then Response.Write("selected")%>>黄色</option>
                        <option value="blue"  <%if rs("kfstyle")="blue" then Response.Write("selected")%>>蓝色</option>
                        <option value="green" <%if rs("kfstyle")="green" then Response.Write("selected")%>>绿色</option>
                        <option value="orange" <%if rs("kfstyle")="orange" then Response.Write("selected")%>>橙色</option>
                        <option value="white" <%if rs("kfstyle")="white" then Response.Write("selected")%>>白色</option>
                        <option value="syellow" <%if rs("kfstyle")="syellow" then Response.Write("selected")%>>深黄</option>
                        <option value="dblue" <%if rs("kfstyle")="dblue" then Response.Write("selected")%>>淡蓝</option>
                        <option value="sblue" <%if rs("kfstyle")="sblue" then Response.Write("selected")%>>深蓝</option>
                        <option value="black" <%if rs("kfstyle")="black" then Response.Write("selected")%>>黑色</option>
                        <option value="red" <%if rs("kfstyle")="red" then Response.Write("selected")%>>红色</option>
                        <option value="tblue" <%if rs("kfstyle")="tblue" then Response.Write("selected")%>>天空蓝</option>
      </select>                 </td>
</tr>
<tr >
                                        <td  >QQ样式：</td>
  
  <td>
                      <select  name="qqstyle" id="qqstyle">
                        <option value="41" <%if rs("qqstyle")="41" then Response.Write("selected")%>>样式一</option>
                        <option value="42" <%if rs("qqstyle")="42" then Response.Write("selected")%>>样式二</option>
                        <option value="44" <%if rs("qqstyle")="44" then Response.Write("selected")%>>样式三</option>
                        <option value="45" <%if rs("qqstyle")="45" then Response.Write("selected")%>>样式四</option>
                        <option value="46" <%if rs("qqstyle")="46" then Response.Write("selected")%>>样式五</option>
                        <option value="47" <%if rs("qqstyle")="47" then Response.Write("selected")%>>样式六</option>
                      </select>                   </td>
                                      </tr>
                                      

<tr >
  <td  >旺旺样式：</td>
  
  <td>
                      <select  name="wwys" id="wwys">
                        <option value="1" <%if rs("wwys")="1" then Response.Write("selected")%>>样式一</option>
                        <option value="2" <%if rs("wwys")="2" then Response.Write("selected")%>>样式二</option>
                      </select>   </td>
</tr>
<tr >
  <td  >离顶部距离：</td>
  
  <td>
                      <input name="top" type="text" id="top" value="<%=rs("top")%>" size="10" />
                    px</td>
</tr>

<tr >
  <td  >位置：</td>
  
  <td>
                      <select  name="wz" id="wz">
                        <option value="left" <%if rs("wz")="left" then Response.Write("selected")%>>居左</option>
                        <option value="right" <%if rs("wz")="right" then Response.Write("selected")%>>居右</option>
                      </select>                    </td>
</tr>
<tr >
<td  >是否显示手机或者微信二维码：</td>
  
  <td>
<select  name="ifewm" id="ifewm">
<option value="true" <%if rs("ifewm")="true" then Response.Write("selected")%>>显示</option>
<option value="false" <%if rs("ifewm")="false" then Response.Write("selected")%>>不显示</option>
</select></td>
</tr>
<tr >
  <td  >二维码选择类型：</td>
  
  <td>
<select  name="xzewm" id="xzewm">
<option value="true" <%if rs("xzewm")="true" then Response.Write("selected")%>>手机二维码</option>
<option value="false" <%if rs("xzewm")="false" then Response.Write("selected")%>>微信二维码</option>
</select></td>
</tr>
<tr >
  <td  >二维码说明：</td>
  
  <td>
    <input name="ewmsm" type="text" id="ewmsm" value="<%=rs("ewmsm")%>" size="30" />    </td>
</tr>




<tr >
  <td  colspan="2">提示:前6种样式参数开始</td>
</tr>
<tr >
  <td  >是否展开：</td>
  
  <td>
                      <select  name="zk" id="zk">
                        <option value="false" <%if rs("zk")="false" then Response.Write("selected")%>>展开</option>
                        <option value="true" <%if rs("zk")="true" then Response.Write("selected")%>>收缩</option>
                      </select>
                    </td>
</tr>
<tr >
  <td  >是否显示分割线：</td>
  
  <td>
                      <select  name="fgx" id="fgx">
                        <option value="true" <%if rs("fgx")="true" then Response.Write("selected")%>>显示</option>
                        <option value="false" <%if rs("fgx")="false" then Response.Write("selected")%>>不显示</option>
                      </select></td>
</tr>


<tr >
  <td colspan="2">小提示：后6种样式参数开始</td>
</tr>
<tr >
  <td  >工作时间:</td>
  
  <td>

    <select name="gzks" id="gzks">
      <option value="周一" selected="selected" <%if rs("gzks")="周一" then Response.Write("selected")%>>周一</option>
      <option value="周二" <%if rs("gzks")="周二" then Response.Write("selected")%>>周二</option>
      <option value="周三" <%if rs("gzks")="周三" then Response.Write("selected")%>>周三</option>
      <option value="周四" <%if rs("gzks")="周四" then Response.Write("selected")%>>周四</option>
      <option value="周五" <%if rs("gzks")="周五" then Response.Write("selected")%>>周五</option>
      <option value="周六" <%if rs("gzks")="周六" then Response.Write("selected")%>>周六</option>
      <option value="周日" <%if rs("gzks")="周日" then Response.Write("selected")%>>周日</option>
    </select>
    至
   <select name="gzjs" id="gzjs">
      <option value="周一" selected="selected" <%if rs("gzjs")="周一" then Response.Write("selected")%>>周一</option>
      <option value="周二" <%if rs("gzjs")="周二" then Response.Write("selected")%>>周二</option>
      <option value="周三" <%if rs("gzjs")="周三" then Response.Write("selected")%>>周三</option>
      <option value="周四" <%if rs("gzjs")="周四" then Response.Write("selected")%>>周四</option>
      <option value="周五" <%if rs("gzjs")="周五" then Response.Write("selected")%>>周五</option>
      <option value="周六" <%if rs("gzjs")="周六" then Response.Write("selected")%>>周六</option>
      <option value="周日" <%if rs("gzjs")="周日" then Response.Write("selected")%>>周日</option>
    </select>    </td>
</tr>
<tr >
  <td  >在线时间：</td>
  
  <td>
  
    <input name="zxsjks" type="text" id="zxsjks" value="<%=rs("zxsjks")%>" size="10" />
    至
    <input name="zxsjjs" type="text" id="zxsjjs" value="<%=rs("zxsjjs")%>" size="10" /></td>
</tr>
<tr >
  <td  >服务电话：</td>
  
  <td>
    
    <input name="fwdh" type="text" id="fwdh" value="<%=rs("fwdh")%>" />
    </td>
</tr>
<tr >
  <td  >是否展开：</td>
  
  <td>

    <select name="sfzk" id="sfzk">
      <option value="1" <%if rs("sfzk")="1" then Response.Write("selected")%>>展开</option>
      <option value="0" <%if rs("sfzk")="0" then Response.Write("selected")%>>收缩</option>
    </select> </td>
</tr>
<tr >
  <td  >距离左边或者右边的距离：</td>
  
  <td>
    
    <input name="jl" type="text" id="jl" value="<%=rs("jl")%>" size="10" /> px
    </td>
</tr>
<tr >
                                        <td >
  <input  class="bnt"type="submit" value="确认修改" id="submit4" name="submit" />                                        </td>
                                      </tr>
   </table>

</form>
</body>
</html>
<%  conn.Close
    Set conn = Nothing%>