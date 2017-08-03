<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->

<%
if Instr(session("AdminPurview"),"|91,")=0 then 
response.Write "<script language=javascript>alert(""非常抱歉！您没有相对应的模块的权限！"");window.location.href=""right.asp""</script>"
  response.end
end if
'========判断是否具有管理权限
%>
<%  

zxys=request.Form("zxys")
If request.Form("submit") = "保存" Then
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"' "
rs.Open sql, conn, 1, 3
rs("zxys")=zxys
    rs.update
    rs.Close
    Set rs = Nothing
    call addlog("浮动客户咨询样式选择")
  response.Write "<script language=javascript>alert(""恭喜您修改成功"");window.location.href=""messageskin.asp""</script>"
    response.End
End If
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>zmcms_管理页面</title>

<script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>

<link href="css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.table1{ font-size: 14px; background: #F9f9f9;border:none;}
.table1 td{padding:10px; font-family:"微软雅黑";color:#ccc; font-weight:bold;}
.table1 td a{color:#ffffff}
</style>
</head>


<body>

<form method="POST" action="messageskin.asp" id="form" name="form">
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from wzset where lang='"&lang&"'"
rs.Open sql, conn, 1, 1
If rs.EOF Then
response.Write "<script language=javascript>alert(""不存在此数据！"");window.history.back();</script>"
response.End
End If
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0"  align="center" class="table1">
<tr>
 <td  style="border-bottom:none;border-right:none;color:#cccccc"><div align="center">及时咨询样式设置</div></td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="99%" class="dtable3">
 <tr>
 <td ><div class="dmeun"><a href="set.asp" target="main">返回菜单</a></div></td>
 

 </tr>
</table>

<table width="99%" border="0" cellpadding="0" cellspacing="0" class="dtable3">

  
  <tr><td >
  
  
  <table cellpadding="0" cellspacing="0" border="0" >
          <tr>
		  
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin01.png" alt="样式一" width="80" /></td>
                </tr>
                <tr>
                  <td >样式一
                    <input type="radio" name="zxys" value="1" <%
If rs("zxys")="1" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
			  
            <td ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin02.png" alt="样式二" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二
                    <input type="radio" name="zxys" value="2" <%
If rs("zxys")="2" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin03.png" alt="样式三" width="80" /></td>
                </tr>
                <tr>
                  <td >样式三
                    <input type="radio" name="zxys" value="3"<%
If rs("zxys")="3" Then Response.Write("checked")%> ></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td >
			<table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin04.png" alt="样式四" width="80" /></td>
                </tr>
                <tr>
                  <td >样式四
                    <input type="radio" name="zxys" value="4" <%
If rs("zxys")="4" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
        
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin05.png" alt="样式五" width="80" /></td>
                </tr>
                <tr>
                  <td >样式五
                    <input type="radio" name="zxys" value="5" <%
If rs("zxys")="5" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin06.png" alt="样式六" width="80" /></td>
                </tr>
                <tr>
                  <td >样式六
                    <input type="radio" name="zxys" value="6" <%
If rs("zxys")="6" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin07.png" alt="样式七" width="80" /></td>
                </tr>
                <tr>
                  <td >样式七
                    <input type="radio" name="zxys" value="7" <%
If rs("zxys")="7" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin08.png" alt="样式八" width="80" /></td>
                </tr>
                <tr>
                  <td >样式八
                    <input type="radio" name="zxys" value="8" <%
If rs("zxys")="8" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
          </tr>
        </table>
<table cellpadding="0" cellspacing="0" border="0" >
          <tr>
		  
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin09.png" alt="样式九" width="80" /></td>
                </tr>
                <tr>
                  <td >样式九
                    <input type="radio" name="zxys" value="9" <%
If rs("zxys")="9" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
			  
            <td ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin10.png" alt="样式十" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十
                    <input type="radio" name="zxys" value="10" <%
If rs("zxys")="10" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin11.png" alt="样式十一" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十一
                    <input type="radio" name="zxys" value="11"<%
If rs("zxys")="11" Then Response.Write("checked")%> ></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td >
			<table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin12.png" alt="样式十二" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十二
                    <input type="radio" name="zxys" value="12" <%
If rs("zxys")="12" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
        
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin13.png" alt="样式十三" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十三
                    <input type="radio" name="zxys" value="13" <%
If rs("zxys")="13" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin14.png" alt="样式十四" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十四
                    <input type="radio" name="zxys" value="14" <%
If rs("zxys")="14" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin15.png" alt="样式十五" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十五
                    <input type="radio" name="zxys" value="15" <%
If rs("zxys")="15" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin16.png" alt="样式十六" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十六
                    <input type="radio" name="zxys" value="16" <%
If rs("zxys")="16" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
          </tr>
        </table>
		
		<table cellpadding="0" cellspacing="0" border="0" >
          <tr>
		  
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin17.png" alt="样式十七" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十七
                    <input type="radio" name="zxys" value="17" <%
If rs("zxys")="17" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
			  
            <td ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin18.png" alt="样式十八" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十八
                    <input type="radio" name="zxys" value="18" <%
If rs("zxys")="18" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin19.png" alt="样式十九" width="80" /></td>
                </tr>
                <tr>
                  <td >样式十九
                    <input type="radio" name="zxys" value="19"<%
If rs("zxys")="19" Then Response.Write("checked")%> ></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td >
			<table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin20.png" alt="样式二十" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二十
                    <input type="radio" name="zxys" value="20" <%
If rs("zxys")="20" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
        
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin21.png" alt="样式二一" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二一
                    <input type="radio" name="zxys" value="21" <%
If rs("zxys")="21" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin22.png" alt="样式二二" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二二
                    <input type="radio" name="zxys" value="22" <%
If rs("zxys")="22" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin23.png" alt="样式二三" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二三
                    <input type="radio" name="zxys" value="23" <%
If rs("zxys")="23" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin24.png" alt="样式二四" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二四
                    <input type="radio" name="zxys" value="24" <%
If rs("zxys")="24" Then Response.Write("checked")%>></td>
                </tr>
              </table></td>
            <td ></td>
          </tr>
        </table>
		
		<table cellpadding="0" cellspacing="0" border="0" >
          <tr>
		  
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin25.png" alt="样式二五" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二五
                    <input type="radio" name="zxys" value="25" <%
If rs("zxys")="25" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
			  
            <td  ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin26.png" alt="样式二六" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二六
                    <input type="radio" name="zxys" value="26" <%
If rs("zxys")="26" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td ><table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin27.png" alt="样式二七" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二七
                    <input type="radio" name="zxys" value="27"<%
If rs("zxys")="27" Then Response.Write("checked")%> ></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
            <td >
			<table cellpadding="3" cellspacing="0">
                <tr>
                  <td ><img src="Images/zxys/Skin28.png" alt="样式二八" width="80" /></td>
                </tr>
                <tr>
                  <td >样式二八
                    <input type="radio" name="zxys" value="28" <%
If rs("zxys")="28" Then Response.Write("checked")%>></td>
                </tr>
              </table>
		    </td>
            <td ></td>
			
        
            <td >&nbsp;</td>
            <td ></td>
            <td >&nbsp;</td>
            <td ></td>
            <td >&nbsp;</td>
            <td ></td>
            <td >&nbsp;</td>
            <td ></td>
          </tr>
        </table>
  <tr>
    <td style=" border-bottom:none"><label>
      <input type="submit" name="Submit" value="保存" class="bnt" />
    </label></td>
  </tr>
</table>






</td></tr>
</table>

</form>

</body>
</html>

