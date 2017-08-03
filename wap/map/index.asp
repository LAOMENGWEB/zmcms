<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../Inc/library.asp"-->
<!-- #include file="../../Inc/sub.asp"--> 

<%
echo ob_get_contents(""&template&"wap/map.asp")
%>

<script>
	
    
    var map,marker,infoWin;
    var companyName = decodeURIComponent("<%=wapmc%>");
    var phoneNum = decodeURIComponent("<%=mapphone%>");
    var remarks = decodeURIComponent("<%=wapcontent%>");
    var navLink = "http://map.baidu.cn/mobile/#place/linesearch/foo=bar/from=plac&start=&end=";
    var endStr = "word="+companyName+"&point=<%=wapmap%>&uid=jpapVMZI6GkEYSIA23wTjT6U";
    navLink += encodeURIComponent(endStr)+"&tab=line";
    var html = [];
    html.push('<h4 style="padding: 0px;margin: 0px;">单位信息 </h4>');
    html.push('<table border="0" cellpadding="1"  cellspacing="1" class="map">');
    html.push('  <tr>');
    html.push('      <td>名称：</td>');
    html.push('      <td colspan="2">'+companyName+'</td>');
    html.push('  </tr>');
    html.push('  <tr>');
    html.push('      <td  >电话：</td>');
    html.push('      <td colspan="2" >'+phoneNum+'</td>');
    html.push('  </tr>');
    html.push('  <tr>');
    html.push('      <td >备注：</td>');
    html.push('      <td colspan="2">'+remarks+'</td>');
    html.push('  </tr>');
    html.push('</table>');

    html.push("<div style='margin:10px auto;display:block;width:220px;'>");
    html.push("<a href='tel:<%=mapphone%>' style='margin-right:10px;text-align:center;display:inline-block;border-radius: .5em;height: 40px;font-size: 12pt;line-height: 40px;width:100px;");
    html.push('color: #606060;border: solid 1px #b7b7b7;background: #fff;background: -webkit-gradient(linear, left top, left bottom, from(#fff), to(#ededed));background: -moz-linear-gradient(top,  #fff,  #ededed);filter:  progid:DXImageTransform.Microsoft.gradient(startColorstr=\'#ffffff\', endColorstr=\'#ededed\');');
    html.push(">电话</a>");

    html.push("<a href='javascript:void(0)' onclick='daohang();' style='text-align:center;display:inline-block;border-radius: .5em;height: 40px;font-size: 12pt;line-height: 40px;width:100px;");
    html.push('color: #606060;border: solid 1px #b7b7b7;background: #fff;background: -webkit-gradient(linear, left top, left bottom, from(#fff), to(#ededed));background: -moz-linear-gradient(top,  #fff,  #ededed);filter:  progid:DXImageTransform.Microsoft.gradient(startColorstr=\'#ffffff\', endColorstr=\'#ededed\');');
    html.push(">导航</a>");

    html.push("</div>");
    if(window.pageYOffset == 0){
        window.scrollTo(0, window.pageYOffset + 1);
    }
    var container = $("#container");
    container.height($(window).height()+60);
    container.attr("companyName",companyName);
    container.attr("phoneNum",phoneNum);
    container.attr("remarks",remarks);
    map = new BMap.Map("container");
    var pointer = new BMap.Point(<%=wapmap%>);
    container.attr("lng",<%=wapx%>);
    container.attr("lat",<%=wapy%>);

    map.addControl(new BMap.NavigationControl({anchor: BMAP_ANCHOR_TOP_LEFT, type: BMAP_NAVIGATION_CONTROL_ZOOM}));
    map.addControl(new BMap.MapTypeControl({mapTypes: [BMAP_NORMAL_MAP,BMAP_HYBRID_MAP]}));
    infoWin = new BMap.InfoWindow(html.join(""));
    marker = new BMap.Marker(pointer);  // 创建标注
    //marker.setAnimation(BMAP_ANIMATION_BOUNCE); //跳动的动画
    map.addOverlay(marker);         // 将标注添加到地图中
    marker.openInfoWindow(infoWin);// 打开信息窗口
    marker.addEventListener("click",function(){
        this.openInfoWindow(infoWin);
    });
    map.centerAndZoom(pointer, 15);
    map.panTo(pointer);
    

/**** 百度导航 ****/
function daohang(){
var geolocation = new BMap.Geolocation();
geolocation.getCurrentPosition(function(r){
if(this.getStatus() == BMAP_STATUS_SUCCESS){
var driving = new BMap.DrivingRoute(map, {renderOptions:{map: map, autoViewport: true}});
driving.search(r.point,pointer);
}else {
alert("定位失败！")
}
},{enableHighAccuracy: true});
}    
</script>