<link href="<%=template%>css/system.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="<%=SitePath%>js/jquery.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/jquery-migrate-1.1.0.min.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/myfocus-2.0.4.min.js"></script><!--幻灯调用-->
<script type="text/javascript" src="<%=SitePath%>js/jquery.fixed.1.3.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/laydate/laydate.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/jquery.lightbox-0.5.js" ></script>
<script type="text/javascript" src="<%=SitePath%>js/qp.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/jquery.SuperSlide.2.1.1.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/qq.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/js.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/Functions.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/jwplayer5.7.min.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/jquery.idTabs.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/index.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/shoppingcart.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/Validform_v5.3.2_min.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/passwordStrength-min.js"></script>
<script type="text/javascript" src="<%=SitePath%>js/laypage.js" ></script>
<script type="text/javascript" src="<%=SitePath%>js/layer.js"></script>
<script type="text/javascript">
window.onerror=function(){return true;}
</script>
<%if kq=1 then%>
<script src="http://siteapp.baidu.com/static/webappservice/uaredirect.js" type="text/javascript"></script>
<%end if%>
<link href="<%=SitePath%>images/qq.css" rel="stylesheet" type="text/css" >
<script type="text/javascript">
$(function(){
$('#gallery a').lightBox();
});
</script> 
<script type="text/javascript">
$(function(){
$(".checkform").Validform({
tiptype:2,
usePlugin:{
passwordstrength:{
minLen:6,
maxLen:18
}
}
});

$.Tipmsg.r="<%=word(362)%>";
$.Tipmsg.c="<%=word(648)%>";

})
</script>
<%=messageskin%>




<%call zmqylanyy%>





<script type="text/javascript">
$(document).ready(function(){
if (window.location.hash!="")
{$(".fade").idTabs(window.location.hash);}
else
{$(".fade").idTabs();}
});
function RefreshImage(valImageId) {
var objImage = document.images[valImageId];
if (objImage == undefined) {
return;
}
var now = new Date();
objImage.src = objImage.src.split('?')[0] + '?x=' + now.toUTCString();
}
</script> 
<script>
//加载公共
$(document).ready(function(){ 
setaccount();
}); 
//加载公共
$(document).ready(function(){ 
showcount();
}); 
//用户状态
function setaccount()
{
$.ajax({
type:"get",
url:"<%=template%>ajax/gmember.asp",
data:"r="+Math.round(Math.random()*10000),
beforesend:function(){$("#fldaccount").html("载入中...");},
error:function(){$("#fldaccount").html("获取信息失败");},
success:function(data){$("#fldaccount").html(data);}
});
}
</script>
<script>
//---建立xmlhttp对象
var xmlhttp;
try{
xmlhttp= new ActiveXObject('Msxml2.XMLHTTP');// IE 
}catch(e){
try{
xmlhttp= new ActiveXObject('Microsoft.XMLHTTP');// IE 
}catch(e){
try{
xmlhttp= new XMLHttpRequest();//Mozilla FF
}catch(e){alert("纳尼?浏览器不支持AJAX!")}
}
}
</script>
<script language="javascript" type="text/javascript">
function showDiv_addcart(){				
layer.open({
    type: 2,
    title: '<%=word(518)%>',
    area: ['500px','170px'],
    shade: 0.8,
    closeBtn:1,
    shadeClose: true,
    content: '<%=template%>ajax/cart_result.asp',
	skin: 'layui-layer-lan',
	btn: ['<%=word(519)%>', '<%=word(520)%>'],
	yes: function(index, layero){
       location.href="<%=template%>member/productbuy.asp?step=1&m=<%=m1%>&Language=<%=Language%>";
    }
});
}
</script>
<script> 
function showcount() 
//购物车及时统计
{
url="<%=template%>ajax/count.asp";
url+=((url.indexOf("?")==-1)?"?":"&")+"rnd="+Math.random();
xmlHttp=GetXmlHttpObject(stateChanged1) 
xmlHttp.open("GET", url , true) 
xmlHttp.send(null) 
}
function stateChanged1() 
{ 
if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete") 
{ 
document.getElementById("txtHint").innerHTML=xmlHttp.responseText 
} 
}
function GetXmlHttpObject(handler) 
{ 
var objXmlHttp=null
if (navigator.userAgent.indexOf("Opera")>=0) 
{ 
alert("This example doesn't work in Opera") 
return; 
} 
if (navigator.userAgent.indexOf("MSIE")>=0) 
{ 
var strName="Msxml2.XMLHTTP"
if (navigator.appVersion.indexOf("MSIE 5.5")>=0) 
{ 
strName="Microsoft.XMLHTTP" 
} 
try 
{ 
objXmlHttp=new ActiveXObject(strName) 
objXmlHttp.onreadystatechange=handler 
return objXmlHttp 
} 
catch(e) 
{ 
alert("Error. Scripting for ActiveX might be disabled") 
return 
} 
} 
if (navigator.userAgent.indexOf("Mozilla")>=0) 
{ 
objXmlHttp=new XMLHttpRequest() 
objXmlHttp.onload=handler 
objXmlHttp.onerror=handler 
return objXmlHttp 
} 
} 
</script> 
<script>
//添加到购物车
function add_2_my_cart(){	
list_hid=document.getElementsByName("hidv_0001_<%=proid%>");
var hasNull = false;
for(var j=0;j<list_hid.length;j++){
if(!list_hid[j].value) 
{
hasNull = true; 
}
}
if(hasNull==true) 
{
layer.msg('<%=word(369)%>', {icon: 5}); 
return false; 
}
var id=document.getElementById('id').value;
var quantity=document.getElementById('quantity').value;//数量
if(quantity.length>0){//不为空时
strURL="<%=template%>member/productbuy.asp?language=<%=language%>&quantity="+ quantity + "&id="+ id;
strURL+=((strURL.indexOf("?")==-1)?"?":"&")+"rnd="+Math.random();
xmlhttp.open("GET", strURL, true);
xmlhttp.onreadystatechange = function(){
if(xmlhttp.readyState == 4){

if(xmlhttp.status == 200){
if(xmlhttp.responseText!=""){
var data=escape(xmlhttp.responseText);
add_2_my_cart_result(data);
showcount();
}
}
else{
}
}
else{
}
}
xmlhttp.setRequestHeader('Content-type','application/x-www-form-urlencoded');
xmlhttp.send(null);//使用GET后这里就用Null
}
else{
document.getElementById('show_add_to_cart2').innerHTML="<img src=../image/gantanhao.gif  style=margin-top:3px; margin-right:2px;/>请不要重复操作。";
}
}
function add_2_my_cart_result(data){
var resultbox=document.getElementById('addcart_result');
if(data==1){
showDiv_addcart()
}
else{
if(data==3){
show_err3()
}
else if(data==2){
showDiv_addcart()
}
else {
showDiv_addcart()
}
}
}
</script>
<script> 
// 加载购物车按钮状态
function showcart() 
{
var id=document.getElementById('id').value;
url="<%=template%>ajax/showcart.asp?language=<%=language%>&id="+ id;
url+=((url.indexOf("?")==-1)?"?":"&")+"rnd="+Math.random();
xmlHttp=GetXmlHttpObject(stateChanged) 
xmlHttp.open("GET", url , true) 
xmlHttp.send(null) 
}
function stateChanged() 
{ 
if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete") 
{ 

document.getElementById("carttip").innerHTML=xmlHttp.responseText 

} 

}
function GetXmlHttpObject(handler) 
{ 
var objXmlHttp=null
if (navigator.userAgent.indexOf("Opera")>=0) 
{ 
alert("This example doesn't work in Opera") 
return; 
} 
if (navigator.userAgent.indexOf("MSIE")>=0) 
{ 
var strName="Msxml2.XMLHTTP"
if (navigator.appVersion.indexOf("MSIE 5.5")>=0) 
{ 
strName="Microsoft.XMLHTTP" 
} 
try 
{ 
objXmlHttp=new ActiveXObject(strName) 
objXmlHttp.onreadystatechange=handler 
return objXmlHttp 
} 
catch(e) 
{ 
alert("Error. Scripting for ActiveX might be disabled") 
return 
} 
} 
if (navigator.userAgent.indexOf("Mozilla")>=0) 
{ 
objXmlHttp=new XMLHttpRequest() 
objXmlHttp.onload=handler 
objXmlHttp.onerror=handler 
return objXmlHttp 
} 
} 
</script> 
<%if kq=1 then%>
<script type="text/javascript">uaredirect("<%=wz%><%=sitepath%>wap/?language=<%=language%>");</script>
<%end if%>
<script type="text/javascript">
$(function() {
	setTimeout("changeDivStyle();", 100); // 0.1秒后展示结果，因为HTML加载顺序，先加载脚本+样式，再加载body内容。所以加个延时
});
function changeDivStyle(){
var o_status = $("#o_status").val();	//获取隐藏框值
if(o_status==0){
$('#create').css('background', '#DD0000');
$('#createText').css('color', '#DD0000');
}else if(o_status==1){
$('#check').css('background', '#DD0000');
$('#checkText').css('color', '#DD0000');
}else if(o_status==2){
$('#produce').css('background', '#DD0000');
$('#produceText').css('color', '#DD0000');
}else if(o_status==3){
$('#delivery').css('background', '#DD0000');
$('#deliveryText').css('color', '#DD0000');
}else if(o_status>=4){
$('#received').css('background', '#DD0000');
$('#receivedText').css('color', '#DD0000');
}
}
</script>


<script>
function webRequest(url,xmlstr,method,isAsync,callback,callbackargs)
{
    var xmlhttp = null;
	if (window.ActiveXObject) {
       xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
    }else if (window.XMLHttpRequest) {
       xmlhttp = new XMLHttpRequest();
    }
	if(!method) method = "GET";
	if(!isAsync) isAsync = false;
	if(isAsync && callback){
		var args = new Array();
		args.push(xmlhttp);
		if(typeof(callbackargs) == 'object'){
			args = args.concat(callbackargs);
		}
		xmlhttp.onreadystatechange = function(){callback.apply(null,args);}
	}
	xmlhttp.open(method,url,isAsync);
    if(method.toLowerCase() == 'post'){
		xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	}
    xmlhttp.send(xmlstr);
	if(!isAsync){
		return xmlhttp.responseText;
	}
}

function select_sp(id,i,ib,values,price,listcode,nostockorder){
	var selectover = false;
	
	var result = document.getElementById("result_"+listcode+"_"+id);
	list=document.getElementsByName("sv_n_"+id+"_"+i);
	for(var j=0;j<list.length;j++){
		if (list[j].className == "spva-img"||list[j].className == "spva-imgon"){list[j].className="spva-img";}else{list[j].className="spva";}
		//list[j].style.background="#fff";
	}
	if (document.getElementById("sv_"+id+"_"+i+"_"+ib).className=="spva-img"||document.getElementById("sv_"+id+"_"+i+"_"+ib).className=="spva-imgon"){document.getElementById("sv_"+id+"_"+i+"_"+ib).className="spva-imgon";}else{document.getElementById("sv_"+id+"_"+i+"_"+ib).className="spvaon";}
	document.getElementById("hidv_"+listcode+"_"+id+"_"+i).value=values;

	list_hid=document.getElementsByName("hidv_"+listcode+"_"+id);
	var spvalue="";
	for(var j=0;j<list_hid.length;j++){
		if (list_hid[j].value != "")
		{
			spvalue+=list_hid[j].value.split("|")[1];
			if (j<list_hid.length-1)
			{
				spvalue+=",";
			}
		}
	}
	for(var j=0;j<list_hid.length;j++){
		if (!spvalue.split(",")[j])
		{
			selectover = true
		}
	}
	if (selectover==false)
	{    
	
	
	document.getElementById('myDiv').style.display='none'
		 var url="<%=template%>ajax/getsinfo.asp?action=getspinfo&language=<%=language%>&id="+id+"&name="+escape(spvalue)+"&price="+price+"&t="+Math.random();
		 webRequest(url, "", "GET", true, function() {
			var xmlhttp = arguments[0];
			var args = new Array();
			if (xmlhttp.readyState == 4) {
				var returntxt = xmlhttp.responseText;
				var stock = returntxt.split("|")[0];
				var price = returntxt.split("|")[1];
				var sprice = returntxt.split("|")[2];
				var itemnum = returntxt.split("|")[3];
				var mes = returntxt.split("|")[4];
				
				if(mes.length > 0)
				{
					result.style.display = 'block';
					result.innerHTML = mes;
				
					showcart();
				}
				else
				{
					
					result.style.display = 'none';
				
					
					result.innerHTML = "";
				}
				document.getElementById("p-sprice").innerHTML = sprice;
				document.getElementById("p-price").innerHTML = price;
				document.getElementById("p-itemnum").innerHTML = itemnum;
				
				if (stock < 1)
				{
					if (nostockorder == 0)
					{
						document.getElementById("button_"+listcode+"_"+id).disabled = true;
					}else{
						document.getElementById("button_"+listcode+"_"+id).disabled = false;
					}
				}else{
						
				}
					document.getElementById('kkk').style.display='none';
			
			}
		 }, "");
	}
}


	</script>







