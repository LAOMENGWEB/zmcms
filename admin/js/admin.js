          


function addUniteRows(v0,v1,v2,v3,v4,v5)
{
	//创建td节点
	var td0= document.createElement("td");
	var td1= document.createElement("td");
	var td2= document.createElement("td");
	var td3= document.createElement("td");
	td3.setAttribute("class","listrow");
	var td4= document.createElement("td");
	td4.setAttribute("align","");

	td0.innerHTML = v1+'<INPUT TYPE="hidden" NAME="uBianhao" value="'+v1+'"> <INPUT TYPE="hidden" id="uPid" NAME="uPid" value="'+v0+'"><INPUT TYPE="hidden" NAME="pWeight" id="pWeight" value="'+v5+'">';
	td1.innerHTML = v2+'<INPUT TYPE="hidden" NAME="uProduct_name" value="'+v2+'">';
	td2.innerHTML = v3+'<INPUT TYPE="hidden" NAME="uSprice" value="'+v3+'">';
	td3.innerHTML = '<input type=text name=uPNum size=2 value="'+v4+'" maxlength=3 class=form onkeyup=value=value.replace(\/[^\\d]\/g,\'\');  onblur=sumSPrice();>';
	td4.innerHTML = '<A onclick="return delUniteRows(event);" style="cursor:pointer;">删除</a>';

	//创建tr节点
	var trNode = document.createElement("tr");
	trNode.appendChild(td0);
	trNode.appendChild(td1);
	trNode.appendChild(td2);
	trNode.appendChild(td3);
	trNode.appendChild(td4);

	//创建tbody节点，必须使用tbody，不然无法使用dom方式添家行，只能用insertRow
	var trBody = document.createElement("tbody");
	trBody.appendChild(trNode);
	document.getElementById("tableUP").appendChild(trBody);
}
function getEvent()
{ 
    if(document.all)   return window.event;   
    func=getEvent.caller;       
    while(func!=null){ 
        var arg0=func.arguments[0];
        if(arg0)
        {
          if((arg0.constructor==Event || arg0.constructor ==MouseEvent) || (typeof(arg0)=="object" && arg0.preventDefault && arg0.stopPropagation))
          { 
          return arg0;
          }
        }
        func=func.caller;
    }
    return null;
}
function delUniteRows(event)
{
    var Evt = getEvent();
    var srcElements = Evt.srcElement == undefined ? Evt.target : Evt.srcElement
    var i = srcElements.parentNode.parentNode.rowIndex;
    document.getElementById("tableUP").deleteRow(i);
	sumSPrice();
	sumWeight();
}
//function delUniteRows(event)
//{
//	if(event == null)
//	{
//		event = window.event; // For IE
//	}
//	var eventObj = event.srcElement? event.srcElement : event.target;   
//	var tbodyNode = eventObj.parentNode.parentNode.parentNode;
//	var trNode = eventObj.parentNode.parentNode
//	var rowIndex = trNode.rowIndex;
//	var objTable = document.getElementById("tableUP");
//
//	objTable.removeChild(tbodyNode);
//	sumSPrice();
//	return false;
//	//sumWeight();
//}

function trim(str) //去空格
{
	return str.replace(/(^\s*)|(\s*$)/g, "");
}
function toUpdateGradeIco()  // WangLei 更新记录排序  2008-3-10
{   
	var i = 0,icoStr="";   
	var grade_ico = document.all.grade_ico;
	for (var i=0; i<grade_ico.length; i++) 
	{
		icoStr+=grade_ico[i].value+",";
	}
	if (icoStr.length == 0 )
	{
		alert("暂无内容可更新！");
		return;
	}else{
		//alert(icoStr);
		location.href="?action=updateGradeIco&icoStr="+icoStr;  
	}
}

function checknum(str){
	var strSource ="0123456789.";
	var ch;
	var i;
	var temp; 
	for (i=0;i<=(str.length-1);i++)
	{

		ch = str.charAt(i);
		temp = strSource.indexOf(ch);
		if (temp==-1) 
		{
		 return 0;
		}
	}
	if (strSource.indexOf(ch)==-1)
	{
		return 0;
	}
	else
	{
		return 1;
	} 
}
function SinnerHTML(pObj,iobj,errMsg){
   iobj.innerHTML = "<img src=\""+ AdminPath +"images/succes.gif\" align=\"absmiddle\"> <font color=\"red\">"+errMsg+"</font>";
}
function EinnerHTML(pObj,iobj,errMsg){
   iobj.innerHTML = "<img src=\""+ AdminPath +"images/error.gif\" align=\"absmiddle\"> <font color=\"red\">"+errMsg+"</font>";
   pObj.focus(); 
}

function tdshow(obj,arr)
{
	var cArr = arr||["#fff","#F0FAFF","#DEEFF8"];
	var trArr=document.getElementById(obj).getElementsByTagName("tr");
	for(var i=0,j=trArr.length;i<j;i++)
	{
		var o=trArr[i];
		o.style.backgroundColor=cArr[i%2];
		o.onmouseover=function(){this.style.backgroundColor=cArr[2];}
		o.onmouseout=function(){this.style.backgroundColor=cArr[this.sectionRowIndex%2];}
	}
}
function resizeEditor(name,size) {   
	var newheight = parseInt(document.getElementById(name +"___Frame").style.height, 10) + size;   
	if(newheight >= 150) {   
		document.getElementById(name +"___Frame").style.height = newheight + 'px';   
	}   
}  
function $FatherTab(obj){return document.getElementById(obj)}
function FatherTab(Cname,Lenght,j){
	for(i=1;i<Lenght;i++){eval("$FatherTab('"+Cname+i+"').style.display='none'");
	eval("$FatherTab('"+Cname+j+"').style.display='block'");}
}
//** Tips
function AddTips(s){
	try{
	clearTimeout(RemoveTipsTimer);
	}catch(e){}
	//if( s == '') return;
	var args = arguments;
	var id,position,width,height;
	var top = document.documentElement.scrollTop ? document.documentElement.scrollTop : document.body.scrollTop;
	var browser=navigator.userAgent;
	if (browser.indexOf("Firefox")>0){
	　　top = document.body.scrollTop;
	}
	if(args[1]) id = args[1];
	if(args[2]) position = args[2]; //显示位置,有 top bottom left right mouse(鼠标跟随) 等多种属性
	width = args[3];
	if(args[4]) height = args[4];
	if(!id) id = "dvTips";
	var dvTips = document.getElementById(id);
	if(!dvTips){
		dvTips = document.createElement("DIV");
		dvTips.id = id;
		dvTips.className = "dvTips";
		dvTips.style.position = 'absolute';
		dvTips.style.padding = '0px';
		if(width) dvTips.style.width = width + "px";
		if(height) dvTips.style.height = height + "px";
		document.body.appendChild(dvTips);
	}
	dvTips.onmouseover = function(){clearTimeout(RemoveTipsTimer);}
	dvTips.onmouseout = function(){RemoveTips(id,true);}
	if(s != '') dvTips.innerHTML = s;
	try{
	mX = event.x ? event.x : event.pageX;
	mY = event.y ? event.y : event.pageY;
	if(position == 'top'){
		dvTips.style.top = getOffsetTop(event.srcElement) - ( dvTips.clientHeight + 5 ) + "px";
		dvTips.style.left = getOffsetLeft(event.srcElement) + "px";
	}else if(position == 'bottom'){
		dvTips.style.top = getOffsetTop(event.srcElement) + ( event.srcElement.clientHeight + 5 ) + "px";
		dvTips.style.left = getOffsetLeft(event.srcElement) + "px";
	}else if(position == 'left'){
		dvTips.style.top = getOffsetTop(event.srcElement) + "px";
		dvTips.style.left = getOffsetLeft(event.srcElement) - ( dvTips.clientWidth + 5 ) + "px";
	}else if(position == 'right'){
		dvTips.style.top = getOffsetTop(event.srcElement) + "px";
		dvTips.style.left = getOffsetLeft(event.srcElement) + ( event.srcElement.clientWidth + 5 ) + "px";
	}else{
		var left = mX + 5 + document.body.scrollLeft;
		if (left+parseInt(width)>document.body.scrollWidth)	left=mX - 5 - width;		
		dvTips.style.left = left + "px";	
		dvTips.style.top = mY + 5 + top + "px";
	}
	dvTips.style.display = '';
	} catch(e){}
}
var RemoveTipsTimer;
function RemoveTips(){
	var args = arguments;
	var id,delay;
	if(args[0]) id = args[0];
	if(typeof(args[1]) != 'undefined') delay = args[1];
	if(!id) id = 'dvTips';
	var dvTips = document.getElementById(id);
	if(dvTips){
		if(typeof(delay) == 'undefined' || delay == true) RemoveTipsTimer = setTimeout(function(){dvTips.style.display = 'none'},200);
		else dvTips.style.display = 'none';
	}
}
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
function ShowMenu(id,url,pid,target)
{
	 var url=AdminPath +"Ajax_menu.asp?id="+id+"&pid="+pid+"&url="+escape(url)+"&target="+target+"&t="+Math.random();
     $.ajax({
     type:"POST",
     url:url,
     data:"",
     success:function(returntxt){
		if(returntxt.length > 0)
		{
			$("#parentmenu").html(returntxt);
		}
		else
		{
			$("#parentmenu").html("Loading………");
		}
     }
   })		
}
function getSelectValue(typeid,value,id,mode,lang)
{
	 if (typeid == 0)
	 {
		 typeid = $("#SelectType").val();
	 }
	 var url="Ajax_data.asp?action=selectvalue&typeid="+typeid+"&value="+escape(value)+"&id="+id+"&mode="+mode+"&lang="+lang+"&t="+Math.random();
     $.ajax({
     type:"POST",
     url:url,
     data:"",
     success:function(returntxt){
		if(returntxt.length > 0)
		{
			 if (mode == 0)
			 {
				 $("#SelectValue").html(returntxt);
				 $("#SelectValueBox").show();
			 }
			 if (mode == 1)
			 {
				 $("#SpecificationValue").html(returntxt);
				 $("#SpecificationValueBox").show();
			 }
			 if (mode == 2)
			 {
				 $("#ProductValue"+lang).html(returntxt);
				 $("#ProductValueBox"+lang).show();
			 }
		}
		else
		{
			 if (mode == 0)
			 {
				 $("#SelectValue").html("");
				 $("#SelectValueBox").hide();
			 }
			 if (mode == 1)
			 {
				 $("#SpecificationValue").html("");
				 $("#SpecificationValueBox").hide();
			 }
			 if (mode == 2)
			 {
				 $("#ProductValue"+lang).html("");
				 $("#ProductValueBox"+lang).hide();
			 }
		}
     }
   })		
}
function getspecification(id,name,list)
{
	 var url="Ajax_data.asp?action=getspecification&id="+id+"&name="+escape(name)+"&list="+list+"&t="+Math.random();
     $.ajax({
     type:"POST",
     url:url,
     data:"",
     success:function(returntxt){
		if(returntxt.length > 0)
		{
			$("#Get_Specification"+id+"").html(returntxt);
		}
		else
		{
			$("#Get_Specification"+id+"").html("");
		}
     }
   })		
}
function getspecificationpic(id,name,pic)
{
	 var url="Ajax_data.asp?action=getspecificationpic&id="+id+"&pic="+pic+"&name="+escape(name)+"&t="+Math.random();
     $.ajax({
     type:"POST",
     url:url,
     data:"",
     success:function(returntxt){
		if(returntxt.length > 0)
		{
			$("#Get_SpecificationPic"+id+"").html(returntxt);
		}
		else
		{
			$("#Get_SpecificationPic"+id+"").html("");
		}
     }
   })		
}
function GetStockTips(title,id,name,list) 
{
	var url="Ajax_data.asp?action=getspecification&id="+id+"&name="+escape(name)+"&list="+list+"&t="+Math.random();
	AddTips('<dl><dt>'+title+'</dt><dd>正在加载信息...</dd></dl>', null, 'mouse', '500', null);
	webRequest(url, "", "GET", true, function() {
		var xmlhttp = arguments[0];
		var args = new Array();
		if (xmlhttp.readyState == 4) {
			var content = xmlhttp.responseText;
			AddTips('<dl><dt>'+title+'</dt><dd>'+content+'</dd></dl>', null, 'mouse', '500', null);
		}
	}, new Array(event));
}
function GetOrderPayLogTips(title,id) 
{
	var url="Ajax_data.asp?action=getorderpaylog&id="+id+"&t="+Math.random();
	AddTips('<dl><dt>'+title+'</dt><dd>正在加载信息...</dd></dl>', null, 'mouse', '500', null);
	webRequest(url, "", "GET", true, function() {
		var xmlhttp = arguments[0];
		var args = new Array();
		if (xmlhttp.readyState == 4) {
			var content = xmlhttp.responseText;
			AddTips('<dl><dt>'+title+'</dt><dd>'+content+'</dd></dl>', null, 'mouse', '500', null);
		}
	}, new Array(event));
}
function search_suggest(key,from,lang) 
{ 
	var str = document.getElementById(key).value;
	var div = document.getElementById("search_suggest_"+ key); 
	var sql = "";
	var suggest;
	if (str !="")
	{
	var url="ajax_data.asp?action=search_suggest&key="+escape(str)+"&lang="+lang;
	if(from.indexOf("|")!= -1){
		url+="&from="+from.split('|')[0];
		url+="&sql="+from.split('|')[1];
	}else{
		url+="&from="+from;
	}
	url+="&t="+Math.random();
    $.ajax({
    type:"GET",
    url:url,
    data:"",
    success:function(returntxt){
		if(returntxt.length > 0)
		{
			div.innerHTML = ""; 
			var str = returntxt.split("|");
			for (var i=0; i<str.length; i++) 
			{ 
				suggest = "<div>";
				keystr = str[i].split("---");
				for (var k=0; k<keystr.length; k++) 
				{ 
					suggest += "<a onclick=javascript:setSearch('"+key+"',this.innerHTML);>" + keystr[k] + "</a>  "; 
				} 
				suggest += "</div>";
				div.innerHTML += suggest; 
				div.style.display = ""; 
//			var suggest = "<div onclick=javascript:setSearch('"+key+"',this.innerHTML);>" + str[i] + "</div>"; 
//			div.innerHTML += suggest; 
//			div.style.display = ""; 
			} 
		}
		else
		{
			div.style.display = "none"; 
		}
     }
    })		
	}
}
function setSearch(key,div_value) 
{ 
	if(div_value.indexOf("---")!= -1){
		div_value = div_value.split('---')[1]
	}
	document.getElementById(key).value = div_value; 
	document.getElementById("search_suggest_"+ key).style.display = "none"; 
} 
function select_sp(id,i,ib,values,price,listcode){
	var selectover = 1;
	list=document.getElementsByName("sv_n_"+id+"_"+i);
	for(var j=0;j<list.length;j++){
		if (list[j].className == "spva-img"||list[j].className == "spva-imgon"){list[j].className="spva-img";}else{list[j].className="spva";}
		list[j].style.background="#fff";
	}
	if (document.getElementById("sv_"+id+"_"+i+"_"+ib).className=="spva-img"||document.getElementById("sv_"+id+"_"+i+"_"+ib).className=="spva-imgon"){document.getElementById("sv_"+id+"_"+i+"_"+ib).className="spva-imgon";}else{document.getElementById("sv_"+id+"_"+i+"_"+ib).className="spvaon";}
	document.getElementById("hidv_"+listcode+"_"+id+"_"+i).value=values;
	document.getElementById("sv_"+id+"_"+i+"_"+ib).style.background="url("+ AdminPath +"images/select-on.gif) no-repeat right bottom";
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
		if (spvalue.split(",")[j] =="")
		{
			selectover = 0
		}
	}
	if (selectover == 1)
	{
		 var url="Ajax_data.asp?action=getspinfo&id="+id+"&name="+escape(spvalue)+"&price="+price+"&t="+Math.random();
		 $.ajax({
		 type:"POST",
		 url:url,
		 data:"",
		 success:function(returntxt){
//			var stock = returntxt.split("|")[0];
//			var mes = returntxt.split("|")[1];
//			if (stock == 0)
//			{
//				$("#button_"+listcode+"_"+id+"").disabled = true;
//			}else{
//				$("#button_"+listcode+"_"+id+"").disabled = false;
//			}
			if(returntxt.length > 0)
			{
				$("#result_"+listcode+"_"+id+"").html(returntxt);
			}
			else
			{
				$("#result_"+listcode+"_"+id+"").html("");
			}
		 }
		})	
	}
}