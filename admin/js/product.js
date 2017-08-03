var $=jQuery;
var ke = '1';
var li = 0;
function getSValue(i){
	var str="";
	var list=document.getElementsByName("SValue"+i);
	for(k=0;k<list.length;k++){
		if(list[k].checked){
			str+=list[k].value.split('|')[1]+",";
		}
	}
	str=str.substr(0,str.length-1);
	var arr2=str.split(',');
	return arr2;
}
function createHeader(arr,arrSValue){
	var objt=document.getElementById("Table_Specification");
	var rows=objt.insertRow(0);
	var list=arr.split(',');
	for(var i=0;i<list.length;i++){
	var cell=rows.insertCell(i);
	cell.innerHTML=list[i];
	cell.classname = "listhead";
	}
	var cell1 = rows.insertCell(list.length);
	cell1.style.cssText = "width:20%;";
	cell1.classname = "listhead";
	cell1.innerHTML = Lang_Product.split('$')[0];
	var cell2=rows.insertCell(list.length+1);
	cell2.innerHTML = Lang_Product.split('$')[1];
	cell2.style.cssText = "width:10%;";
	cell2.classname = "listhead";
	var cell3=rows.insertCell(list.length+2);
	cell3.innerHTML = Lang_Product.split('$')[2];
	cell3.style.cssText = "width:10%;";
	cell3.classname = "listhead";
	var cell4=rows.insertCell(list.length+3);
	cell4.innerHTML = Lang_Product.split('$')[3];
	cell4.style.cssText = "width:10%;";
	cell4.classname = "listhead";
	var cell5=rows.insertCell(list.length+4);
	cell5.innerHTML = Lang_Product.split('$')[4];
	cell5.style.cssText = "width:10%;";
	cell5.classname = "listhead";

}
function createBody(arr,arrSValue,id){
	var resultArray = arrSValue[0];
	var copyArray;
	for (var i = 1; i < arrSValue.length; i++)
	{
		copyArray = resultArray.splice(0, resultArray.length);
		for (var k = 0; k < copyArray.length; k++)
			for (var j = 0; j < arrSValue[i].length; j++)
				resultArray.push(copyArray[k]+"," + arrSValue[i][j]);
	}
	var obj_t=document.getElementById("Table_Specification");
	var length = obj_t.rows.length; 
	if(resultArray.length>0&&ke!='0'||length==0){
		createHeader(arr,arrSValue);
		ke = '0';
	}
	var kucun1 = 0;
	var kucun2 = 0;
	var shu = document.getElementById("kucun"+id+"").value;
	li = resultArray.length;
	document.getElementById("SpecificationCount").value=li;
	for(var c=0;c<resultArray.length;c++){
		kucun1 = parseInt(shu/resultArray.length);
		if(c==resultArray.length-1){
			kucun1 = parseInt(shu)-kucun2;
		}
		kucun2 += kucun1;
		var str=resultArray[c];
		var rows=obj_t.insertRow(c+1);
		rows.style.cssText=' ';
		var list=str.split(',');
		for(var b=0;b<list.length;b++){
			var cell_s=rows.insertCell(b);
			cell_s.innerHTML=list[b].split('$$$')[0];
			cell_s.style.cssText = " class='listrow'";
			str = str.replace(",","@")
		}
		var cell_s1 = rows.insertCell(list.length);
		cell_s1.innerHTML = '<input style=\"width:130px;\" class=\"form\" type="text" name="ajax_p_no" id="p_no'+c+'" value="-'+c+'" />';
		cell_s1.style.cssText = " class='listrow'";
		var cell_s2=rows.insertCell(list.length+1); 
		cell_s2.innerHTML='<input type="hidden" name="ajax_s_id" id="s_id'+c+'" value="0" /><input type="hidden" name="ajax_s_value" id="s_value'+c+'" value="'+str+'" />';
		cell_s2.innerHTML+='<input type="text" style=\"width:40px;\" class=\"form\" onblur=\"countle('+resultArray.length+','+id+')\" name="ajax_p_stock" id="p_stock'+c+'" value="'+kucun1+'" />';
		cell_s2.style.cssText = " class='listrow'";
		var cell_s3=rows.insertCell(list.length+2);
		cell_s3.innerHTML='<input type="text" style=\"width:40px;\" class=\"form\"  name="ajax_p_price" id="p_price'+c+'" value="" />';
		cell_s3.style.cssText = " class='listrow'";
		var cell_s4=rows.insertCell(list.length+3);
		cell_s4.innerHTML='<input type="text" style=\"width:40px;\" class=\"form\" name="ajax_p_cbprice" id="p_cbprice'+c+'" value="'+document.getElementById("cbprice").value+'" />';
		cell_s4.style.cssText = " class='listrow'";
		
		var cell_s5=rows.insertCell(list.length+4);
		cell_s5.innerHTML='<a style="cursor:pointer;" onclick="delRow('+resultArray.length+','+id+');">'+ Lang_Product.split('$')[4] +'</a>';
		cell_s5.style.cssText = " class='listrow'";
	}
}
function countle(num,id){
	var  count = 0.0;
    for(var i=0;i<num;i++){
		try{count=count+ parseInt(document.getElementById("p_stock"+i).value);}
		catch(e){continue;}
    }
    document.getElementById("kucun"+id+"").value = count;
	document.getElementById("total_kucun"+id+"").value = count;
}
function clearRow(){ 
	var objTable= document.getElementById("Table_Specification"); 
	var length = objTable.rows.length; 
	for(var i=1;i<length;i++) 
	{ 
		objTable.deleteRow(1); 
	}
}
function ResetRow(id){ 
	var objTable= document.getElementById("Table_Specification"); 
	var length = objTable.rows.length; 
	for(var i=0;i<length;i++) 
	{ 
		objTable.deleteRow(-1); 
	}
	var listc=document.getElementsByName("Hidden_SName");
	for(var i=0;i<listc.length;i++){
		listc[i].checked=false;
		var lists=document.getElementsByName("SValue"+listc[i].value.split('|')[0]);
		for(var s=0;s<lists.length;s++){
			lists[s].checked=false;
		}
	}
	display(id);
}
function delRow(len,id){
	var currRowIndex=event.srcElement.parentNode.parentNode.rowIndex;
	document.getElementById("Table_Specification").deleteRow(currRowIndex);
	countle(len,id);
}
function delValueRow(id,name,delid,len){
	 var url="Ajax_data.asp?action=getspecification&id="+id+"&name="+escape(name)+"&delid="+delid+"&t="+Math.random();
     $.ajax({
     type:"POST",
     url:url,
     data:"",
     success:function(returntxt){
		if(returntxt.length > 0)
		{
			$("#Get_Specification"+id+"").html(returntxt);
			countle(len,id);
		}
		else
		{
			$("#Get_Specification"+id+"").html("");
		}
     }
   })	
}
function getdiv(id){
	var arr=new Array();
	var arr2=new Array();
	var arrSValue=new Array();
	var flag=false;
	var Hidden_SName="";
	var allSValue="";
	var picbox="";
    var chkArr = document.getElementsByName("Hidden_SName");
    for (var i=0;i<chkArr.length ;i++ ){ 
		if(chkArr[i].checked){
			arr += chkArr[i].value.split('|')[1]+",";
			arr2 += chkArr[i].value+",";
			flag=true;
		}
	}
	arr = arr.substring(0,arr.length-1);
	arr2 = arr2.substring(0,arr2.length-1);
	var list=arr2.split(',');
	for(var i=0;i<list.length;i++){
		if(getSValue(list[i].split('|')[0])!=""){
			arrSValue[i]=getSValue(list[i].split('|')[0]);
			allSValue+=getSValue(list[i].split('|')[0])+",";
		}else{
			flag=false;
		}
	}
	allSValue = allSValue.substring(0,allSValue.length-1);
	var slist=allSValue.split(',');
	picbox = '<table cellspacing="0" cellpadding="0" class="tablelist">'
	for(var i=0;i<slist.length;i++){
		picbox += '<tr><td>'+ slist[i].split('$$$')[0] +'</td><td><input type="text" name="SValuePic" id="url'+i+'" size="20" class="form" /></td><td>请手动输入要关联的图片地址</td></tr>';
	}
	picbox += '</table>'
	$("#Get_SpecificationPic"+id+"").html(picbox);
	if(flag){
		$("#Table_Specification").show();
		clearRow();
		document.getElementById("kucun"+id+"").setAttribute("disabled", "disabled");
		document.getElementById("Get_Specification"+id+"").innerHTML = "";
		createBody(arr,arrSValue,id);
	}else{
		display(id);
		clearRow();
	}
}
function display(id){
	document.getElementById("kucun"+id+"").removeAttribute("disabled");
	document.getElementById("Get_Specification"+id+"").innerHTML = "";
	document.getElementById("Get_SpecificationPic"+id+"").innerHTML = "";
}