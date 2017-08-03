function checkout1(t){
	var id=t;
	if(id=='')return false;
	var quantity=document.getElementById("quantity"+id).value.replace(/\D/g, '');
	document.getElementById("quantity"+id).value=quantity;
	var price=document.getElementById("price"+id).value;
	var money=(quantity * price).toFixed(2);
	return money;
}//end fn

function allpayfn(){

	var $chkpay=$(":input[name=chkpay]");6
	var allpays=0;
	for(i=0;i<$chkpay.size();i++){
		allpays = allpays + new Number(checkout1(  $chkpay.eq(i).val() ));
		//checkout1(  $chkpay.eq(i).val() )
		//alert( $chkpay.eq(i).val() 

	}
	allpays=allpays.toFixed(2);
	$("#allpay").html('￥'+allpays)

}

function checkout(t){
	var id=t;
	if(id=='')return false;
	var quantity=document.getElementById("quantity"+id).value.replace(/\D/g, '');
	document.getElementById("quantity"+id).value=quantity;
	var price=document.getElementById("price"+id).value;
	var money=(quantity * price).toFixed(2);
	document.getElementById("money"+id).innerHTML=money;
	allpayfn();
	$.ajax({type: "POST",url:"?myaction=modifyNumber",data:"sid="+escape(id)+"&quantity="+escape(quantity),timeout:20000});

}//end fn



function product_plus(t,s){
	var id=t;
	var kc=s;
	if(id=='')return false;

	var quantity=new Number(document.getElementById("quantity"+id).value.replace(/\D/g, ''))+1;

	if(quantity<=kc){document.getElementById("quantity"+id).value=quantity;checkout(id);}
	showcount()
}//end fn
function product_minus(t){
	var id=t;
	if(id=='')return false;

	var quantity=new Number(document.getElementById("quantity"+id).value.replace(/\D/g, ''))-1;
	if(quantity>=1){document.getElementById("quantity"+id).value=quantity;checkout(id)}
	showcount()
}//end fn
