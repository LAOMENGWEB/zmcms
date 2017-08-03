

<script>
function showyuyan()
{
	var htmlurl=parent.frames["main"].location.href;
$.ajax({
	type:"get",
	url:"yuyanqh.asp?htmlurl="+ encodeURIComponent(htmlurl),
	data:"r="+Math.round(Math.random()*10000),
	beforesend:function(){$("#yuyan").html("载入中...");},
	error:function(){$("#yuyan").html("获取信息失败");},
	success:function(data){$("#yuyan").html(data);}
});
}

$(document).ready(function(){ 
showyuyan();
}); 
</script>
