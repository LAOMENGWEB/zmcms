/*宽度调整*/
function widthtz(){
	if(document.documentElement.clientWidth>800){
		$(".zmcont").width(340);
	}
	$(".footerbox").width($(".zmcont").width());
	$(".footerboxnav").width($(".zmcont").width());
	$(".footerlist").width($(".zmcont").width());
	$(".focus .bd li img").width($(".zmcont").width());
	$(".focus1 .bd li img").width($(".zmcont").width());
	$(".swipe  li img").width($(".zmcont").width());
}
widthtz();
$(window).on('resize', function(e) {
	widthtz();
});
//等高
function metHeight(group){
	tallest=0;
	group.each(function(){
		thisHeight=$(this).height();
		if(thisHeight>tallest){
			tallest=thisHeight;
		}
	});
	group.height(tallest);
}



//右上角功能----------------------------------------------------------
function topmorebox(tboxjq,dom){
	var dok = dom.css('display');
	$(".topmorebox").hide();
	$(".top-right li").removeClass('now');
	if(dok=='none'){
		dom.show();
		tboxjq.addClass('now');
	}
}
//语言
var ler = $(".top-right .lang");
if(ler.size()>0){
	var lerdom = ler;
	lerdom.click(function(){
		topmorebox(ler,$(".langlist"));
	});
}
//图片轮播----------------------------------------------------------
$('.iflash_slides').slider({loop:true,autoPlay:false,interval:4000,arrow:false,dots:false});

if($('.flash_slides li').size()>1){//Banner
	$('.flash_slides').slider({loop:true,autoPlay:false,interval:4000,arrow:false,dots:false});
	$(".flash_slides .ui-imglazyload").css({ 'width':'100%','height':falsh_y+'px' });
}

//over


//搜索
var ser = $(".top-right .seach");
if(ser.size()>0){
	var serdom = ser;
	serdom.click(function(){
		topmorebox(ser,$(".seachbox"));
		seachboxwd();
	});
}

//栏目下拉菜单
var cer = $(".top-right .column");
if(cer.size()>0){
	cer.on('click',function(){
		topmorebox(cer,$(".type3"));
	});
}
//搜索框宽度
function seachboxwd(){
	var swd = $('.seachbox_box').width() - 30 - ($('.seachbox input.submit').width() + 20);
	$(".seachbox input.text").width(swd);
	var ls = $(".top-right li.tlist").size();
	var wd = 11;
	if(ls==2)wd = 46 + 11;
	if(ls==3)wd = 46 + 46 + 11;
	$(".seachbox_box i").css("right",wd+'px');
}
$(window).on('resize', function(e) {
	seachboxwd();
});















//图片延迟加载---------------------------------------------------------
function metmobile_imgloading(dom){
	dom.each(function(){
		var imgurl = $(this).attr('src');
		$(this).after("<div class=\"ui-imglazyload\" data-url='"+imgurl+"'></div>");
		$(this).remove();
	});
	$('.ui-imglazyload').imglazyload({
		container: $('.zmcont')
	});
}
metmobile_imgloading($(".flexslider img,.sidebar img"));
$.fn.imglazyload.detect();
//over



//主导航面板
if($('nav.panel').size()>0){
	var myScroll1 = new iScroll('wrapper_nav');
}
$(function ($) {
	$('.panel').panel({
		contentWrap: $('.metcont'),
		scrollMode: 'hide'
	});
	$('.panel').on('close',function(){
		$('.panel .tapmengban').show();
		$('#navmianban').removeClass('now');
		$('.panel').hide();
	});
	$('.panel .tapmengban').height($(window).height());
	$('#navmianban').tap(function () {
		$(".topmorebox").hide();
		$(".top-right li").removeClass('now');
		$('#navmianban').addClass('now');
		setTimeout(" $('.panel .tapmengban').hide() ", 600);
		$('.panel').panel('toggle','push');//overlay' | 'reveal' | 'push'
		//$('.metcont').height($(window).height());
		//onCompletion(myScroll1);
	});
	$('.panel h3.title').on('click', function () {
		$('.panel').panel('close');
	});
}(Zepto));

//内页分类导航面板----------------------------------------------------------
if($('#sidebar').size()>0){
	var myScroll2 = new iScroll('wrapper_sidebar'); 
}
$('.sidebar_jsbox').height($(window).height());
$(function ($) {
	if($('#sidebar').size()>0){
		$('#sidebar').panel({
			contentWrap: $('.sidebar_jsbox'),
			scrollMode: 'fix'
		});
		$('#sidebar').on('close',function(){
			$('#sidebar').hide();
			$('.sidebar_jsbox').hide();
		});
		$('.moresidebar').click(function () {
			$('.sidebar_jsbox').show();
			$('#sidebar').panel('toggle','overlay');
		});
		$('#sidebar h3.title').on('click', function () {
			$('#sidebar').panel('close');
		});
	}
}(Zepto));
//over

$("#navnow1_"+classnow).addClass("now");

/*内页选项卡*/
var tags = $(".Tabtitle div.l");
if(tags.size()>1){
	tags.each(function(i){
		tagsdom = tags[i];
		tagsdom.addEventListener('touchstart', function(e){
			//alert
			$(".mb_Tabbox").hide();
			tags.removeClass('now');
			$(".mb_Tabbox_"+i).show();
			$(this).addClass("now");
			onCompletion(myScroll);
		}, false);
	});
}


/*手机键盘弹出时重定义高度*/
var wht = $(window).height();
$(window).on('orientationchange', function(e) {
        _landscape2 = !!(window.orientation & 2);
}).on('resize', function(e) {
	if($(window).height()<wht){
		$("#footer").hide();
		$("#wrapper").css('bottom','0px');
	}else{
		$("#footer").show();
		$("#wrapper").css('bottom','45px');
	}
	onCompletion(myScroll);
});

/*表格横向滚动*/
var tables = $(".editor table");//<div id="wrapper_nav"><div id="scroller_nav">
if(tables.size()>0){
	tables.wrap('<div class="wrapper_table"><div class="scroller_table"></div></div>');
	tables.each(function(){
		$(this).css({'margin-left':'0px','margin-right':'0px'});
		$(this).parent('.scroller_table').width($(this).width()+20);
	});
	$(".wrapper_table").iScroll({'vScroll':false});
	//var myScrolltable = new iScroll('wrapper_table',{'vScroll':false});
}

//右下角弹出图标----------------------------------------------------------
        window.addEventListener("DOMContentLoaded", function () {
            btn = document.getElementById("info-nr-btn");
            btn.onclick = function () {
                var divs = document.getElementById("info-nr-phone").querySelectorAll("div");
                var className = className = this.checked ? "on" : "";
                for (i = 0; i < divs.length; i++) {
                    divs[i].className = className;
                }
                document.getElementById("jisou-info").style.display = "on" == className ? "block" : "none";
            }
        }, false);
        $(function () {
            new Swipe(document.getElementById('jisou-banner'), {
                speed: 500,
                auto: 3000,
                callback: function () {
                    var lis = $(this.element).next("ol").children();
                    lis.removeClass("on").eq(this.index).addClass("on");
                }
            });
        });
//over

//首页index区块位置----------------------------------------------------------
niming();

$(window).on('resize', function(e){niming();});

function niming(){

var headerH = $(".zmcont header").height();          //顶部高
var footerH = $("#footernav").height();              //底部高
var bodyH = $(window).height();                    //屏幕高
    innerH = bodyH - headerH - footerH;            //屏幕高度减去顶部底部还剩下的高度

var idElement = $("#indexId");            
var indexH = idElement.height();                   //内容部分高度   
var metWidth = $(".zmcont").width();               //手机模板宽度
var indexSize = idElement.data('value');           //首页方块整体位置
var footText = $('.index-foot-text').height();     //底部信息高度
var headnav_ok=$("header").data("value");            //头部是否开启
var bannerH = $(".iflexslider").data("height");        //banner高度

indexbox = indexH + footText + 70;
boxHeight = innerH > indexbox?innerH:indexbox;
contentH = boxHeight > bannerH?boxHeight:bannerH;

$(".iflexslider").height(contentH);
$(".iflexslider li").css('height',contentH);
$(".iflexslider li a").height(contentH);
	
if(headnav_ok==1){headnav_top="30px";}else{headnav_top = headerH + 30;};

indexBlock=(contentH-footText-indexH)/2;              //除去底部，显示区域的一半；

switch(indexSize)
{
   case 1:idElement.css({"top":headnav_top,"left":"10px"});break;
   case 2:idElement.css({"top":headnav_top,"left":(metWidth-160)/2});break;
   case 3:idElement.css({"top":headnav_top,"right":"10px"});break;
   case 4:idElement.css({"top":indexBlock+headerH,"left":"10px"});break;
   case 5:idElement.css({"top":indexBlock+headerH,"left":(metWidth-160)/2});break;
   case 6:idElement.css({"top":indexBlock+headerH,"right":"10px"});break;
   case 7:idElement.css({"bottom":footText+30,"left":"10px"});break;
   case 8:idElement.css({"bottom":footText+30,"left":(metWidth-160)/2});break;
   default :idElement.css({"bottom":footText+30,"right":"10px"});break; 
}

}

//over