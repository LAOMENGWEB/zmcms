var swfu,selQueue=[],arrUrl=[],arrName=[],allSize=0,uploadSize=0,processers={};
var aien =function(selecter){return document.getElementById(selecter);};
function formatBytes(bytes) {
	var s = ['Byte', 'KB', 'MB', 'GB', 'TB', 'PB'];
	var e = Math.floor(Math.log(bytes)/Math.log(1024));
	return (bytes/Math.pow(1024, Math.floor(e))).toFixed(2)+" "+s[e];
}
function showButton(){
	//aien("clearU").disabled=false;
	aien("startU").disabled=false;
}

function hideButton(){
	//aien("clearU").disabled=true;
	aien("startU").disabled=true;
}

function removeAllFiles()
{
	var file;
	for(i in selQueue)
	{
		file=selQueue[i];
		removeOne(file.id);
		allSize-=file.size;
	}
	allSize=0;
	hideButton();
	aien("progressArea").style.display="none";
}

function removeOne(id){
	file=selQueue[id];
	swfu.cancelUpload(id);
	delete selQueue[id];
	aien(id).parentNode.parentNode.removeChild(aien(id).parentNode);
	allSize-=file.size;
	if(allSize==0){
		hideButton();
		aien("progressArea").style.display="none";
	}
}
function startUploadFiles()
{
	if(swfu.getStats().files_queued)
	{
		hideButton();
		aien("progressArea").style.display="block";
		swfu.startUpload();
		processers["pall"] = new Processer("pall",aien("processer"),920,8);
		processers["pall"].init();
	}
	else alert('上传前请先添加文件');
}
function setFileState(file,txt)
{
	
	if(aien("process" + file.id)) aien("process" + file.id).innerHTML=txt;
	
}
function fileQueued(file)//队列添加成功
{
	if(selQueue.length==0)showButton();
	selQueue[file.id]=file;
	allSize+=file.size;
	//aien('#listarea').append('<tr id="'+file.id+'"><td>'+file.name+'</td><td>'+formatBytes(file.size)+'</td><td id="'+file.id+'_state">就绪</td></tr>');
	var str="<div id=\"" + file.id + "\" class=\"filelist\">";
	str+="<div class=\"filename\">" + file.name + "</div>";
	str+="<div id=\"process" + file.id + "\" class=\"fileprocess\">准备就绪,等待上传中...[<a href=\"javascript:void(0)\" onclick=\"removeOne('" + file.id + "');\">取消</a>]</div>";
	str+="</div>";
	var c = document.createElement("div");
	c.innerHTML = str;
	aien('listArea').appendChild(c);
}
function fileQueueError(file, errorCode, message)//队列添加失败
{
	var errorName='';
	switch (errorCode)
	{
		case SWFUpload.QUEUE_ERROR.QUEUE_LIMIT_EXCEEDED:
			errorName = "只能同时上传 "+this.settings.file_upload_limit+" 个文件";
			break;
		case SWFUpload.QUEUE_ERROR.FILE_EXCEEDS_SIZE_LIMIT:
			errorName = "选择的文件超过了当前大小限制："+this.settings.file_size_limit;
			break;
		case SWFUpload.QUEUE_ERROR.ZERO_BYTE_FILE:
			errorName = "零大小文件";
			break;
		case SWFUpload.QUEUE_ERROR.INVALID_FILETYPE:
			errorName = "文件扩展名必需为："+this.settings.file_types_description+" ("+this.settings.file_types+")";
			break;
		default:
			errorName = "未知错误";
			break;
	}
	alert(errorName);
}
function uploadStart(file)//单文件上传开始
{
	setFileState(file,'上传中…');
	processers["p" + file.id] = new Processer(file.id,aien("process" + file.id),330,8);
	processers["p" + file.id].init();
	
}
function uploadProgress(file, bytesLoaded, bytesTotal)//单文件上传进度
{
	var percent=Math.ceil((uploadSize+bytesLoaded)/allSize*100);
	var percent_single = Math.ceil(bytesLoaded/file.size*100);
	//processers["p" + file.id].update(bytesLoaded,bytesTotal);
	processers["pall"].update(uploadSize+bytesLoaded,allSize);
	setFileState(file,'总大小：' + formatBytes(bytesTotal) + '，已上传：' + formatBytes(bytesLoaded));
	if(bytesLoaded ==bytesTotal){
		setFileState(file,'总大小：' + formatBytes(bytesTotal) + '，已上传：' + formatBytes(bytesLoaded) + "。正在保存数据...");
	}
}
function uploadSuccess(file, serverData)//单文件上传成功
{
	var data;
	if(Setting['debug']==true)alert(serverData);
	try{eval("data=" + serverData);}catch(ex){};
	if(data['err']!=undefined&&data['msg']!=undefined)
	{
		if(!data.err){
			uploadSize+=file.size;
			var url=data.msg;
			var name=data.name;
			//setFileState(file,'上传成功,文件大小' + formatBytes(file.size) );
			arrUrl.push(url);
			arrName.push(name);
			setFileState(file,'总大小：' + formatBytes(file.size) + '。上传完成！');
			if(typeof window.CallBack=="function"){
				window.CallBack.apply(data,[]);	
			}
		}else{
			setFileState(file,data.msg);
		}
	}
	else setFileState(file,'上传失败！');
}
function uploadError(file, errorCode, message)//单文件上传错误
{
	setFileState(file,'上传失败！' + errorCode);
}
function uploadComplete(file)//文件上传周期结束
{
	if(swfu.getStats().files_queued)swfu.startUpload();
	else uploadAllComplete();
}
function uploadAllComplete()//全部文件上传成功
{
	var strUrl=arrUrl.join('|');
	var strName=arrName.join('|');
	//processers["pall"].update(1,1);
	//callback(strUrl+'|+|'+strName);
}

function Processer(id,showto,width,height){
	showto.innerHTML="";
	this.id = id;
	this.showto = showto;
	this.width = width;
	this.height = height;
	this.skin=null;
	this.process=null;
	this.percent=null;
	this.init = function(){
		var skin = document.createElement("div");
		skin.style.cssText="padding:0px; margin:0; border:1px #ccc solid;height:" + this.height + "px;width:" + (this.width-30) + "px; float:left;";
		var process = document.createElement("div");
		process.style.cssText="padding:0;margin:0;height:" + (this.height) + "px;width:0; background-color:#eee;";
		skin.appendChild(process);
		
		var percent = document.createElement("div");
		percent.style.cssText = "width:25px;height:" + this.height + "px;float:left; padding-left:5px;";
		this.skin=skin;
		this.process = process;
		this.percent = percent;
		this.showto.appendChild(skin);
		this.showto.appendChild(percent);
		
		var br = document.createElement("div");
		br.style.cssText = "width:0;height:0; clear:both";
		this.showto.appendChild(br);
	};
	this.update=function(current,total){
		var bili = current / total;
		var len = (this.width-30) * bili;
		this.process.style.width= len + "px";
		this.percent.innerHTML= Math.ceil(current/total*100) + "%";
	};
}
function pageInit()
{
	aien("progressArea").style.display="none";
	swfu = new SWFUpload({
		// Flash组件
		flash_url : "swfupload/upload.swf",

		// 服务器端
		upload_url: Setting["upload_url"],
		file_post_name : Setting["file_post_name"],
		post_params: {},

		// 文件设置
		file_types : Setting["file_types"],//文件格式限制
		file_types_description : Setting["file_types_description"],//文件格式描述
		file_size_limit : Setting["file_size_limit"],	// 文件大小限制
		file_upload_limit : Setting["file_upload_limit"],//上传文件总数
		file_queue_limit:0,//上传队列总数
		custom_settings : {},

		// 事件处理
		file_queued_handler : fileQueued,//添加成功
		file_queue_error_handler : fileQueueError,//添加失败
		upload_start_handler : uploadStart,//上传开始
		upload_progress_handler : uploadProgress,//上传进度
		upload_error_handler : uploadError,//上传失败
		upload_success_handler : uploadSuccess,//上传成功
		upload_complete_handler : uploadComplete,//上传结束

		// 按钮设置
		button_placeholder_id : "divAddFiles",
		button_width: 71,
		button_height: 18,
		button_window_mode: SWFUpload.WINDOW_MODE.TRANSPARENT,
		button_cursor: SWFUpload.CURSOR.HAND,
		button_text: '<span>上传文件</span>',
		button_text_style:"color:#ffffff",
		button_text_left_padding: 14,
		button_text_top_padding: 3,

		// 调试设置
		debug: false
	});
}