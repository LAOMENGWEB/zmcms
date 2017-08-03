<%
''' 系统核心函数

dim Cls,sysdata

set Cls=new Cls_function
	Cls.init
class Cls_function
	private reg,match,matches
	private fso,stm
	public db,temp,page,sysdb
	
	private sub class_initialize
		set reg=new regexp
			reg.ignorecase=true
			reg.global=true
		set fso=createobject("scripting.filesystemobject") 
		set stm=server.createobject("adodb.stream")
		set sysdb=server.createobject("scripting.dictionary")
	end sub
	
	private sub class_terminate
		set reg=nothing
		set fso=nothing
		set stm=nothing
		set page=nothing
		set db=nothing
	end sub
	
	public sub init()
		set db=new Cls_dbbase
	end sub
	
	public function runtime()
		runtime=formatnumber(timer()-startime,4,true,false,true)
	end function
	
	public sub echo(byval t0)
		response.write t0
	end sub
	
	public sub die()
		response.end()
	end sub
		
	public sub go(byval t0)
		response.redirect t0
	end sub
	
	'字符长度
	function strlen(byval t0)
		if isnull(t0) then strlen=0:exit function
		strlen=len(t0)
	end function
	
	''返回上一页
	public sub back(byval t0)
		echo "<script>alert("""&t0&""");location.href=""javascript:history.go(-1)""</script>"
	end sub
	
	''返回指定地址
	public sub backurl(byval t0,byval t1)
		echo "<script>alert("""&t0&""");location.href="""&t1&"""</script>"
	end sub
	
	''指定script代码
	public sub script(byval t0)
		echo "<script>"&t0&"</script>"
	end sub
	
	''IF语句
	public function iif(byval t0,byval t1,byval t2)
		if t0 then iif=t1 else iif=t2
	end function
	
	''ElseIf语句
	public function iiif(byval t0,byval t1,byval t2,byval t3,byval t4)
		if t0 then
			iiif=t2
		elseif t1 then
			iiif=t3
		else
			iiif=t4
		end if
	end function
	
	''过滤Sql语句
	public function sqlstr(byval t0)
		sqlstr="'"&replace(t0,"'","''")&"'"
		sqlstr=replace(sqlstr,"%27","")
	end function
	
	''querystring参，t1为0时过滤空格
	public function fget(byval t0,byval t1)
		fget=request.querystring(t0)
		if t1=0 then fget=trim(fget)
	end function
	
	''post参数，t1为0时过滤前后空格
	public function fpost(byval t0,byval t1)
		fpost=request.form(t0)
		if t1=0 then fpost=trim(fpost)
	end function
	
	'query参数，t1为0时过滤前后空格
	public function fquery(byval t0,byval t1)
		fquery=request(t0)
		if t1=0 then fquery=trim(fquery)
	end function
	
	''判断是否数字,返回True|False
	public function isnum(byval t0)
		isnum=true
		if not isnumeric(t0) then isnum=false
	end function
	
	''判断是否货币，返回True|False
	public function ismoney(byval t0)
		ismoney=true
		if not isnumeric(t0) then ismoney=false:exit function
		if t0<0 then ismoney=false
	end function
	
	''判断系统组件是否安装，返回True|False
	public function isinstall(byval t0)
		err.clear
		on error resume next
		isinstall=false
		dim obj
		set obj=server.createobject(t0)
			if err.number=0 then isinstall=true
		set obj=nothing
		err.clear()
	end function
	
	''生成随机数,t0为字符数
	public function getrnd(byval t0)
		randomize
		dim n1,n2,n3
		t0=getint(t0,10)
		do while len(getrnd)<t0
			n1=cstr(chrw((57-48)*rnd+48))
			n2=cstr(chrw((90-65)*rnd+65))
			n3=cstr(chrw((122-97)*rnd+97))
			getrnd=getrnd&n1&n2&n3 
		loop
	end function
	
	''生成时间格式随机文件名
	public function getrndfilename()
		dim t0,t1
		randomize
		t0=int(900*rnd)+100
		t1=now()
		getrndfilename=year(t1)&month(t1)&day(t1)&hour(t1)&minute(t1)&second(t1)&t0
	end function
			
	public function getkey()
		getkey=getrnd(5)
	end function
	
	'获取本目录名称
	public function get_thisfolder
		dim a,b,c
		a=request.servervariables("script_name")
		b=split(a,"/")
		c=ubound(b)
		if c<=1 then
			get_thisfolder=""
		else
			get_thisfolder=b(c-2)
		end if
	end function
	
	''取整数，如果t0非数字，返回默认值t1
	public function getint(byval t0,byval t1)
		on error resume next
		if isnum(t0) then getint=clng(t0) else getint=t1
	end function
	
	''格式化日期，t2控制是否带年，t1控制年与月中间的符号，如- .
	public function getdate(byval t0,byval t1,byval t2)
		getdate=""
		if t2=1 then getdate=getdate&year(t0)&t1
		getdate=getdate&right("0"&month(t0),2)&t1
		getdate=getdate&right("0"&day(t0),2)
	end function
	
	''格式化日期
	public function format_time(t0,t1)
		Dim y, m, d, h, mi, s
		format_time = ""
		If IsDate(t0) = False Then Exit Function
		y = cstr(year(t0))
		m = cstr(month(t0))
		If len(m) = 1 Then m = "0" & m
		d = cstr(day(t0))
		If len(d) = 1 Then d = "0" & d
		h = cstr(hour(t0))
		If len(h) = 1 Then h = "0" & h
		mi = cstr(minute(t0))
		If len(mi) = 1 Then mi = "0" & mi
		s = cstr(second(t0))
		If len(s) = 1 Then s = "0" & s
		Select Case t1
		Case 0
			format_time = m & "-" & d
		Case 1
			' yyyy-mm-dd hh:mm:ss
			format_time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
		Case 2
			' yyyy-mm-dd
			format_time = y & "-" & m & "-" & d
		Case 3
			' yyyy.mm.dd
			format_time = y & "." & m & "." & d
		Case 4
			' hh:mm:ss
			format_time = h & ":" & mi & ":" & s
		Case 5
			' yyyy年mm月dd日
			format_time = y & "年" & m & "月" & d & "日"
		Case 6
			' yyyymmdd
			format_time = y & m & d
		End Select
	end function
	
	''格式化日期为2014-01-01
	public function numtodate(byval t0)
		if not isnumeric(t0) or len(t0)<>8 then
			numtodate=""
		else
			numtodate=mid(t0,1,4)&"-"&mid(t0,5,2)&"-"&mid(t0,7,2)
		end if
	end function
	
	''把字符转换为数字，小数点后2位
	public function getnum(byval t0)
		getnum=formatnumber(t0,2,true,false,true)
	end function
	
	''把字符转换为货币，小数点后2位
	public function getmoney(byval t0)
		getmoney=formatcurrency(t0,2,true,false,true)
	end function
	
	'转换百分比,t1为小数点后位数
	public function getpcent(byval t0,byval t1)
		dim t2
		if t1="" then
			t1=0
			t2="0%"
		else
			t2="0.00%"
		end if
		if isnum(t0) then getpcent=formatpercent(t0,t1,true) else getpcent=t2
	end function
	
	public function geturlparam()
		dim x,queryarray,tempstr
		for each x in split(request.querystring,"&")
		   queryarray=split(x,"=")
		   if ubound(queryarray)>0 then
			   if lcase(queryarray(0))<>"page" then
				   tempstr=tempstr&queryarray(0)&"="&queryarray(1)&"&"
			   end if
		   end if
		next
		geturlparam=tempstr
	end function
	
	''获取当前IP
	public function getip()
		getip=request.servervariables("http_x_forwarded_for")
		if getip="" then getip=request.servervariables("remote_addr")
		if not(checkstr(getip,"ip")) then getip="unknow"
	end function
	
	''通过xmlhtml方式获取远程文件内容
	public function gethttp(byval t0,byval t1)
		select case t1
			case "get","post"
			case else:t1="get"
		end select
		on error resume next
		dim http
		set http=server.createobject("microsoft.xmlhttp")
			http.open t1,t0,false
			http.setRequestHeader "If-Modified-Since","0"
			http.send()
			gethttp=http.responsetext
		set http=nothing
	end function
	
	'获取参数1,2,3,4,5
	public function deal_strmid(byval t0,byval t1)
		if len(t0)=0 or len(t1)=0 then deal_strmid="":exit function
		dim i,t2,t3
		t2=split(t0,t1)
		t3=""
		for i=0 to ubound(t2)
			if len(t2(i))>0 then
				if len(t3)=0 then
					t3=t2(i)
				else
					t3=t3&t1&t2(i)
				end if
			end if
		next
		deal_strmid=t3
	end function
	
	'判断是否图片，返回值1|0
	public function is_pic(byval t0)
		if isnull(t0) or len(t0)=0 then is_pic=0:exit function
		if instr(t0,".")<0 then is_pic=0:exit function
		if len(t0)<4 then is_pic=0:exit function
		select case lcase(right(t0,3))
			case "gif","jpg","bmp","png":is_pic=1
			case else:is_pic=0
		end select
	end function
	
	'判断是否在线视频
	public function is_video(byval t0)
		if isnull(t0) or len(t0)=0 then is_video="other":exit function
		if instr(t0,".")<0 then is_video="other":exit function
		if len(t0)<4 then is_video="other":exit function
		if instr(lcase(t0),".swf")>0 and instr(lcase(t0),"http://") then is_video="swf":exit function
		select case lcase(right(t0,3))
			case "flv":is_video="flv"
			case "swf":is_video="swf"
			case else:is_video="other"
		end select
	end function
	
	''格式化HTML
	public function nohtml(byval t0)
		if strlen(t0)=0 then nohtml="":exit function
		reg.pattern="<.+?>"
		set matches=reg.execute(t0)
		for each match in matches
			t0=replace(t0,match.value,"")
		next
		t0=replace(t0,"&nbsp;"," ")
		t0=replace(t0,"　","")
		nohtml=t0
	end function
	
	''转换HTML代码，过滤代码
	public function enhtml(byval t0)
		if isnull(t0) then enhtml="":exit function
		if t0="<p>&nbsp;</p>" then enhtml="":exit function
		t0=replace(t0,"&","&amp;")
		t0=replace(t0,"'","&#39;")
		t0=replace(t0,"""","&#34;")
		t0=replace(t0,"<","&lt;")
		t0=replace(t0,">","&gt;")
		reg.pattern="(w)(here)"
		t0=reg.replace(t0,"$1h&#101;re")
		reg.pattern="(s)(elect)"
		t0=reg.replace(t0,"$1el&#101;ct")
		reg.pattern="(i)(nsert)"
		t0=reg.replace(t0,"$1ns&#101;rt")
		reg.pattern="(c)(reate)"
		t0=reg.replace(t0,"$1r&#101;ate")
		reg.pattern="(d)(rop)"
		t0=reg.replace(t0,"$1ro&#112;")
		reg.pattern="(a)(lter)"
		t0=reg.replace(t0,"$1lt&#101;r")
		reg.pattern="(d)(elete)"
		t0=reg.replace(t0,"$1el&#101;te")
		reg.pattern="(u)(pdate)"
		t0=reg.replace(t0,"$1p&#100;ate")
		reg.pattern="(\s)(or)"
		t0=reg.replace(t0,"$1o&#114;")
		reg.pattern="(java)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		reg.pattern="(j)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		reg.pattern="(vb)(script)"
		t0=reg.replace(t0,"$1scri&#112;t")
		if instr(t0,"expression")<>0 then
			t0=replace(t0,"expression","e&#173;xpression",1,-1,0)
		end if
		enhtml=t0
	end function
	
	''逆向转换HTML
	public function dehtml(byval t0)
		if isnull(t0) then
			dehtml=""
			exit function
		end if
		t0=replace(t0,"&amp;","&")
		t0=replace(t0,"&#39;","'")
		t0=replace(t0,"&#34;","""")
		t0=replace(t0,"&lt;","<")
		t0=replace(t0,"&gt;",">")
		t0=replace(t0,chr(10),vbcrlf)
		dehtml=t0
	end function
	
	''检查字符
	public function checkstr(byval t0,byval t1)
		dim t2
		select case t1
			case "en":t2="^[a-zA-Z]+$"
			case "cn":t2="^[\u4e00-\u9fa5]+$"
			case "int":t2="^[-\+]?\d+$"
			case "price":t2="^\d+(\.\d+)?$"
			case "username":t2="^[a-zA-Z0-9_]{5,20}$"
			case "password":t2="^[a-zA-Z0-9_]{6,16}$"
			case "email":t2="^[\w\-\.]+@[a-zA-Z0-9]+\.(([a-zA-Z0-9]{2,4})|([a-zA-Z0-9]{2,4}\.[a-zA-Z]{2,4}))$"
			case "qqemail":t2="^[\w\-\.]+@qq.com"
			case "tel":t2="^((\(\+?\d{2,3}\))|(\+?\d{2,3}\-))?(\(0?\d{2,3}\)|0?\d{2,3}-)?[1-9]\d{4,7}(\-\d{1,4})?$"
			case "mobile":t2="^(\+?\d{2,3})?0?1(3\d|5\d|8[06789])\d{8}$"
			case "zipcode":t2="^\d{6}$"
			case "qq":t2="^[1-9]\d{4,15}$"
			case "url":t2 = "^(http|https|ftp):\/\/[a-zA-Z0-9]+\.[a-zA-Z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\""])*$"
			case "ip":t2="^((25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)\.){3}(25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)$"
			case "file":t2="^[a-zA-Z0-9/_-]{1,50}$"
			case "filename":t2="^[a-zA-Z0-9._/]{1,50}$"
		end select
		reg.pattern=t2
		checkstr=reg.test(trim(t0))
	end function
	
	''切割字符
	public function cutstr(byval t0,byval t1,byval t2)
		dim l,t,c,i
		if strlen(t0)=0 then cutstr="":exit function
		l=len(t0)
		t1=int(t1)
		t=0
		for i=1 to l
			c=ascw(mid(t0,i,1))
			if c<0 or c>255 then t=t+2 else t=t+1
			if t>=t1 then
				cutstr=left(t0,i)
				if t2=1 then cutstr=cutstr&"..."
				exit for
			else
				cutstr=t0
			end if
		next
	end function
	
	''查检是否外部提交
	public function checkpost
		checkpost=false
		dim t0,t1,t2
		t0=cstr(request.servervariables("http_referer"))
		t1=cstr(request.servervariables("server_name"))
		t2=request.servervariables("http_user_agent")
		if instr(lcase(t2),"firefox")>0 Then
			checkpost=true
			exit function
		end if
		if instr(t0,replace(replace(t1,"http://",""),"www.",""))=0 then
			checkpost=false
		else
			checkpost=true
		end if
	end function
	
	''清除所有格式代码
	public function closehtml(byval t0)
		dim t1,i,t2,t3,j
		t1=array("p","div","span","table","ul","font","b","u","i","h1","h2","h3","h4","h5","h6")
		for i=0 to ubound(t1)
			t2=0:t3=0
			reg.pattern="\<"&t1(i)&"( [^\<\>]+|)\>"
			set matches=reg.execute(t0)
			for each match in matches
				t2=t2+1
			next
			reg.pattern="\</"&t1(i)&"\>"
			set matches=reg.execute(t0)
			for each match in matches
				t3=t3+1
			next
			for j=1 to t2-t3
				t0=t0+"</"&t1(i)&">"
			next
		next
		closehtml=t0
	end function
	
	''高亮某字符
	public function highlight(byval t0,byval t1)
		reg.pattern="("&t1&")"
		highlight=reg.replace(t0,"<font color=""red"">$1</font>")
	end function
	
	'获取代码里的图片
	public function getpic(byval t0)
		dim pic:pic=""
		reg.pattern="<img[^>]+src=""([^"">]+)""[^>]*>"
		set matches=reg.execute(t0)
		if matches.count>0 then
			for each match in matches
				if left(lcase(match.submatches(0)),7)<>"http://" then
					pic=pic&"|"&match.submatches(0)
				end if
			next
		end if
		if len(pic)>1 then pic=right(pic,len(pic)-1)
		getpic=pic
	end function
	
	
	public function get_outfile(byval t0)
		if strlen(t0)=0 then get_outfile="":exit function
		dim filename:filename=""
		dim filedata,t1
		set filedata=server.createobject("scripting.dictionary")
		reg.pattern="(http://(.[^>]+?)\.(gif|jpg|png|bmp))"
		set matches=reg.execute(t0)
		if matches.count>0 then
			for each match in matches
				t1=match.submatches(0)
				if filedata.exists(t1) then
					filedata.item(t1)=t1
				else
					filedata.add t1,t1
				end if
			next
		end if
		dim labeltag:labeltag=filedata.keys
		dim labelval:labelval=filedata.items
		if filedata.count>0 then
			dim i
			for i=0 to filedata.count-1
				if filename="" then
					filename=labelval(i)
				else
					filename=filename&"|"&labelval(i)
				end if
			next
		end if
		set filedata=nothing
		get_outfile=filename
	end function
		
	public function deal_outfile(byval t0,byval t1)
		dim htmlcontent:htmlcontent=t0
		dim filelist:filelist=get_outfile(htmlcontent)
		if filelist="" then deal_outfile=htmlcontent:exit function
		dim filearr,i
		filearr=split(filelist,"|")
		for i=0 to ubound(filearr)
			dim filetxt:filetxt=mid(filearr(i),instrrev(filearr(i),".")+1)
			dim filepath:filepath=uploaddir("")
			dim filename:filename=getrndfilename
			dim newfile:newfile=filepath&filename&"."&filetxt
			if save_outfile(filearr(i),array(filepath,filename,filetxt),t1) then
				htmlcontent=replace(htmlcontent,filearr(i),newfile)
			end if
		next
		deal_outfile=htmlcontent
	end function
	
	''保存外部图片
	public function save_outfile(byval t0,byval t1,byval t2)
		on error resume next
		dim http,filedata,filesize,filename
		filename=t1(0)&t1(1)&"."&t1(2)
		set http=server.createobject("microsoft.xmlhttp")
		with http
			.open "get",t0,false,"",""
			.send
			filedata=.responsebody
			filesize=getnum(round(lenb(filedata))/1024)
		end with
		with stm
			.type=1
			.open
			.write filedata
			.savetofile server.mappath(filename), 2
			.cancel
		.close
		end with
		if err then
			err.clear
			save_outfile=false
		else
			save_outfile=true
		end if
	end function
	
	''写入cookie
	public sub setcookie(byval t0,byval t1)
		response.cookies(prefix)(t0)=t1
	end sub
	
	''读取cookie
	public function loadcookie(t0)
		loadcookie=request.cookies(prefix)(t0)
	end function
	
	''写入session
	public sub setsession(byval t0,byval t1)
		session(prefix&t0)=t1
	end sub
	
	''读取session
	public function loadsession(byval t0)
		loadsession=session(prefix&t0)
	end function
	
	''设置时间格式上传目录
	public function uploaddir(byval t0)
		if strlen(t0)>0 then t0=t0&"/"
		Dim t1:t1=now()
		Dim t2:t2=year(t1)&right("0"&month(t1),2)&"/"
		uploaddir=webroot&"upfile/"&t0&t2
		newfolder uploaddir
	end function
	
	''判断是否目录存在
	public function isfolder(byval t0)
		isfolder=fso.folderexists(server.mappath(t0))
	end function
	
	'判断是否文件存在
	function isfile(byVal t0)
		isfile=fso.fileexists(server.mappath(t0))
	end function
	
	''新建目录
	public sub newfolder(byval t0)
		dim t1,t2,i
		err.clear
		on error resume next
		t0=server.mappath(t0)
		if fso.folderexists(t0) then exit sub
		t1=split(t0,"\")
		t2="" 
		for i=0 to ubound(t1)  
			t2=t2&t1(i)&"\"
			if not fso.folderexists(t2) then fso.createfolder(t2)
		next
	end sub
	
	''移动目录
	public sub editfolder(byval t0,byval t1)
		err.clear
		on error resume next
		if fso.folderexists(server.mappath(t0)) then
			fso.movefolder server.mappath(t0),server.mappath(t1)
		end if
	end sub
	
	''删除目录
	public sub delfolder(byval t0)
		err.clear
		on error resume next
		dim f
		set f=fso.getfolder(server.mappath(t0))
		if not isnull(t0) then f.delete true
		set f=nothing
	end sub
	
	''读取目录文件
	public function loadfile(byval t0)
		if len(t0)=0 then loadfile="文件路径不能为空":exit function
		err.clear
		on error resume next
		dim t1
		with stm
			.type=2
			.mode=3 
			.charset="utf-8"
			.open
			.loadfromfile server.mappath(t0)
			t1=.readtext
			.close
		end with
		if err then
			loadfile=t0&"<br />"&err.description&"<br />":err.clear
		else
			if len(t1)=0 then t1=t0&":<br />文件大小为0"
			loadfile=t1
		end if
	end function
	
	''新建文件
	public sub newfile(byval t0,byval t1,byval t2,byval t3)
		if t3="" then t3="utf-8"
		newfolder t0
		err.clear
		on error resume next
		with stm
			.type=2
			.mode=3
			.charset=t3
			.open
			.writetext t2
			.savetofile server.mappath(t0&t1),2
			.close
		end with
	end sub
	
	''移动文件
	public sub editfile(byval t0,byval t1)
		err.clear
		on error resume next
		if fso.fileexists(server.mappath(t0)) then
			fso.movefile server.mappath(t0),server.mappath(t1)
		end if
	end sub
	
	'删除文件
	public sub delfile(byval t0)
		err.clear
		on error resume next
		if fso.fileexists(server.mappath(t0)) then fso.deletefile server.mappath(t0)
	end sub
	
	''字符转换为二进制
	public function stringtobytes(byval t0)
		with stm
			.type=2
			.charset="utf-8"
			.open
			.writetext t0
			.position=0
			.type=1
			.read 3
			stringtobytes=.read(-1)
			.close
		end with
	end function
	
	''二进制转换为字符
	public function bytestostring(byval t0)
		with stm
			.type=1
			.open
			.write t0
			.position=0
			.type=2
			.charset="utf-8"
			bytestostring=.readtext(-1)
			.close
		end with
	end function
	
	''base64e加密
	public function base64encode(byval t0)
		dim t1,t2
		set t1=server.createobject("msxml2.domdocument")
		set t2=t1.createelement("b64")
		t2.datatype="bin.base64"
		if vartype(t0)=(vbbyte or vbarray) then
			t2.nodetypedvalue=t0
		else
			t2.nodetypedvalue=stringtobytes(t0)
		end if
		base64encode=t2.text
		set t2=nothing
		set t1=nothing
	end function
	
	''base64解密
	public function base64decode(byval t0)
		dim t1,t2
		set t1=server.createobject("msxml2.domdocument")
		set t2=t1.createelement("b64")
		t2.datatype="bin.base64"
		t2.text=t0
		base64decode=bytestostring(t2.nodetypedvalue)
		set t2=nothing
		set t1=nothing
	end function
		
	''获取Tags
	public function get_tags_arr(byval t0)
		if strlen(t0)=0 then
			get_tags_arr=array()
			exit function
		end if
		get_tags_arr=split(t0,",")
	end function
	
	''获取城市名称
	public function get_cityname(byval t0)
		dim t1:t1=getip
		if t1="unknow" or t1="127.0.0.1" or t1="" then
			get_cityname=t0
			exit function
		end if
		if Cls.loadcookie("get_cityname_"&t1)<>"" then
			get_cityname=left(enhtml(Cls.loadcookie("get_cityname_"&t1)),15)
		else
			dim str:str=gethttp("http://whois.pconline.com.cn/ip.jsp?ip="&t1&"","")
			dim t3
			if strlen(str)>0 then
				dim t2:t2=split(str," ")
				t3=t2(0)
			else
				t3=t0
			end if
			get_cityname=t3
			Cls.setcookie "get_cityname_"&t1,t3
		end if
	end function
	
	''判断字符是否存在另一个字符
	public function is_current(byval t0,byval t1,byval t2)
		if instr(","&t0&",",","&t1&",")>0 then
			is_current=t2
		end if
	end function
	
end class
%>