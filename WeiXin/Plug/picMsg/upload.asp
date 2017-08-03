<!--#include file="../../../inc/conn.asp"-->
<!--#include file="upload_class_vbs.asp"-->
<!--#include file="Images_Class.asp"-->
<%	
add=trim(request.QueryString("add"))
	dim sql,rs
    sql="select * from wzset"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
	Dim JpegType,JpegLocation,JpegTxt,JpegSize,JpegFont,JpegColor,JpegPic
	Dim strSpicW,strSpicH,strPicW,strPicH

if add="one" then
'调用小图尺寸处理图片,不给加水印
	strSpicW = rs("PicW")	'设置生成后缩略图宽度
	strSpicH = rs("PicH")	'设置生成后缩略图高度
	strPicW = rs("PicW")	'设置原图宽度大于该数字时开启缩略图生成功能
	strPicH = rs("PicH")	'设置原图高度大于该数字时开启缩略图生成功能
	JpegType = rs("Piclx")'水印类型,0:关闭水印功能,1:文字水印,2:图片水印
	if JpegType>0 then JpegType=3
else
	strSpicW = rs("PicBW")'设置生成后缩略图宽度
	strSpicH = rs("PicBH")'设置生成后缩略图高度
	strPicW = rs("PicBW")	'设置原图宽度大于该数字时开启缩略图生成功能
	strPicH = rs("PicBH")	'设置原图高度大于该数字时开启缩略图生成功能
	JpegType = rs("Piclx")'水印类型,0:关闭水印功能,1:文字水印,2:图片水印
end if
	
	
	
	'---------- 水印位置,0:左上,1:右上,2:左下,3:右下,4:居中
	JpegLocation = rs("picLocation")
	'---------- 水印文字
	JpegTxt = rs("Pictxt")
	'---------- 水印文字,字号
	JpegSize = rs("PicSize")
	'---------- 水印文字,字体
	JpegFont = "宋体"
	'---------- 水印文字,颜色
	JpegColor = rs("PicColor")
	'---------- 图片水印,图片
	JpegPic = "../"&rs("sytp")
	
Server.ScriptTimeout = 9999999
Dim Upload,path,tempCls,e
'===============================================================================
set Upload=new AnUpLoad				 				'创建类实例
Upload.SingleSize=clng(1000 * 1024) * 1024            			'设置单个文件最大上传限制,按字节计；默认为不限制
Upload.MaxSize=clng(1000 * 1024) * 1024            				'设置最大上传限制,按字节计；默认为不限制
Upload.Exe="jpg|gif|png|bmp"          				'设置合法扩展名,以|分割,忽略大小写
Upload.Charset="utf-8"								'设置文本编码，默认为gb2312
Upload.GetData()										'获取并保存数据,必须调用本方法
'===============================================================================

%>

<%if add="one" then%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>zmcms</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

</head>
 <style>   

#sfile {padding:10px; float:left}
#sfile li { list-style-type:none}
#sfile li .submit4 {background: #09F;border:none; margin:40px 0 0 30px; padding:10px; color:#FFF; font-size:14px; font-weight:bold; height:40px; width:100px
}

#sfile li img{ width:160px; height:120px; display:block; float:left; padding:3px; border:#CCC 1px solid; margin-right:10px}
</style>
<script type="text/javascript" src="../../../js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="../../../javascript/colorbox/jquery.colorbox-min.js"></script>
<script language="javascript" type="text/javascript">
		
        function test()
        {
         var value = $('#fuImage').val();
         parent.$('#weepic').val(value);	 
         parent.$("#mysmaimg").html("<img src='../../../"+value+"'  id='sc'/><br /><a href='delslt.asp?sltdz="+value+"' target='_blank' onclick='deltp()' id='lj'>删除</a>"); 
		 			 
		     parent.$.fn.colorbox.close();return false;
        }

 </script>
<body> 
<%
if Upload.ErrorID>0 then								'判断错误号,如果myupload.Err<=0表示正常
	response.Write("<Script Language=JavaScript >alert('" & Upload.description & " '); parent.$.fn.colorbox.close();</Script>")
else
	if Upload.files(-1).count>0 then 					'这里判断你是否选择了文件
		path=server.mappath("../../../UpLoadFile/image/maxpic") 					'文件保存路径(这里是files文件夹)
		'保存第一个文件(以新文件名保存)
		set tempCls=Upload.files("file1") 
		result = tempCls.SaveToFile(path,0,true)
		if result then
			strFile = tempCls.FileName
		UpLoadDir="../../../UpLoadFile/image/maxpic/"
			if  JpegType <>0 then
			sFileName = oJpegPic(UpLoadDir&strFile,UpLoadDir&"S"&strFile,strSpicW,strSpicH)
			'删除掉已经处理的文件
			Set fs = Server.CreateObject("Scripting.FileSystemObject")
            File = Server.MapPath(UpLoadDir&strFile)
            on Error Resume Next
            fs.DeleteFile File, True '强制删除只读文件
             set fs=nothing
			response.Write("<div id='sfile'><input  type='hidden' ID='fuImage'  value=UpLoadFile/image/maxpic/S"& tempCls.filename&" /><li><img src='../../../UpLoadFile/image/maxpic/S"& tempCls.filename &"' />上传完毕,请点击下面按钮插入图片<input type=""button"" value=""插入图片"" onclick=""test()"" class=""submit4""  id=""Button1"" /></li></div>")
			
		else
			sFileName = UpLoadDir&strFile
			response.Write "<div id='sfile'><input  type='hidden' ID='fuImage' value=UpLoadFile/image/maxpic/"& tempCls.filename &" /><li><img src='../../../UpLoadFile/image/maxpic/"& tempCls.filename &"'/> 上传完毕,请点击下面按钮插入图片<input type=""button"" value=""插入图片"" onclick=""test()"" class=""submit4""  id=""Button1"" /></li></li></div>"
		end if
		end if
		set tempCls=nothing
	end if
end if
set Upload=nothing


else


if Upload.ErrorID>0 then								'判断错误号,如果myupload.Err<=0表示正常
	response.Write("{err:true,msg:'" & Upload.description & "'}")
else
	if Upload.files(-1).count>0 then 					'这里判断你是否选择了文件
		path=server.mappath("../../../UpLoadFile/image/maxpic") 					'文件保存路径(这里是files文件夹)
		'保存第一个文件(以新文件名保存)
		set tempCls=Upload.files("filedata") 
		result = tempCls.SaveToFile(path,0,true)
		if result then
		
		strFile = tempCls.FileName
		UpLoadDir="../../../UpLoadFile/image/maxpic/"
			if  JpegType <>0 then
			sFileName = oJpegPic(UpLoadDir&strFile,UpLoadDir&"S"&strFile,strSpicW,strSpicH)
			'删除掉已经处理的文件
			Set fs = Server.CreateObject("Scripting.FileSystemObject")
            File = Server.MapPath(UpLoadDir&strFile)
            on Error Resume Next
            fs.DeleteFile File, True '强制删除只读文件
             set fs=nothing
			 response.Write("{err:false,msg:'upload',name:'UpLoadFile/image/maxpic/S" & tempCls.filename & "'}")
		    else
			sFileName = UpLoadDir&strFile
			response.Write("{err:false,msg:'upload',name:'UpLoadFile/image/maxpic/" & tempCls.filename & "'}")
			end if	
		else
		response.Write("{err:true,msg:'" & tempCls.Exception & "'}")
	end if
	set tempCls=nothing
	end if
end if
set Upload=nothing
end if
Function GetExtendName(FileName) 
	Dim ExtName 
	ExtName = LCase(FileName) 
	ExtName = right(ExtName,3) 
	ExtName = right(ExtName,3-Instr(ExtName,".")) 
	GetExtendName = ExtName 
End Function
Function oJpegPic(photopath,thisPic,widthx,heighty)
	Dim oStr,m_File
	oStr = Lcase(GetExtendName(photopath))
	If oStr = "jpg" or oStr = "gif" or  oStr = "png" Then 
		If IsObjInstalled("persits.jpeg") Then
			Set ImagesCls = New ImgWHInfo
			ImagesW = ImagesCls.imgH(Server.Mappath(photopath))
			ImagesH = ImagesCls.imgW(Server.Mappath(photopath))
			Set ImagesCls = Nothing
'			If ImagesW > 600 And ImagesH > 600 Then
           If ImagesW > strPicW or ImagesH > strPicH Then
				Set oJpeg = server.CreateObject("persits.jpeg")
				oJpeg.Open server.mappath(photopath)
				IF ImagesW < ImagesH then '计算宽高比缩放,图片不变形
				oJpeg.Width = widthx
   		    	oJpeg.height =ImagesW/(ImagesH/widthx)
				else
				oJpeg.Width = ImagesH/(ImagesW/heighty)
   		    	oJpeg.height = heighty
				end if
				oJpeg.save server.mappath(thisPic)
				m_File = StartJpeg(thisPic,thisPic)
				oJpegPic = m_File
			Else
				oJpegPic = StartJpeg(photopath,thisPic)
			End If
		Else
			oJpegPic = photopath
		End If
	Else
		oJpegPic = photopath
	End If
	
End Function

Function StartJpeg(CodePic,IsCodePic)
	'================aspJpeg 开始====================
	If (CodePic = "" Or IsNull(CodePic)) Then
		Exit Function
	End If
	If JpegType <> 0 Then
		
		Set BG = Server.CreateObject("Persits.Jpeg")
		BG.Open Server.MapPath(CodePic)
		BG_W = BG.Width
		BG_H = BG.Height
		
		If JpegType = 1 Then
		
		BG.Canvas.Font.Color = "&H" & JpegColor ' 颜色,这里是设置成:黑  
		BG.Canvas.Font.Family = JpegFont' 设置字体  
		BG.Canvas.Font.Bold = true '是否设置成粗体  
		BG.Canvas.Font.Size = JpegSize '字体大小  
		BG.Canvas.Font.Quality = 4 ' 文字清晰度 
		BG.Canvas.Font.ShadowColor = &H000000
		BG.Canvas.Font.ShadowYOffset = 1 
		BG.Canvas.Font.ShadowXOffset = 1
		
	'	BG.Canvas.Font.Interpolation=0 
            

			Select Case JpegLocation
				Case 0
					x = 20 : y = 20
				Case 1
					x = BG_W - Len(JpegTxt) * 20 : y = 20
				Case 2
					x = 20 : y = BG_H - 20
				Case 3
					x = BG_W - Len(JpegTxt) * 20 : y = BG_H - 20*2
				Case 4
					x = BG_W\2 - Len(JpegTxt) * 20\2 : y = BG_H\2 - 20*2
			End Select
			BG.Canvas.PrintText x, y, JpegTxt
		End If
		
		If JpegType = 2 Then
			Set logo = Server.CreateObject("Persits.Jpeg")
			logo.Open Server.MapPath(JpegPic)
			logo_w = logo.Width
			logo_h = logo.Height
			Select Case JpegLocation
				Case 0
					x = 20 : y = 20
				Case 1
					x = BG_W - logo_w - 20 : y = 20
				Case 2
					x = 20 : y = logo_h - 20
				Case 3
					x = BG_W - logo_w - 20 : y = BG_H - logo_h - 20
				Case 4
					x = BG_W\2 - logo_w\2 : y = BG_H\2 - logo_h - 20
			End Select
			'BG.Canvas.drawimage x, y, logo, 0.4, &HFFFFFF  '低版本ASPJPEG请注释下面哪行 启用本行
			BG.Canvas.DrawPNG x, y,Server.MapPath(JpegPic)

			Set logo = Nothing
		End If
		BG.Quality = 100
		BG.Save Server.MapPath(IsCodePic)
		Set BG = Nothing
		StartJpeg = IsCodePic
	End If
	'================aspJpeg 结束====================
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function


rs.close
set rs=nothing
conn.close
set conn=nothing

%>