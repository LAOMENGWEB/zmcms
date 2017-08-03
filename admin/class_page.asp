<%

Const PW_BTN_F = "首页"

Const PW_BTN_P = "上一页"

Const PW_BTN_N = "下一页"

Const PW_BTN_L = "尾页"

Const PW_BTN_NL_CSS = ""
Const PW_BTN_L_CSS = ""
'================================================
'数字导航按钮样式
'================================================
Const PN_BTN_F = "首页"

Const PN_BTN_P = "上一页"

Const PN_BTN_N = "下一页"

Const PN_BTN_L = "尾页"

Const PN_BTN_NL_CSS = " class='n_nl'"
Const PN_BTN_L_CSS = " class='n_l'"
'================================================

Class HHYY_Pageclass

	Private CurrPage
	Private Int_TotalRecord
	Private Int_TotalPage
	
	Private Page_RS
	Private Page_Listnum
	Private Page_Pathurl
	Private Page_URLstr
	Private Page_Navtype
	Private Error_Str

	Private Sub Class_initialize()
		CurrPage = 1
		Page_Listnum = 10
		Page_Navtype = 1
		Error_Str = ""
	End Sub
	
	Private Sub Class_Terminate()
		Error_Str = ""
		'Page_RS.Close
		Set Page_RS = Nothing
	End Sub
	
	Public Property Let Setpage_RS(ByRef obj_RS)
		If IsObject(obj_RS) Then
			Set Page_RS = obj_RS
		End If
	End Property
	
	'================================================
	'设置每页显示的信息数
	'================================================
	Public Property Let Setpage_Listnum(ByVal P_LN)
		If P_LN = "" Or IsEmpty(P_LN) Then
			Page_Listnum = 10
		Else
			If IsNumeric(P_LN) = False Then
				Error_Str = "不可识别的信息显示数!"
				'Call Show_Error(Error_Str)
			Else
				Page_Listnum = Clng(P_LN)
			End If
		End If
	End Property
	
	'================================================
	'设置页面转向的文件地址
	'================================================
	Public Property Let Setpage_Pathurl(str_URL)
		Page_Pathurl = Cstr(str_URL)
	End Property
	
	'================================================
	'设置分页导航样式
	'================================================
	Public Property Let Setpage_Navtype(ByVal P_NT)
		Page_Navtype = P_NT
	End Property
	
	'================================================
	'输出错误信息
	'================================================
	Public Property Get OutErrInfo()
		OutErrInfo = Error_Str
	End Property
	
	'================================================
	'分页初始化
	'================================================
	Public Sub Page_Init()
		
		CurrPage = Replace(Replace(Trim(Request("page")),"'",""),"%","")
		
		If CurrPage = "" Or IsEmpty(CurrPage) Then
			CurrPage = 1
		Else
			If Check_Digital(CurrPage) = False Then
				CurrPage = 1
			Else
				If CurrPage < 1 Or CurrPage <= 0 Then
					CurrPage = 1
				Else
					CurrPage = Clng(CurrPage)
				End If
			End If
		End If
		
		Int_TotalRecord = Page_RS.RecordCount
		Page_RS.PageSize = Page_Listnum
		
		If Int_TotalRecord >=1 Then
			Int_TotalPage = Page_RS.PageCount
			If (CurrPage-Int_TotalPage) >0 Then
				CurrPage = Int_TotalPage
			Else
				CurrPage = Clng(CurrPage)
			End If
		Else
			CurrPage = 1
		End If
		
		Page_RS.AbsolutePage = CurrPage
		
	End Sub
	
	'================================================
	'显示分页导航(任意位置调用)
	'================================================
	Public Sub Page_Navigation()
		Dim Nav_Htmlstr
		Page_URLstr = Format_URL(Page_Pathurl)
		'显示分页信息，各个模块根据自己要求更改显求位置
		Response.Write "<table width=""99%""align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0""style=""border-top:none"">"
		Response.Write "<tr >"
		Response.Write "<td align=""center""  style=""border-bottom:none;padding:10px 0;border-right:none"">"
		
		Nav_Htmlstr = Show_FirstPrv
		Response.Write Nav_Htmlstr &"&nbsp;"
		
		If Page_Navtype = 2 Then
			Nav_Htmlstr = Show_NumBtn
			Response.Write Nav_Htmlstr
		End If
		
		Nav_Htmlstr = Show_NextLast
		Response.Write Nav_Htmlstr
		
		Nav_Htmlstr = Show_PageInfo
		Response.Write "&nbsp;"& Nav_Htmlstr
		
		Nav_Htmlstr = Show_PageOption
		Response.Write Nav_Htmlstr
		
		Response.Write "&nbsp;</td>"
		Response.Write "</tr>"
		Response.Write "</table>"
	End Sub
	
	
	'/*显示首页、前一页*/
	Private Function Show_FirstPrv()
		Dim FP_BTN_F, FP_BTN_P, FP_BTN_NL_CSS, FP_BTN_L_CSS
		Dim FP_STR, INT_PPAGE
		
		If Page_Navtype = "" Or IsEmpty(Page_Navtype) Then
			FP_BTN_F = PW_BTN_F
			FP_BTN_P = PW_BTN_P
			FP_BTN_NL_CSS = PW_BTN_NL_CSS
			FP_BTN_L_CSS = PW_BTN_L_CSS
		Else
			Select Case Page_Navtype
				Case 1
					FP_BTN_F = PW_BTN_F
					FP_BTN_P = PW_BTN_P
					FP_BTN_NL_CSS = PW_BTN_NL_CSS
					FP_BTN_L_CSS = PW_BTN_L_CSS
				Case Else
					FP_BTN_F = PN_BTN_F
					FP_BTN_P = PN_BTN_P
					FP_BTN_NL_CSS = PN_BTN_NL_CSS
					FP_BTN_L_CSS = PN_BTN_L_CSS
			End Select
		End If
		
		If CurrPage <= 1 Then 
			FP_STR = "<span"& FP_BTN_NL_CSS &">"& FP_BTN_F &"</span>&nbsp;<span"& FP_BTN_NL_CSS &">"& FP_BTN_P &"</span>"
		Else 
			INT_PPAGE = CurrPage-1 
			FP_STR = "<a"& FP_BTN_L_CSS &" href='"& Page_URLstr &"page=1'>"& FP_BTN_F &"</a>&nbsp;<a"& FP_BTN_L_CSS &" href='"& Page_URLstr &"page="& Cstr(INT_PPAGE) &"'>" & FP_BTN_P &"</a>"
		End If
		Show_FirstPrv = FP_STR
	End Function
	
	'/*显示下一页、末页*/
	Private Function Show_NextLast()
		Dim NL_BTN_N, NL_BTN_L, NL_BTN_NL_CSS, NL_BTN_L_CSS
		Dim NL_STR, INT_NPAGE
		
		If Page_Navtype = "" Or IsEmpty(Page_Navtype) Then
			NL_BTN_N = PW_BTN_N
			NL_BTN_L = PW_BTN_L
			NL_BTN_NL_CSS = PW_BTN_NL_CSS
			NL_BTN_L_CSS = PW_BTN_L_CSS
		Else
			Select Case Page_Navtype
				Case 1
					NL_BTN_N = PW_BTN_N
					NL_BTN_L = PW_BTN_L
					NL_BTN_NL_CSS = PW_BTN_NL_CSS
					NL_BTN_L_CSS = PW_BTN_L_CSS
				Case Else
					NL_BTN_N = PN_BTN_N
					NL_BTN_L = PN_BTN_L
					NL_BTN_NL_CSS = PN_BTN_NL_CSS
					NL_BTN_L_CSS = PN_BTN_L_CSS
			End Select
		End If
		
		If CurrPage >= Int_TotalPage Then
			NL_STR = "<span"& NL_BTN_NL_CSS &">"& NL_BTN_N &"</span>&nbsp;<span"& NL_BTN_NL_CSS &">"& NL_BTN_L &"</span>"
		Else 
			INT_NPAGE = CurrPage+1 
			NL_STR = "<a"& NL_BTN_L_CSS &" href='"& Page_URLstr &"page="& Cstr(INT_NPAGE) &"'>"& NL_BTN_N &"</a>&nbsp;<a"& NL_BTN_L_CSS &" href='"& Page_URLstr &"page="& Cstr(Int_TotalPage) &"'>" & NL_BTN_L &"</a>"
		End If
		
		Show_NextLast = NL_STR
	End Function
	
	'/*显示数字导航*/
	Private Function Show_NumBtn()
		
		Dim str_tmp, P_I, EndPage
		
		If Int_TotalPage > CurrPage+4 Then
			EndPage = CurrPage+4
		Else
			EndPage = Int_TotalPage
		End If
		
		For P_I = CurrPage-3 to EndPage
			If Not P_I < 1 Then
				If P_I = CurrPage Then
					str_tmp = str_tmp &"<span style='font-weight:bold; color:#FF0000; font-size:14px;'>["& P_I &"]</span>&nbsp;"
				Else
					str_tmp = str_tmp &"<a href='"& Page_URLstr &"page="& Cstr(P_I) &"' style='font-size:14px;'>["& P_I &"]</a>&nbsp;"
				End If
			End If
		Next
		
		If CurrPage+4 < Int_TotalPage Then
			str_tmp = str_tmp &"....."
			str_tmp = str_tmp &"<a href='"& Page_URLstr &"page="& Cstr(Int_TotalPage) &"' style='font-size:14px;'>["& Int_TotalPage &"]</a>&nbsp;"
		End If
		
		Show_NumBtn = str_tmp
	End Function
	
		
	'/*显示跳转*/
	Private Function Show_PageOption()
		Dim str_tmp, i
		str_tmp = vbcrlf
		str_tmp = str_tmp & "<script language=""javascript"" type=""text/javascript"">"& vbcrlf
		str_tmp = str_tmp & "function page_jump(){"& vbcrlf
		str_tmp = str_tmp & "var kg = /^[\s　]*|[\s　]*$/;"& vbcrlf
		str_tmp = str_tmp & "var pn_Value = document.getElementById(""pn"").value;"& vbcrlf
		str_tmp = str_tmp & "pn_Value = pn_Value.replace(kg,"""");"& vbcrlf
		str_tmp = str_tmp & "var PN_Regex = /^\d+$/;"& vbcrlf

		str_tmp = str_tmp & "if(pn_Value == """"){"& vbcrlf
		str_tmp = str_tmp & "alert(""请输入页码!"");"& vbcrlf
		str_tmp = str_tmp & "document.getElementById(""pn"").focus();"& vbcrlf
		str_tmp = str_tmp & "}"& vbcrlf
		str_tmp = str_tmp & "else{"& vbcrlf
		str_tmp = str_tmp & "if(!PN_Regex.test(pn_Value)){"& vbcrlf
		str_tmp = str_tmp & "alert(""页码输入错误!"");"& vbcrlf
		str_tmp = str_tmp & "document.getElementById(""pn"").focus();"& vbcrlf
		str_tmp = str_tmp & "}"& vbcrlf
		str_tmp = str_tmp & "else{"& vbcrlf
		str_tmp = str_tmp & "window.location.href="""& Page_URLstr &"page=""+pn_Value ;"& vbcrlf
		str_tmp = str_tmp & "}"& vbcrlf
		str_tmp = str_tmp & "}"& vbcrlf

		str_tmp = str_tmp & "}"& vbcrlf
		str_tmp = str_tmp & "</script>"& vbcrlf
		str_tmp = str_tmp & "&nbsp;<input type='text' name='pn' id='pn' style='height:20px;line-height:18px;' SIZE=1 Maxlength=5 value='"& CurrPage &"'>"
		str_tmp = str_tmp & "&nbsp;<input type='button' style='height:20px;line-height:18px; border:none' value='跳转' onClick=""page_jump()"">"
		
		Show_PageOption = str_tmp
	End Function
	
	'/*分页信息*/
	'/*根据要求自行修改*/
	Private Function Show_PageInfo() 
		Dim str_tmp
		str_tmp = "页次：<font color='#FF0000'><b>"& CurrPage &"</b></font>/"& Int_TotalPage &"页&nbsp;共"& Int_TotalRecord &"条&nbsp;"& Page_Listnum &"条/页" 
		Show_PageInfo = str_tmp 
	End Function
	
	Private Function Format_URL(strUrl)
		If strUrl = "" Or IsEmpty(strUrl) Then
			Format_URL = ""
			Exit Function
		End If
		
		If InStr(strUrl,"?") < Len(strUrl) Then 
			If InStr(strUrl,"?") > 1 Then
				If InStr(strUrl,"&") < Len(strUrl) Then 
					Format_URL = strUrl & "&"
				Else
					Format_URL = strUrl
				End If
			Else
				Format_URL = strUrl & "?"
			End If
		Else
			Format_URL = strUrl
		End If
	End Function
	
	Private Function Check_Digital(tmp_str)
		If tmp_str = "" Or IsEmpty(tmp_str) Then
			Check_Digital = False
			Exit Function
		End If
		
		Dim obj_regEx, Results
		Set obj_regEx = New RegExp
		obj_regEx.Pattern = "^\d+$"
		obj_regEx.IgnoreCase = True
		obj_regEx.Global = True
		Results = obj_regEx.Test(tmp_str)
		Check_Digital = tmp_str
	End Function
	
	Private Sub Show_Error(STR)
		If STR <> "" Then
			Response.Write STR
			Response.End
		End If
	End Sub
	
End Class
%>