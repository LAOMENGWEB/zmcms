<%
'
'
'///////////////////////////////////////////////////////////////////////////////////
'调用及说明
'///////////////////////////////////////////////////////////////////////////////////
''创建类的实例
'Set Page = new Page_List
''获得数据库对象
'Page.Con = Conn
''获得SQL查询语句
'Page.Sql = "Select * from [TableName] Where ABC = '" & ABC & "'"
''每页显示记录条数，可选属性，默认为 10
'Page.PageSize = 20
''分页的样式(1~6)，各样式分别见内部代码说明，可选属性，默认为 5
'Page.Style = 5
''计数单位，可选属性，默认为 条记录
'Page.Unit = "条记录"
''跳转的URL地址，默认为当前页，可自定所传字符串，如 Page.PageUrl = "Search.asp?keyword=" & KeyWord
'Page.PageUrl = Url
''获得记录集对象
'Set Rs = Page.Rs
'
'If Rs.Bof And Rs.Eof Then
'    Response.Write "没有数据"
'Else
'    For i = 1 To Page.PageSize     '开始循环输出，用户可以自定义输出方式，不用判断是否到达记录集末尾，类里已经处理
'        Response.Write i     '显示数据
'        Rs.MoveNext     '移动记录集指针
'    Next    '继续循环直到结束
'End If
'
''显示分页页码
'Call Page.ShowPage
'
''断开并清空记录集
'Rs.Close
'Set Rs = Nothing
''销毁分页类
'Set Page = Nothing
'///////////////////////////////////////////////////////////////////////////////////




'分页类开始
Class Page_List
    '定义私有变量
    Private PageCount, myCon, myRs, mySQL, CurrentPage, TotalNumber, sURL, URL, FileName, myPageSize, myStyle, sUnit
	
	Private Function NewURL()
		sURL = Request.ServerVariables("URL")
		sQUERY_STRING = Request.ServerVariables("QUERY_STRING")
		sTmp = Split(sURL,"/")
		sURL = sTmp(Ubound(sTmp))
		If sQUERY_STRING <> "" Then	sQUERY_STRING=Replace(sQUERY_STRING,"page=" & Request("page"),"")
		If sQUERY_STRING = "" Then
			sURL = sURL & "?"
			Else
				sURL = sURL & "?" & sQUERY_STRING & "&"
		End If
		sURL = Replace(sURL,"&&","&")
		NewURL=sURL
	End Function

    '显示页面信息，可以根据需要修改
    Private Function ShowPageInfo()
        Dim str_tmp
        str_tmp = "共 <font color=""#ff0000""><b>" & TotalNumber & "</b></font> " & sUnit & " 每页 <font color=""#FF0000""><b>" & myPageSize & "</b></font> " & sUnit & " 页次：<font color=""#ff0000""><b>" & CurrentPage & "</b></font>/<b>" & PageCount & "</b>页"
        ShowPageInfo = str_tmp
    End Function

     '页码跳转输入框
    Private Function Page_Input()
        Page_Input = " <input id=""page"" type=""text"" value=""" & CurrentPage & """ style=""font-family: Verdana; font-size: 10px; height: 15px; border: 1px solid #666666"" size=""2"" title=""定页"" onmouseover=""javascript:this.focus();this.select()"" onKeyDown=""javascript: if(window.event.keyCode == 13) location='" & URL & "page='+this.value"">"
    End Function

    '页码跳转Go按钮
    Private Function Page_Button()
        Page_Button = " <input type=""button"" style=""font-family: Verdana; font-size: 10px; height: 15px; border: 1px solid #666666"" value=""Go"" onclick=""javascript:if(document.getElementById('page')) location='" & URL & "page='+document.getElementById('page').value"">"
    End Function

    '页码跳转选择列表
    Private Function Page_Select()
        Dim str_tmp, i
        str_tmp = " 转到：<select id=""page"" onchange=""javascript:location='" & URL & "page='+this.options[this.selectedIndex].value"" style=""font-size:12px; font-family:Verdana"">"
        For i = 1 To PageCount
            str_tmp = str_tmp & "<option value=""" & i & """"
            If i = CurrentPage Then str_tmp = str_tmp & " selected"
            str_tmp = str_tmp & ">第" & i & "页</option>"
        Next
        str_tmp = str_tmp & "</select>"
        Page_Select = str_tmp
    End Function

    '显示翻页（myStyle表示导航栏样式1~5）
    Public Sub ShowPage()
        Dim i, str
        If TotalNumber > 0 Then
            str = "<table width=""99%"" cellspacing=""0"" cellpadding=""0"" border=""0"" align=""center"" id=""SplitPage"" style=""border:0px #ddd solid; font-size:12px; margin-top:10px""><tr>"
            str = str & "<td align=""center"">" & ShowPageInfo & " "
            If PageCount >= 1 Then
                '调用各个分页样式
                Select Case myStyle
                    Case 1: str = str & PageStyle1
                    Case 2: str = str & PageStyle2
                    Case 3: str = str & PageStyle3
                    Case 4: str = str & PageStyle4
                    Case 5: str = str & PageStyle5
                    Case 6: str = str & PageStyle6
                End Select
            End If
            str = str & "</td></tr></table>"
        Else
            str = " "
        End If
        Response.Write str
    End Sub

    '数据显示页面通过该函数取得Rs对象，然后自定义显示方式
    Public Function Rs()
        On Error Resume Next
        Set myRs = Server.CreateObject("ADODB.RecordSet")
        myRs.PageSize = myPageSize
        myRs.Open mySQL, myCon, 1
        If Err Then
            Response.Write Err.Description & "<br>"
            Response.Write "数据库查询出错，请检查查询字符串！"
            Response.End
        End If
        If Not myRs.EOF Then
            '记录总数
            TotalNumber = myRs.RecordCount
            '页面总数
            PageCount = myRs.PageCount
            '当前输出页码
            If CurrentPage > PageCount Then CurrentPage = PageCount
            myRs.AbsolutePage = CurrentPage
        End If
        Set Rs = myRs
    End Function

    '得到数据库连接对象
    Public Property Let Con(oCon)
        Set myCon = oCon
    End Property

    '得到查询语句
    Public Property Let SQL(str)
        mySQL = str
    End Property

    '得到分页的样式(1~6)
    Public Property Let Style(num)
        If IsNumeric(num) Then
            myStyle = CInt(num)
            If myStyle < 1 or myStyle > 6 Then myStyle = 1
        Else
            myStyle = 1
        End If
    End Property

    '每一页的记录条数，页传递给类
    Public Property Let PageSize(num)
        If IsNumeric(num) Then
            myPageSize = CInt(num)
            If myPageSize < 1 Then myPageSize = 10
        Else
            myPageSize = 10
        End If
    End Property

    '计数单位，如：个、条等等，页传递给类
    Public Property Let Unit(str)
        If str = "" Then
            sUnit = "条记录"
        Else
            sUnit = str
        End If
    End Property

    '得到页面URL信息，页传递给类
    Public Property Let PageUrl(str)
        If str <> "" Then
            sURL = str
            URL = JoinChar(sURL)
        Else
            URL = JoinChar(NewURL)
        End If
    End Property

    '数据显示页面中每一页输出的条数，类传递给页
    Public Property Get PageSize()
        '确定每页输出的条数，照顾到最后一页记录数不足的情况
        If CurrentPage = PageCount Then
            PageSize = TotalNumber - myPageSize * (PageCount - 1)
        ElseIf TotalNumber = 0 Then
            PageSize = 0
        Else
            PageSize = myPageSize
        End If
    End Property

    '类的初始化
    Private Sub Class_Initialize()
        '初始化每页显示记录条数为10
        myPageSize = 10
        '初始化分页样式为1
        myStyle = 6
        '初始化记数单位为 条记录
        sUnit = "条记录"
        '处理传入的Page参数
        CurrentPage = Trim(Request.QueryString("Page"))
        If IsNumeric(CurrentPage) Then
            CurrentPage = CInt(CurrentPage)
            If CurrentPage < 1 Then CurrentPage = 1
        Else
            CurrentPage = 1
        End If
        '初始化获得页面URL
        FileName = Trim(Request.ServerVariables("URL"))
        FileName = Mid(FileName, InStrRev(FileName, "/") + 1)
        URL = JoinChar(NewURL)
    End Sub

    '每一条记录的序号，跟页码相关，例如第2页第一条是在第一页最后序号基础上加1
    Public Function ListNum(num)
        ListNum = myPageSize * (CurrentPage - 1) + num
    End Function

    '处理获得到的URL地址
    Public Function JoinChar(ByVal strUrl)
        If strUrl = "" Then
            JoinChar = ""
            Exit Function
        End If
        If InStr(strUrl, "?") < Len(strUrl) Then
            If InStr(strUrl, "?") > 1 Then
                If InStr(strUrl, "&") < Len(strUrl) Then
                    JoinChar = strUrl & "&"
                Else
                    JoinChar = strUrl
                End If
            Else
                JoinChar = strUrl & "?"
            End If
        Else
            JoinChar = strUrl
        End If
    End Function

    '分页样式1：列出所有页码的页号，页码较少时可以使用
    Private Function PageStyle1()
        Dim str
        str = "转到 第 "
        For i = 1 To PageCount
            If i = CurrentPage Then str = str & "<b>" & i & "</b> " Else str = str & "<a href=""" & URL & "page=" & i & """>" & i & "</a> "
        Next
        str = str & "页"
        PageStyle1 = str
    End Function

    '分页样式2：显示首页、上一页、下一页、末页，均为符号和转到页的输入框
    Private Function PageStyle2()
        Dim str
        If CurrentPage > 1 Then
            str = str & "<a href=""" & URL & "page=1"" title=""首页"" style=""text-decoration: none""><font face=""webdings"">9</font></a> "
            str = str & "<a href=""" & URL & "page=" & CurrentPage - 1 & """ title=""上一页"" style=""text-decoration: none""><font face=""webdings"">3</font></a> "
        Else
            str = str & "<font face=""webdings"">9</font> <font face=""webdings"">3</font> "
        End If
        If CurrentPage < PageCount Then
            str = str & "<a href=""" & URL & "page=" & CurrentPage + 1 & """ title=""下一页"" style=""text-decoration: none""><font face=""webdings"">4</font></a> "
            str = str & "<a href=""" & URL & "page=" & PageCount & """ title=""末页"" style=""text-decoration: none""><font face=""webdings"">:</font></a>"
        Else
            str = str & "<font face=""webdings"">4</font> <font face=""webdings"">:</font>"
        End If
        '显示页码输入框，如果不显示则注释掉
        str = str & Page_Input
        '显示Go字样的按钮，如果不显示则注释掉
        str = str & Page_Button
        PageStyle2 = str
    End Function

    '分页样式3：显示首页、上十页、下十页、尾页以及中间一段页码
    Private Function PageStyle3()
        Dim str, p
        p = (CurrentPage - 1) \ 10
        '首页
        If CurrentPage = 1 Then str = str & "<font face=""webdings"">9</font> " Else str = str & "<a href=""" & URL & "page=1"" title=""首页""><font face=""webdings"">9</font></a> "
        '上十页
        If p * 10 > 0 Then str = str & "<a href=""" & URL & "page=" & CurrentPage - 10 & """ title=""上十页""><font face=""webdings"">7</font></a> "
        '页码
        For i = (p * 10 + 1) To (p * 10 + 10)
            If i = CurrentPage Then str = str & "<b>" & i & "</b> " Else str = str & "<a href=""" & URL & "page=" & i & """ title=""转到第" & i & "页"">" & i & "</a> "
            If i = PageCount Then Exit For
        Next
        '下十页
        If i < PageCount Then str = str & "<a href=""" & URL & "page=" & i & """ title=""下十页""><font face=""webdings"">8</font></a> "
        '尾页
        If CurrentPage = PageCount Then str = str & "<font face=""webdings"">:</font>" Else str = str & "<a href=""" & URL & "page=" & PageCount & """ title=""末页""><font face=""webdings"">:</font></a>"
        PageStyle3 = str
    End Function

    '分页样式4：显示首页和尾页的页码、中间显示一部分页码
    Private Function PageStyle4()
        Dim str, n, pageto, i
        If CurrentPage > 4 Then str = str & "<a href=""" & URL & "page=1"">[1]</a> ..."
        If CurrentPage < PageCount - 3 Then pageto = CurrentPage + 3 Else pageto = PageCount
        For i = CurrentPage - 3 To pageto
            If i >= 1 Then
                If i = CurrentPage Then str = str & " <b>[" & i & "]</b>" Else str = str & " <a href=""" & URL & "page=" & i & """>[" & i & "]</a>"
            End If
        Next
        If CurrentPage < PageCount - 3 Then str = str & " ... <a href=""" & URL & "page=" & PageCount & """>[" & PageCount & "]</a>"
        PageStyle4 = str
    End Function

    '分页样式5：显示首页、上一页、下一页、末页，均为汉字和转到页的选择列表
    Private Function PageStyle5()
        Dim str
        If CurrentPage > 1 Then
            str = str & "<a href=""" & URL & "page=1"" title=""首页"">首页</a> "
            str = str & "<a href=""" & URL & "page=" & CurrentPage - 1 & """ title=""上一页"">上一页</a> "
        Else
            str = str & "<font color=""#999999"">首页 上一页</font>"
        End If
        If CurrentPage < PageCount Then
            str = str & " <a href=""" & URL & "page=" & CurrentPage + 1 & """ title=""下一页"">下一页</a> "
            str = str & "<a href=""" & URL & "page=" & PageCount & """ title=""末页"">末页</a> "
        Else
            str = str & " <font color=""#999999"">下一页 末页</font>"
        End If
        '显示页码的选择列表，如果不显示则注释掉
        str = str & Page_Select
        PageStyle5 = str
    End Function
	
    '分页样式6：显示首页、上一页、下一页、末页，均为汉字和转到页的选择列表，转到页的输入框
    Private Function PageStyle6()
        Dim str, n, pageto, i
        If CurrentPage > 4 Then str = str & "<a href=""" & URL & "page=1"">[1]</a> ..."
        If CurrentPage < PageCount - 3 Then pageto = CurrentPage + 3 Else pageto = PageCount
        For i = CurrentPage - 3 To pageto
            If i >= 1 Then
                If i = CurrentPage Then str = str & " <b>[" & i & "]</b>" Else str = str & " <a href=""" & URL & "page=" & i & """>[" & i & "]</a>"
            End If
        Next
        If CurrentPage < PageCount - 3 Then str = str & " ... <a href=""" & URL & "page=" & PageCount & """>[" & PageCount & "]</a>"

        '显示页码输入框，如果不显示则注释掉
        str = str & Page_Input
        '显示Go字样的按钮，如果不显示则注释掉
        str = str & Page_Button
        PageStyle6 = str
    End Function
'分页类结束
End Class
%>