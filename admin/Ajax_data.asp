<!--#include file="conn.asp"-->
<!--#include file="inc/Function.asp"-->
<%

'**************************************

Select Case request.querystring("action")


	   Case "selectvalue" : selectvalue()
	   Case "getspecification" : getspecification()
	   Case "getspecificationpic" : getspecificationpic()


End Select







Sub selectvalue()
		On Error Resume Next 
		typeid = CheckNum(request("typeid"),0)
		id = CheckNum(request("id"),0)
		mode = CheckNum(request("mode"),0)
		value = Trim(Replace(request("value"),"|",","))
		lang = safereplace(request("lang"))
		If mode = 2 Then 
				If (value = "" Or value = "@@") And typeid <> 0 Then value = "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$@@$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
				pro_name = Split(value,"@@")(0)
				If Split(value,"@@")(1) = "" Then 
						pro_info = Split("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$","$$$")
				Else
						pro_info = Split(Split(value,"@@")(1),"$$$")
				End If 
		Else
				If value = "" Then value = "$$$|||||"
		End If 
		If mode = 2 And typeid = 0 Then 
				If UBound(Split(pro_name,"$")) > 0 Then 
				Response.Write("<div id=""getvalue""><table cellspacing=""0"" cellpadding=""0"" width=""100%"" id=""selectvalue"& lang & mode &""">")
				For ii = 0 To UBound(Split(pro_name,"$"))
				Response.Write("<tr class=""classrow""><td width=""15%"" height=""30"" class=""classrow""><input type=""hidden"" name=""p_ProductValueType"& lang &""" id=""p_ProductValueType"& lang &""" value=""0"" /><input type=""text"" name=""p_ProductName"& lang &""" id=""p_ProductName"& lang &""" size=""10"" value="""& Split(pro_name,"$")(ii) &""" class=""form"" />：</td><td height=""30"" colspan=""2"" class=""classrow""><input type=""text"" name=""p_ProductValue_"& ii & lang &""" id=""p_ProductValue_"& ii & lang &""" size=""25"" value="""& pro_info(ii) &""" class=""form"" /></td></tr>")
				Next 
				Response.Write("</table></div>")
				End If 
		End If 
		IsSelectName = 0
		sql = "select s_name"& lang &",s_id,s_type from [SelectName] where typeid = "& typeid &" and s_mode = "& mode &" order by s_order desc"
		Set rs = conn.execute(sql)
		If Not rs.eof Then 
			IsSelectName = 1
			DataSelectName = rs.getrows()
		End If 
		rs.close : Set rs = Nothing 
		If IsSelectName = 1 Then 
		Response.Write("<div id=""getvalue""><table cellspacing=""0"" cellpadding=""0"" style=""border:none""  width=""100%"" id=""selectvalue"& lang & mode &""">")
		For i = 0 To UBound(DataSelectName,2)
		If i = UBound(DataSelectName,2) Then 
		response.write("<tr ><td width=""15%"" height=""30""  style=""background-color:#faf9d1;border-bottom:none"" >")
		Else
		response.write("<tr ><td width=""15%"" height=""30""  style=""background-color:#faf9d1"" >")
		End If   
		
		If mode = 1 Then response.write("<input type=""checkbox"" name=""Hidden_SName"" value="""& DataSelectName(1,i) &"|"& DataSelectName(0,i) &""""& judge(Split(value,"$$$")(0),DataSelectName(1,i))&" /> ")
			If i = UBound(DataSelectName,2) Then 
		response.write(DataSelectName(0,i) &"：</td><td height=""30"" colspan=""2"" style=""background-color:#faf9d1;border-bottom:none"">")
		Else
		response.write(DataSelectName(0,i) &"：</td><td height=""30"" colspan=""2"" style=""background-color:#faf9d1"">")
		End If 
		
		If mode = 2 Then 
				Response.Write("<input type=""hidden"" name=""p_ProductName"& lang &""" id=""p_ProductName"& lang &""" value="""& DataSelectName(0,i) &""" />")
				If DataSelectName(2,i) = 0 Then 
						Response.Write("<input type=""hidden"" name=""p_ProductValueType"& lang &""" id=""p_ProductValueType"& lang &""" value=""0"" /><select name=""p_ProductValue_"& i & lang &""" id=""p_ProductValue_"& i & lang &"""><option value=""""></option>")
				ElseIf DataSelectName(2,i) = 1 Then 
						Response.Write("<input type=""hidden"" name=""p_ProductValueType"& lang &""" id=""p_ProductValueType"& lang &""" value=""0"" /><input type=""text"" name=""p_ProductValue_"& i & lang &""" id=""p_ProductValue_"& i & lang &""" size=""25"" value="""& pro_info(getN(pro_name,DataSelectName(0,i),"$")) &""" class=""form"" />")
				Else
						Response.Write("<input type=""hidden"" name=""p_ProductValueType"& lang &""" id=""p_ProductValueType"& lang &""" value=""1"" />")
				End If 
		End If 
		IsSelectValue = 0
		sql = "select id,"
		If mode = 2 Then 
				sql = sql &"s_value"& lang
		Else 
				For LangT = 1 To TotalLang
						sql = sql &"s_value"& LanguageTxt(LangT)
						If LangT < TotalLang Then sql = sql &","
				Next 
		End If 
		sql = sql &" from [SelectValue] where s_id = "& DataSelectName(1,i) &" order by s_order desc"
		Set rs = conn.execute(sql)
		If Not rs.eof Then 
			IsSelectValue = 1
			DataSelectValue = rs.getrows()
		End If 
		rs.close : Set rs = Nothing 
		If IsSelectValue = 1 Then 
		For ii = 0 To UBound(DataSelectValue,2)
			Select Case mode 
			Case "0" Response.Write("<input type=""checkbox"" name=""ajax_SelectValue"& lang &""" id=""ajax_SelectValue"& lang &""" value="""& DataSelectName(0,i) &"_"& Trim(DataSelectValue(0,ii)) &""" "& judge(value,DataSelectName(0,i) &"_"& Trim(DataSelectValue(0,ii)))&"> "& Trim(DataSelectValue(1,ii)) &"&nbsp;")
			Case "1" 
			Response.Write("<input type=""checkbox"" name=""SValue"& DataSelectName(1,i) &""" value="""& DataSelectValue(0,ii) &"|")
			For LangT = 1 To TotalLang
					Response.Write(Trim(DataSelectValue(LangT,ii)))
					If LangT < TotalLang Then Response.Write("$$$")
			Next 
			Response.Write(""" "& judge(Split(value,"$$$")(1),Trim(DataSelectValue(0,ii)))&"> "& Trim(DataSelectValue(1,ii)) &"&nbsp;")
			Case "2" 
			If DataSelectName(2,i) = 0 Then Response.Write("<option value="""& Trim(DataSelectValue(1,ii)) &""" "& judge2(pro_info(getN(pro_name,DataSelectName(0,i),"$")),Trim(DataSelectValue(1,ii))) &">"& Trim(DataSelectValue(1,ii)) &"</option>")
			If DataSelectName(2,i) = 2 Then Response.Write("<input type=""checkbox"" name=""p_ProductValue_"& i & lang &""" id=""p_ProductValue_"& i & lang &""" value="""& Trim(DataSelectValue(1,ii)) &""" "& judge(pro_info(getN(pro_name,DataSelectName(0,i),"$")),Trim(DataSelectValue(1,ii))) &" /> "& Trim(DataSelectValue(1,ii)) &"")
			End Select 
		Next 
		End If 
		If mode = 2 And DataSelectName(2,i) = 0 Then 
				Response.Write("</select>")
		End If 
		response.write("</td></tr>")
		Next 
		If mode = 1 Then 
		response.write("<tr><td class=""classrow"" colspan=""3"" style=""border-bottom:none;background-color:#faf9d1""><input type=""button"" name=""btn"" value=""生成"" class=""bnt"" onclick=""getdiv("& id &");"" />&nbsp;&nbsp;<input type=""button"" name=""btn"" value=""重新获取"" class=""bnt"" onclick=""ResetRow("& id &");"" /></td></tr>")
		End If 
		response.write("</table><script>tdshow(""selectvalue"& lang & mode &""");</script></div>")
		End If 
End Sub 

Sub getspecification()
		name = request("name")
		id = checknum(request("id"),0)
		delid = checknum(request("delid"),0)
		list = request("list")
		DataCount = 0
		conn.execute("delete from [Product_Stock] where id = "& delid &"")
		Set rs = conn.execute("select p_no,p_stock,p_price,p_cbprice,p_weight,s_value,id from [Product_Stock] where p_id = "& id &"")
		If Not rs.eof Then 
				DataCount = 1
				DataValue = rs.getrows()
		End If 
		rs.close : Set rs = Nothing
		If DataCount = 1 Then 
				response.write("<table id=""Table_Specification"" cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""border:none; margin-top:10px""><tbody><tr>")
				If InStr(name,",") = 0 Then 
						response.write("<td class='listhead' height='25'>")
						Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(name,0) &"")
						If Not rs.eof Then 
								response.write(rs(0))
						End If 
						rs.close : Set rs = Nothing
						response.write("</td>")
				Else
						For n = 0 To UBound(Split(name,","))
						response.write("<td class='listhead' height='25'>")
						Set rs = conn.execute("select s_name from [SelectName] where s_id = "& checknum(Split(name,",")(n),0) &"")
						If Not rs.eof Then 
								response.write(rs(0))
						End If 
						rs.close : Set rs = Nothing
						response.write("</td>")
						Next 
				End If 
				response.write("<td width='12%' class='listhead'>商品编号</td><td width='12%' class='listhead'>库存</td><td width='12%' class='listhead'>加价</td><td width='12%' class='listhead'>成本价</td>")
				If list = 0 Then response.write("<td width='10%' class='listhead'>移除</td>")
				response.write("</tr>")
				For i = 0 To UBound(DataValue,2)
						response.write("<tr>")
						If InStr(DataValue(5,i),",") = 0 Then 
								response.write("<td class='listrow' height='20'>"& DataValue(5,i) &"</td>")
						Else
								For n = 0 To UBound(Split(DataValue(5,i),","))
								response.write("<td class='listrow' height='20'>"& Split(DataValue(5,i),",")(n) &"</td>")
								Next 
						End If 
						response.write("<td class='listrow' height='20'>")
						If list = 2 Then 
								response.write(DataValue(0,i))
						Else
								response.write("<input style='width:130px;' class='form' type='text' name='ajax_p_no' id='p_no"& i &"' value='"& DataValue(0,i) &"' /><input type='hidden' name='ajax_s_id' id='s_id"& i &"' value='"& DataValue(6,i) &"' /><input type='hidden' name='ajax_s_value' id='s_value"& i &"' value='"& Replace(DataValue(5,i),",","@") &"' />")
						End If 
						response.write("</td><td class='listrow' height='20'>")
						If list = 2 Then 
								response.write(DataValue(1,i))
						Else
								response.write("<input style='width:40px;' class='form' type='text' name='ajax_p_stock' id='p_stock"& i &"' onblur=""countle('"& UBound(DataValue,2) + 1 &"','"& id &"')"" value='"& DataValue(1,i) &"' />")
						End If 
						response.write("</td><td class='listrow' height='20'>")
						If list = 2 Then 
								response.write(DataValue(2,i))
						Else
								response.write("<input style='width:40px;' class='form' type='text' name='ajax_p_price' id='p_price"& i &"' value='"& DataValue(2,i) &"' />")
						End If 
						response.write("</td><td class='listrow' height='20'>")
						If list = 2 Then 
								response.write(DataValue(3,i))
						Else
								response.write("<input style='width:40px;' class='form' type='text' name='ajax_p_cbprice' id='p_cbprice"& i &"' value='"& DataValue(3,i) &"' />")
						End If 
						response.write("</td>")
						If list = 2 Then 
							
						Else
								
						End If 
						response.write("")
						If list = 0 Then response.write("<td class='listrow' height='20'><a style='cursor:pointer;' onclick=""delValueRow('"& id &"','"& name &"','"& DataValue(6,i) &"','"& UBound(DataValue,2) &"');"">删除</a></td>")
						If list = 0 Then response.write("</tr>")
				Next 
				response.write("</tbody></table>")
		End If 
End Sub 


Sub getspecificationpic()
name = Replace(request("name"),"|",",")
		id = checknum(request("id"),0)
		pic = request("pic")
		If name <> "" Then 
				If Right(name,1) = "," Then name = cut_Right(name)
				response.write("<table cellpadding=""0"" cellspacing=""0"" class=""tablelist"">")
				If InStr(name,",") = 0 Then 
						response.write("<tr><td>")
						Set rs = conn.execute("select s_value from [SelectValue] where id = "& checknum(name,0) &"")
						If Not rs.eof Then 
								response.write(rs(0))
						End If 
						rs.close : Set rs = Nothing
						response.write("</td><td><input type=""text"" name=""SValuePic"" id=""SValuePic0"" size=""20"" value="""& pic &""" class=""form"" /></td><td>请手动输入要关联的图片地址</td></tr>")
				Else
						If pic = "" Then pic = "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
						For n = 0 To UBound(Split(name,","))
						response.write("<tr><td>")
						Set rs = conn.execute("select s_value from [SelectValue] where id = "& checknum(Split(name,",")(n),0) &"")
						If Not rs.eof Then 
								response.write(rs(0))
						End If 
						rs.close : Set rs = Nothing
						response.write("</td><td><input type=""text"" name=""SValuePic"" id=""SValuePic"& n &""" size=""20"" value="""& Split(pic,"|")(n) &""" class=""form"" />")
						If Split(pic,"|")(n) <> "" Then 
					
						End If 
						response.write("</td><td>请手动输入要关联的图片地址</td></tr>")
						Next 
				End If 
				response.write("</table>")
		End If 
End Sub 


%>