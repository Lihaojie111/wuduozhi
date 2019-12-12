<!--#include file="Conn.asp"-->
<%

Dim Content,Action,ClassID,i,Query_String,Http_Referer,Script_Name,ErrContent,UpFilePath,EditorPath,Layout,Page,Line,arrShowLine(20)
Dim m_Config,m_WebSet,m_MSvr,m_Word,m_OnLine,m_Temp,m_Temp2,CreateFilePath

MyDbPath = "../"
UpFilePath = "../Images/UpFile/"
EditorPath = "../Editor/"
CreateFilePath = "..\Files"

Query_String = Request.ServerVariables("QUERY_STRING")
Http_Referer = Request.ServerVariables("HTTP_REFERER")
Script_Name = Request.ServerVariables("SCRIPT_NAME")
Action = Trim(Request("Action"))
ClassID = Request("ClassID")
If IsNumeric(ClassID) Then
	ClassID = Cint(ClassID)
Else
	Msg "分类编号必须是整数","",1
End If
Page = Request.QueryString("Page")
Line = 20
if Page<1 then Page=1

ConnectionDatabase()

Sql="Select Config,WebSet,MailSvr,Word,OnLine From Config"
Rs.Open Sql,Conn,1,1
If Not Rs.Eof Then
	m_Config=Split(Rs("Config"),"|")
	m_WebSet=Split(Rs("WebSet"),"|")
	m_MSvr=Split(Rs("MailSvr"),"|")
	m_Word=Rs("Word")
	m_OnLine=Rs("OnLine")
Else
	Response.Redirect("Install.asp")
End If
Rs.Close

If LCase(Right(Script_Name,9))<>"login.asp" Then
	Dim ComeUrl,cUrl
	ComeUrl=Lcase(Trim(Request.ServerVariables("HTTP_REFERER")))
	If ComeUrl="" Then
		Session("m_Admin")=""
	Else
		cUrl=Trim("http://" & Request.ServerVariables("SERVER_NAME"))
		If Mid(ComeUrl,Len(cUrl)+1,1)=":" Then
			cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		cUrl=Lcase(cUrl & Request.ServerVariables("SCRIPT_NAME"))
		If Lcase(Left(ComeUrl,Instrrev(ComeUrl,"/")))<>Lcase(Left(cUrl,Instrrev(cUrl,"/"))) Then
			Session("m_Admin")=""
		End If
	End If
End If

If Session("m_Admin")="" And LCase(Right(Script_Name,9))<>"login.asp" Then
	CloseConn()
	Response.Redirect("Login.asp")
	Response.End
End If

Sub Head()
	Content = ""
	Content = Content & "<html>"
	Content = Content & "<head>"
	Content = Content & "<title>"&m_Config(0)&"后台管理系统</title>"
	Content = Content & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
	Content = Content & "<link href=""Style.css"" type=""text/css"" rel=""stylesheet"">"
	Content = Content & "</head>"
	Content = Content & "<body>"
'	Content = Content & "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""t_body"">"
'	Content = Content & "  <tr>"
'	Content = Content & "    <td><table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">"
'	Content = Content & "      <tr>"
'	Content = Content & "        <td width=""20""><img src=""images/i_home.gif"" width=""16"" height=""18""></td>"
'	Content = Content & "        <td width=""100"">后台管理系统</td>"
'	Content = Content & "        <td width=""150"" align=""center""><a href=""Main.asp?Action=ModPsw"" target=""m_Main"">修改管理员密码</a></td>"
'	If m_WebSet(1)="0" Then
'		Content = Content & "        <td align=""center"" width=""150""><marquee width=""100%"" behavior=""alternate"" onmouseover=""this.stop()"" onmouseout=""this.start()"" scrollAmount=""5""><font color=""#CC3300""><strong><u>网站关闭</u></strong></font></marquee></td>"
'	End If
'	Content = Content & "        <td>&nbsp;</td>"
'	Content = Content & "        <td align=""right""><a href="""&m_Config(4)&""" target=""_blank"">返回&nbsp;"&m_Config(0)&"&nbsp;首页</a></td>"
'	Content = Content & "      </tr>"
'	Content = Content & "    </table></td>"
'	Content = Content & "  </tr>"
'	Content = Content & "</table>"
	If m_WebSet(1)=0 Then
		Content = Content & "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""t_border""><tr><td height=""25"" bgcolor=""#FFFFFF""><marquee direction=""left"" style=""font-weight:bold""><span style=""color:#CC3300"">网站维护中...</span>&nbsp;&nbsp;&nbsp;&nbsp; <a href=""Setting.asp?Action=WebSet"">设置状态</a></marquee></td></tr></table>"
	End If
	Response.Write(Content)
End Sub

Sub Bottom()
	CloseConn()
	Content=""
	Content = Content & "</body>"
	Content = Content & "</html>"
	Response.Write(Content)
End Sub

Sub Msg(Message,Url,Method)
	Content=""
	Content = Content & "<form>"
	Content = Content & "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""t_border"">"
	Content = Content & "<tr>"
	Content = Content & "<td class=""t_title"">"
	If Method=0 Then
		Content = Content & "成功"
	ElseIf Method=1 Then
		Content = Content & "错误"
	End If
	Content = Content & "信息</td>"
	Content = Content & "</tr>"
	Content = Content & "<tr>"
	Content = Content & "<td class=""t_rowh"" style=""padding:10px;line-height:120%"">"
	Message=Split(Message,"!")
	For i=0 To Ubound(Message)-1
		Content = Content & "<li>"&Message(i)&"</li>"
	Next
	Content = Content & "</td>"
	Content = Content & "</tr>"
	Content = Content & "<tr>"
	Content = Content & "<td align=""center"" class=""t_row""><a href="""
	If Method=0 Then
		Content = Content & Url
	ElseIf Method=1 Then
		Content = Content & "javascript:history.back()"
	End If
	Content = Content & """>返回</a></td>"
	Content = Content & "</tr>"
	Content = Content & "</table>"
	Content = Content & "</form>"
	Response.Write(Content)
	Bottom()
	Response.End
End Sub

Sub ShowContent(ID,Title,ClassID,Hits,DateAndTime,t_Content,Action)
	m_Temp = Replace(UpFilePath,MyDbPath,"")
	t_Content = Replace(t_Content,m_Temp,UpFilePath)
	m_Temp = Replace(EditorPath,MyDbPath,"")
	t_Content = Replace(t_Content,m_Temp,EditorPath)
	Content = ""
	Content = Content & "<table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"" style=""border:5px #EEEEEE solid""><tr>"
	Content = Content & "<td bgcolor=""#FCFCFC"" style=""border:1px #CFCFCF solid;padding:20px""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""5"">"
	Content = Content & "<tr><td align=""center"" style=""padding:10px;font-size:16px;font-weight:bold"">"& Title &"</td></tr>"
	Content = Content & "<tr><td align=""center"">所属分类："& ShowClassName(ClassID) &"&nbsp;&nbsp;点击次数："& Hits &"次&nbsp;&nbsp;更新时间："& DateAndTime &"</td></tr>"
	Content = Content & "<tr><td><hr></td></tr>"
	Content = Content & "<tr><td>"& t_Content &"</td></tr>"
	Content = Content & "<tr><td><hr></td></tr>"
	Content = Content & "<tr><td style=""line-height:200%"">"& ShowPNRecord(ID,ClassID,Action) &"</td></tr>"
	Content = Content & "</table></td></tr></table>"
	Response.Write(Content)
End Sub

Function ShowPNRecord(ID,ClassID,Action)
	ShowPNRecord = ""
	ShowPNRecord = ShowPNRecord & "<li>上一条记录："
	Sql="Select ID,Title From "& Action &" Where "
	If ClassID>0 Then Sql = Sql & "ClassID="& ClassID &" And "
	Sql = Sql & "ID < "& ID &" Order By ID Desc"
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		ShowPNRecord = ShowPNRecord & "没有记录"
	Else
		ShowPNRecord = ShowPNRecord & "<a href=""?Action=Show&ID="& Rs("ID") &""">" & Rs("Title") &"</a>"
	End If
	Rs.Close
	ShowPNRecord = ShowPNRecord & "</li>"

	ShowPNRecord = ShowPNRecord & "<li>下一条记录："
	Sql="Select ID,Title From "& Action &" Where "
	If ClassID>0 Then Sql = Sql & "ClassID="& ClassID &" And "
	Sql = Sql & "ID > "& ID &" Order By ID Asc"
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		ShowPNRecord = ShowPNRecord & "没有记录"
	Else
		ShowPNRecord = ShowPNRecord & "<a href=""?Action=Show&ID="& Rs("ID") &""">" & Rs("Title") &"</a>"
	End If
	Rs.Close
	ShowPNRecord = ShowPNRecord & "</li>"
End Function

Function CheckChar(CharContent,CharName,CharSize,Method)
	If Instr(Method,"0") And Len(CharContent) < 1 Then
		ErrContent = ErrContent & CharName & "中没有任何字符!"
	End If
	If Instr(Method,"1") And Server.HTMLEncode(CharContent) <> CharContent Then
		ErrContent = ErrContent & CharName & "中含有非法字符(包括网页代码)!"
	End If
	If Instr(Method,"2") And Len(CharContent) > CharSize Then
		ErrContent = ErrContent & CharName &"不可超过" & CharSize & "个字符!"
	End If
	If Instr(Method,"3") And CharContent <> "" And IsNumeric(CharContent) = False Then
		ErrContent = ErrContent & CharName & "必须为数字!"
	End If
	If Instr(Method,"4") And IsValidEmail(CharContent) = False Then
		ErrContent = ErrContent & CharName & "有误,其中必须含有""@""字符!"
	End If
	If Instr(Method,"5") And IsValidTel(CharContent) = False Then
		ErrContent = ErrContent & CharName & "有误,必须是数字、横线、括号或加号组成!"
	End If
	CheckChar=CharContent
End Function

Sub CheckData(formname,inputs)
	Dim j
	Content = "<script language=""javascript"">"
	Content = Content & "function CheckData(){"
	inputs = Split(inputs,"|")
	For i=0 To Ubound(inputs)
		m_Temp = Split(inputs(i),",")
		If Instr(m_Temp(3),"0") Then
			Content = Content & "	if("& formname &".all["""& m_Temp(0) &"""].value == """"){"
			Content = Content & "		window.alert("""& m_Temp(1) &"不可以是空值!"");"
			Content = Content & "		document."& formname &".all["""& m_Temp(0) &"""].focus();"
			Content = Content & "		return false;"
			Content = Content & "	}"
		End If
		If Instr(m_Temp(3),"1") Then
			Content = Content & "	if(!("& formname &".all["""& m_Temp(0) &"""].value.length "& m_Temp(2) &")){"
			Content = Content & "		window.alert("""& m_Temp(1) &"必须"& m_Temp(2) &"个字符!"");"
			Content = Content & "		document."& formname &".all["""& m_Temp(0) &"""].focus();"
			Content = Content & "		return false;"
			Content = Content & "	}"
		End If
		If Instr(m_Temp(3),"2") Then
			Content = Content & "	if(!Number("& formname &".all["""& m_Temp(0) &"""].value)){"
			Content = Content & "		window.alert("""& m_Temp(1) &"必须是数字!"");"
			Content = Content & "		document."& formname &".all["""& m_Temp(0) &"""].focus();"
			Content = Content & "		return false;"
			Content = Content & "	}"
		End If
		If Instr(m_Temp(3),"3") Then
			Content = Content & "	if("& formname &".all["""& m_Temp(0) &"""].value.indexOf('@') == -1 || "& formname &".all["""& m_Temp(0) &"""].value.indexOf('.') == -1){"
			Content = Content & "		window.alert(""请输入正确的Email地址!"");"
			Content = Content & "		document."& formname &".all["""& m_Temp(0) &"""].focus();"
			Content = Content & "		return false;"
			Content = Content & "	}"
		End If
		If Instr(m_Temp(3),"4") Then
			Content = Content & "	if(!CheckTel("& formname &".all["""& m_Temp(0) &"""].value)){"
			Content = Content & "		window.alert("""& m_Temp(1) &"有误,必须是数字、横线、括号或加号组成!"");"
			Content = Content & "		document."& formname &".all["""& m_Temp(0) &"""].focus();"
			Content = Content & "		return false;"
			Content = Content & "	}"
		End If
	Next
	Content = Content & "}"
	Content = Content & "function CheckTel(tel){"
	Content = Content & "	var tst,cTel=0;"
	Content = Content & "	var cmp=""0123456789-+() "";"
	Content = Content & "	for (var i=0;i<tel.length;i++){ "
	Content = Content & "		tst=tel.substring(i,i+1);"
	Content = Content & "		if (cmp.indexOf(tst)<0){"
	Content = Content & "			cTel++;"
	Content = Content & "		}"
	Content = Content & "	}"
	Content = Content & "	if(cTel!=0)"
	Content = Content & "		return false;"
	Content = Content & "	else"
	Content = Content & "		return true;"
	Content = Content & "}"
	Content = Content & "</script>"
	Response.Write(Content)
End Sub

Sub SelectOption(str1,str2)
	Content=""
	m_Temp=Split(str1,"|")
	For i=0 To Ubound(m_Temp)
		Content = Content &"<option value="""& m_Temp(i) &""""
		If str2=m_Temp(i) Then
			Content = Content &" selected"
		End if
		Content = Content &">"& m_Temp(i) &"</option>"
	Next
	Response.Write(Content)
End Sub

Sub SelectRadio(str,str1,str2)
	Content=""
	m_Temp=Split(str1,"|")
	For i=0 To Ubound(m_Temp)
		Content = Content &"<input name="""& str &""" type=""radio"" id="""& str & i &""" value="""& i &""""
		If str2=i Then
			Content = Content &" checked"
		End if
		Content = Content &">&nbsp;<label for="""& str & i &""">"& m_Temp(i) &"</label>&nbsp;&nbsp;"
	Next
	Response.Write(Content)
End Sub

Sub SelectValue(str1,str2)
	Content=""
	m_Temp=Split(str1,"|")
	For i=0 To Ubound(m_Temp)
	If str2=i Then
		Content = m_Temp(i)
	End if
	Next
	Response.Write(Content)
End Sub

Sub SelectClass(CurrentID)
	Dim RsClass,SqlClass,strTemp,tmpDepth,i
	For i=0 To Ubound(arRshowLine)
		arRshowLine(i)=False
	Next
	SqlClass="Select * From [Class] Where Layout='"&Layout&"' order by RootID,OrderID"
	Set RsClass=Server.CreateObject("Adodb.RecordSet")
	RsClass.Open SqlClass,Conn,1,1
	If RsClass.Bof And RsClass.Bof Then
		Response.Write("<option value='0'>请先添加分类</option>")
	Else
		Do While Not RsClass.Eof
			tmpDepth=RsClass("Depth")
			If RsClass("NextID")>0 Then
				arRshowLine(tmpDepth)=True
			Else
				arRshowLine(tmpDepth)=False
			End If
			strTemp="<option value="""
				strTemp=strTemp & RsClass("ClassID")
			strTemp=strTemp & """"
			If CurrentID>0 And RsClass("ClassID")=CurrentID Then
				 strTemp=strTemp & " selected"
			End If
			strTemp=strTemp & ">"
			
			If tmpDepth>0 Then
				For i=1 To tmpDepth
					strTemp=strTemp & "&nbsp;"
					If i=tmpDepth Then
						If RsClass("NextID")>0 Then
							strTemp=strTemp & "├&nbsp;"
						Else
							strTemp=strTemp & "└&nbsp;"
						End If
					Else
						If arRshowLine(i)=True Then
							strTemp=strTemp & "│"
						Else
							strTemp=strTemp & "&nbsp;"
						End If
					End If
				Next
			End If
			strTemp=strTemp & RsClass("ClassName")
'			strTemp=strTemp & RsClass("ClassName") &"("& RsClass("ClassID") &")"
			If RsClass("Child")>0 Then
				strTemp=strTemp & "("&RsClass("Child")&")"
			End If
			strTemp=strTemp & "</option>"
			Response.Write(strTemp)
			RsClass.MoveNext
		loop
	End If
	RsClass.Close
	Set RsClass=Nothing
End Sub

Function ShowClassName(ClassID)
	If ClassID=0 Then
		ShowClassName="一级目录"
	Else
		Dim RsTemp
		Set RsTemp=Conn.Execute("Select ClassID,ClassName From [Class] Where ClassID="&ClassID)
		If RsTemp.Eof Then
			ShowClassName="无此分类"
		Else
			ShowClassName=RsTemp("ClassName")
		End If
		RsTemp.Close
		Set RsTemp=Nothing
	End If
End Function

Sub ShowClassPath(ClassID)
	Dim RsClass,RsShow
	Content = ""
	Set RsClass=Conn.Execute("Select ClassID,ClassName,ParentPath From [Class] Where ClassID="& ClassID)
	If Not RsClass.Eof Then
		Content = "<a href=""?ClassID="&RsClass("ClassID")&""">" & RsClass("ClassName") & "</a>"
		Set RsShow=Conn.Execute("Select ClassID,ClassName From [Class] Where ClassID In ("&RsClass("ParentPath")&") Order By ClassID Desc")
		Do While Not RsShow.Eof
			Content = "<a href=""?ClassID="& RsShow("ClassID") &""">" & RsShow("ClassName") & "</a>&nbsp;>>&nbsp;" & Content
			RsShow.MoveNext
		Loop
		RsShow.Close
		Set RsShow=Nothing
	Else
		Content = "无此分类"
	End If
	RsClass.Close
	Set RsClass=Nothing
	Response.Write(Content)
End Sub

Sub ShowUpload(Obj)
Content = ""
Content = Content & "<iframe ID=""UpFile"" src=""Upload.asp?Obj="&Obj&""" frameborder=""0"" scrolling=""no"" width=""98%"" HEIGHT=""19"" allowtransparency=""true""></iframe>"
Response.Write(Content)
End Sub

Sub ShowEditor(TextName,TextContent)
	Content = "<script language=""JavaScript"">"
	Content = Content & "function doCheck(){"
	Content = Content & "	if (eWebEditor.getHTML() == """") {"
	Content = Content & "		alert(""内容不能为空!"");"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (myform.Title.value == """") {"
	Content = Content & "		alert(""标题不能为空!"");"
	Content = Content & "		document.myform.Title.focus();"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (myform.ClassID.value == 0) {"
	Content = Content & "		alert(""不能指定为含有子分类的分类!"");"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (myform.Picture.value != """" && myform.Pic.value == """") {"
	Content = Content & "		alert(""请指定默认图片!"");"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	eWebEditor.remoteUpload(""doSubmit();"");"
	Content = Content & "	return false;"
	Content = Content & "}"
	Content = Content & "function doSubmit(){"
	Content = Content & "	document.myform.submit();"
	Content = Content & "}"
	Content = Content & "function doChange(objText, objDrop){"
	Content = Content & "	if (!objDrop) return;"
	Content = Content & "	var str = objText.value;"
	Content = Content & "	var arr = str.split(""|"");"
	Content = Content & "	var nIndex = objDrop.selectedIndex;"
	Content = Content & "	objDrop.length=0;"
	Content = Content & "	for (var i=0; i<arr.length; i++){"
	Content = Content & "		objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);"
	Content = Content & "	}"
	Content = Content & "	objDrop.selectedIndex = nIndex;"
	Content = Content & "}"
	Content = Content & "</script>"
'	Content = Content & "<textarea name="""&TextName&""" rows=""20"" style=""width:98%"">"&TextContent&"</textarea>"
	Content = Content & "<input type=""hidden"" name="""&TextName&""" value="""&TextContent&"""><iframe ID=""eWebEditor"" src="""&EditorPath&"eWebEditor.asp?id="&TextName&"&style=s_mini&savepathfilename=Picture"" frameborder=""0"" scrolling=""no"" width=""98%"" HEIGHT=""400""></iframe>"
	Response.Write(Content)
End Sub

Sub ShowEditor1(TextName,TextContent)
	Content = "<script language=""JavaScript"">"
	Content = Content & "function doCheck(){"
	Content = Content & "	if (eWebEditor.getHTML() == """") {"
	Content = Content & "		alert(""内容不能为空!"");"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (myform.Title.value == """") {"
	Content = Content & "		alert(""标题不能为空!"");"
	Content = Content & "		document.myform.Title.focus();"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (myform.ClassID.value == 0) {"
	Content = Content & "		alert(""不能指定为含有子分类的分类!"");"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (myform.Picture.value != """" && myform.Pic.value == """") {"
	Content = Content & "		alert(""请指定默认图片!"");"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	eWebEditor.remoteUpload(""doSubmit();"");"
	Content = Content & "	return false;"
	Content = Content & "}"
	Content = Content & "function doSubmit(){"
	Content = Content & "	document.myform.submit();"
	Content = Content & "}"
	Content = Content & "function doChange(objText, objDrop){"
	Content = Content & "	if (!objDrop) return;"
	Content = Content & "	var str = objText.value;"
	Content = Content & "	var arr = str.split(""|"");"
	Content = Content & "	var nIndex = objDrop.selectedIndex;"
	Content = Content & "	objDrop.length=0;"
	Content = Content & "	for (var i=0; i<arr.length; i++){"
	Content = Content & "		objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);"
	Content = Content & "	}"
	Content = Content & "	objDrop.selectedIndex = nIndex;"
	Content = Content & "}"
	Content = Content & "</script>"
	Content = Content & "<textarea name="""&TextName&""" rows=""20"" style=""width:98%"">"&TextContent&"</textarea>"
'	Content = Content & "<input type=""hidden"" name="""&TextName&""" value="""&TextContent&"""><iframe ID=""eWebEditor"" src="""&EditorPath&"eWebEditor.asp?id="&TextName&"&style=s_mini&savepathfilename=Picture"" frameborder=""0"" scrolling=""no"" width=""98%"" HEIGHT=""400""></iframe>"
	Response.Write(Content)
End Sub

Sub ShowPage(Colspan)
	Content=""
'	If Query_String<>"" Then
'		Query_String=Replace(Query_String,"Page="&Page,"")
'	End If
	Content = Content & "<script language=""javascript"">"
	Content = Content & "function CheckAll(form){"
	Content = Content & "	for (var i=0;i<form.elements.length;i++){"
	Content = Content & "		var e = form.elements[i];"
	Content = Content & "		if (e.Name != ""chkAll""&&e.disabled==false)"
	Content = Content & "		e.checked = form.chkAll.checked;"
	Content = Content & "	}"
	Content = Content & "}"
	Content = Content & "function SDelete(){"
	Content = Content & "	if(confirm(""是否删除所选定的内容，删除后不能恢复！确定要删除吗？""))"
	Content = Content & "		return true;"
	Content = Content & "	else"
	Content = Content & "		return false;"
	Content = Content & "}"
	Content = Content & "</script>"
	Content = Content & "<tr align=""center"" class=""t_row""><td colspan="""&Colspan&""">"
	Content = Content & "<input name=""chkAll"" id=""chkAll"" type=""checkbox"" onclick=""CheckAll(this.form)""><label for=""chkAll"">全选</label>&nbsp;&nbsp;&nbsp;&nbsp;"
	Content = Content & "<input type=""submit"" value=""删除选择"" onclick=""return SDelete()"">&nbsp;&nbsp;&nbsp;&nbsp;"
		Content = Content & "共有：<strong>"&Rs.RecordCount&"</strong>条&nbsp;每页显示：<strong>"&Rs.PageSize&"</strong>条&nbsp;&nbsp;&nbsp;"
	if Rs.PageCount=1 or Rs.PageCount=0 then
		Content = Content & "[首页]&nbsp;[上页]&nbsp;[次页]&nbsp;[尾页]"
	else
		if Cint(Page)=1 then
			Content = Content & "[首页]&nbsp;[上页]&nbsp;"
			Content = Content & "<a href=""?"&Query_String&"Page="&Page+1&""">[次页]</a>&nbsp;<a href=""?"&Query_String&"Page="&Rs.PageCount&""">[尾页]</a>"
		elseif Cint(Page)=Rs.PageCount then
			Content = Content & "<a href=""?"&Query_String&"Page=1"">[首页]</a>&nbsp;<a href=""?"&Query_String&"Page="&Page-1&""">[上页]</a>&nbsp;"
			Content = Content & "[次页]&nbsp;[尾页]"
		else
			Content = Content & "<a href=""?"&Query_String&"Page=1"">[首页]</a>&nbsp;<a href=""?"&Query_String&"Page="&Page-1&""">[上页]</a>&nbsp;"
			Content = Content & "<a href=""?"&Query_String&"Page="&Page+1&""">[次页]</a>&nbsp;<a href=""?"&Query_String&"Page="&Rs.PageCount&""">[尾页]</a>"
		end if
	end if
	Content = Content & "&nbsp;&nbsp;&nbsp;页次：<strong>"&Page&"</strong>/"&Rs.PageCount&"页&nbsp;转到："
	Content = Content & "<select name=""Page"" onchange=""javascript:window.location='?"&Query_String&"Page='+this.options[this.selectedIndex].value+'';"">"
	for i=1 to Rs.PageCount
		Content = Content & "<option value="""&i&""""
		if Cint(Page) = i then Content = Content & " selected"
		Content = Content & ">第"&i&"页</option>"
	next
	Content = Content & "</select>"
	Content = Content & "</td></tr>"
	Response.Write(Content)
End Sub


Sub DelFiles(strUploadFiles)
	If strUploadFiles<>"" Then
		Dim fso,arrUploadFiles,i,FileName
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Instr(strUploadFiles,"|")>0 Then
			arrUploadFiles=Split(strUploadFiles,"|")
			For i=0 To Ubound(arrUploadFiles)
				If arrUploadfiles(i)<>"" Then
					FileName=Split(arrUploadfiles(i),"/")
					If fso.FileExists(server.MapPath(UpFilePath & FileName(Ubound(FileName)))) Then
						fso.DeleteFile(server.MapPath(UpFilePath & FileName(Ubound(FileName))))
					End If
				End If
			Next
		Else
			FileName=Split(strUploadfiles,"/")
			If fso.FileExists(server.MapPath(UpFilePath & FileName(Ubound(FileName)))) Then
				fso.DeleteFile(server.MapPath(UpFilePath & FileName(Ubound(FileName))))
			End If
		End If
		Set fso = Nothing
	End If
End Sub

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

Function IsValidTel(word)
	Dim names, name, i, c
	IsValidTel = True
   For i = 1 To Len(word)
	 c = Lcase(Mid(word, i, 1))
	 If InStr("1234567890-+() ", c) <= 0 And Not IsNumeric(c) Then
	   IsValidTel = False
	   Exit Function
	 End If
   Next
End Function

Function IsValidEmail(email)
	Dim names, name, i, c
	IsValidEmail = True
	names = Split(email, "@")
	If UBound(names) <> 1 Then
	   IsValidEmail = False
	   Exit Function
	End If
	For Each name In names
	   If Len(name) <= 0 Then
		 IsValidEmail = False
		 Exit Function
	   End If
	   For i = 1 To Len(name)
		 c = Lcase(Mid(name, i, 1))
		 If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
		   IsValidEmail = False
		   Exit Function
		 End If
	   Next
	   If Left(name, 1) = "." Or Right(name, 1) = "." Then
		  IsValidEmail = False
		  Exit Function
	   End If
	Next
	If InStr(names(1), ".") <= 0 Then
	   IsValidEmail = False
	   Exit Function
	End If
	i = Len(names(1)) - InStrRev(names(1), ".")
	If i <> 2 And i <> 3 Then
	   IsValidEmail = False
	   Exit Function
	End If
	If InStr(email, "..") > 0 Then
	   IsValidEmail = False
	End If
End Function

Function ShowClassTitle(ShowClassID,Method)
	Dim RsClass,RsShow,ParentPath
	Set RsClass=Conn.Execute("Select ClassID,ClassName,ParentPath,ParentID From [Class] Where ClassID="& ShowClassID)
	If Not RsClass.Eof Then
		If Method=0 Or ClassID=0 Then
			ParentPath = RsClass("ParentPath") &","& RsClass("ClassID")
		Else
			If ClassID=RsClass("ClassID") Then
				ParentPath = RsClass("ClassID")
			Else
				ParentPath = Replace(RsClass("ParentPath") &","& RsClass("ClassID"),","&ClassID&",","*")
'				Response.Write(RsClass("ParentPath") &","& RsClass("ClassID") &" > "& ParentPath & " = ")
				ParentPath = Replace(ClassID & Mid(ParentPath,Instr(ParentPath,"*")),"*",",")
			End If
		End If
'		Response.Write(ParentPath & " | ")
		Set RsShow=Conn.Execute("Select ClassID,ClassName From [Class] Where ClassID In ("&ParentPath&") Order By ClassID Desc")
		Do While Not RsShow.Eof
			If ShowClassTitle="" Then
'				ShowClassTitle = "<a href=""?ClassID="& RsShow("ClassID") &""">" & RsShow("ClassName") & "</a>"
				ShowClassTitle = RsShow("ClassName")
			Else
'				ShowClassTitle = "<a href=""?ClassID="& RsShow("ClassID") &""">" & RsShow("ClassName") & "</a>-" & ShowClassTitle
				ShowClassTitle = RsShow("ClassName") & " > " & ShowClassTitle
			End If
			RsShow.MoveNext
		Loop
		RsShow.Close
		Set RsShow=Nothing
	Else
		ShowClassTitle = "所有产品"
	End If
	RsClass.Close
	Set RsClass=Nothing
End Function

%>
