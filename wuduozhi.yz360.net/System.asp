<!--#include file="inc/Conn.asp"-->
<%
Dim Content,Page,i,Query_String,Http_Referer,Word,Action,ClassID,ID,IP,ShowUrl,arrShowLine(10)
Dim m_Config,m_WebSet,m_MSvr,m_Word,m_Count,m_OnLine,m_Temp,m_Temp2,TempSession
ShowUrl="Show.asp"
Http_Referer = Request.ServerVariables("HTTP_REFERER")
IP = Request.ServerVariables("REMOTE_ADDR")
Action = Trim(Request("Action"))
Word = Trim(Request("Word"))
ID = Request("ID")
If IsNumeric(ID) Then
	ID = Cint(ID)
Else
	Response.Write("<script>alert(""��¼��ű���������"");history.back();</script>")
	Response.End
End If
ClassID = Request("ClassID")
If IsNumeric(ClassID) Then
	ClassID = Cint(ClassID)
Else
	ClassID = 0
End If
Page = Request.QueryString("Page")
If IsNumeric(Page) Then
	Page = Cint(Page)
	If Page<1 Then Page=1
Else
	Page=1
End If

ConnectionDatabase()
Sql="Select Config,WebSet,MailSvr,Word,Counts,OnLine From Config"
Rs.Open Sql,Conn,1,2
If Not Rs.Eof Then
	m_OnLine = Split(Rs("OnLine"),"|")
	If IP<>m_OnLine(0) Or DateDiff("s",m_OnLine(1),Now())>1200 Then
		Rs("Counts")=Rs("Counts")+1
		Rs("OnLine")=IP &"|"& Now()
	End If
	Rs.Update
	m_Config=Split(Rs("Config"),"|")
	m_WebSet=Split(Rs("WebSet"),"|")
	m_MSvr=Split(Rs("MailSvr"),"|")
	m_Word=Rs("Word")
	m_Count=Rs("Counts")
	If m_Config(2)="" Then
		m_Config(2)=m_Config(0)
	End If
	If m_WebSet(1)="0" Then
		Response.Write(m_WebSet(2))
		Response.End
	End If
Else
	Response.Redirect "Admin/"
End If
Rs.Close

Select Case Action
	Case "SendGuest"
		Call SendGuest()
End Select

Sub Intro()
	Sql="Select Intro From Config"
	Rs.Open Sql,Conn,1,1
		Content = ""
'		Content = Content & "<table width=""98%"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""5"">"
'		Content = Content & "<tr><td height=""5""></td></tr><tr><td style=""padding:5px;line-height:200%"">"
		Content = Content & Rs("Intro")
'		Content = Content & "</td></tr><tr><td height=""5""></td></tr>"
'		Content = Content & "</table>"
	Rs.Close
	Response.Write(Content)
End Sub

Sub News(ClassID)
	If Word<>"" Then
		Query_String = "Word="& Word &"&"
	ElseIf ClassID>0 Then
		Query_String = "ClassID="& ClassID &"&"
	End If
	Content = ""
	Content = Content & "<table width=""90%"" align=""center"" border=""0"" cellpadding=""5"" cellspacing=""0"" style=""border-collapse:collapse"">"
'	Content = Content & "<tr bgcolor=""#CCCCCC""><td style=""border-bottom:1px #666666 solid""><strong>"&m_Temp&"</strong></td><td width=""120"" align=""center"" style=""border-bottom:1px #666666 solid""><strong>ʱ��</strong></td></tr>"
	Sql="Select * From News"
	If Word<>"" Then
		Sql = Sql & " Where Title Like '%"& Word &"%'"
	ElseIf Cint(ClassID)>0 Then
		Dim RsTemp,ChildID
		Set RsTemp=Conn.Execute("Select ParentPath From [Class] Where ClassID="&ClassID)
		If Not RsTemp.Eof Then
			Dim ParentPath
			ParentPath = RsTemp("ParentPath")
		Else
			Content = Content & "<tr><td align=""center"">��ѡ��Ҫ��ѯ�ķ���</td></tr>"
		End If
		Set RsTemp=Conn.Execute("Select ClassID From [Class] Where ParentID=" & ClassID & " Or ParentPath Like '"& ParentPath &","& ClassID &"%'")
		If Not RsTemp.Eof Then
			Do While Not RsTemp.Eof
				If ChildID="" Then
					ChildID=RsTemp(0)
				Else
					ChildID=ChildID & "," & RsTemp(0)
				End If
				RsTemp.MoveNext
			Loop
			Sql = Sql & " Where ClassID In ("&ClassID &","& ChildID &")"
		Else
			Sql = Sql & " Where ClassID="& ClassID
		End If
		RsTemp.Close
		Set RsTemp=Nothing
	End If
	Sql = Sql & " Order By Topis Desc,DateAndTime Desc"
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Content = Content & "<tr><td align=""center"">û���κμ�¼</td></tr>"
	Else
		i=0
		Rs.PageSize = Cint(m_WebSet(4))
		Rs.AbsolutePage = Page
		Do While Not Rs.Eof And i < Rs.PageSize
			Content = Content & "<tr><td><img src=""Images/"
			If Rs("Topis")=1 Then
				Content = Content & "m_topis"
			Else
				Content = Content & "m_redis"
			End If
			Content = Content & ".gif"" width=""23"" height=""7"" align=""absmiddle"">"
			Content = Content & "<a href="""&ShowUrl&"?Action=News&ID="&Rs("ID")&""" style="""
			If Rs("TitB")=1 Then
				Content = Content & "font-weight:bold;"
			End If
			If Rs("Red")=1 Then
				Content = Content & "color:#FF0000;"
			End If
			Content = Content & """>"& Rs("Title") & "</a>"
			Content = Content & "&nbsp;(<font color=""#CC3300"">"& rs("hits") &"</font>)"
			If DateDiff("d",Rs("DateAndTime"),Now())<7 Then
				Content = Content & "<img src=""Images/m_newis.gif"">"
			End If
			If Rs("Hits")>=Cint(m_WebSet(3)) Then
				Content = Content & "<img src=""Images/m_hotis.gif"">"
			End If
			Content = Content & "</td><td align=""center"">"
			Content = Content & "<font color=""#666666"">"& FormatDateTime(Rs("DateAndTime"),2) &"</font></td></tr>"
			i=i+1
			Rs.Movenext
		Loop
'		If Rs.RecordCount>20 Then
		Content = Content & "<tr><td colspan=""2"" align=""center"">"
		ShowPage(1)
		Content = Content & "</td></tr>"
'		End If
	End If
	Rs.Close
	Content = Content & "</table>"
	Response.Write(Content)
End Sub


Sub Product()
	If Word<>"" Then
		Query_String = "Word="& Word &"&"
	ElseIf ClassID>0 Then
		Query_String = "ClassID="& ClassID &"&"
	End If
	Content = ""
	Content = Content & "<script language=""javascript"">function Shop(id,title){var l,t;l=(screen.width-680)/2;t=(screen.Height-380)/2;ShopWin=window.open(""System.asp?Action=Shop&ID=""+id+""&Title=""+title,""Shop"",""width=680,height=380,left=""+l+"",top=""+t+"""");ShopWin.focus();}</script>" & vbCrlf
	Content = Content & "<table width=""98%"" align=""center"" border=""0"" cellpadding=""3"" cellspacing=""0"">"
	Content = Content & "<tr><td height=""5""></td></tr>"
	Sql="Select * From Product P,[Class] C Where"
	If Word<>"" Then
		Sql = Sql & " Where Title Like '%"& Word &"%'"
	ElseIf Cint(ClassID)>0 Then
		Dim RsTemp,ChildID
		Set RsTemp=Conn.Execute("Select ParentPath From [Class] Where ClassID="&ClassID)
		If Not RsTemp.Eof Then
			Dim ParentPath
			ParentPath = RsTemp("ParentPath")
		Else
			Content = Content & "<tr><td align=""center"">��ѡ��Ҫ��ѯ�ķ���</td></tr>"
		End If
		Set RsTemp=Conn.Execute("Select ClassID From [Class] Where ParentID=" & ClassID & " Or ParentPath Like '"& ParentPath &","& ClassID &"%'")
		If Not RsTemp.Eof Then
			Do While Not RsTemp.Eof
				If ChildID="" Then
					ChildID=RsTemp(0)
				Else
					ChildID=ChildID & "," & RsTemp(0)
				End If
				RsTemp.MoveNext
			Loop
			Sql = Sql & " P.ClassID In ("&ClassID &","& ChildID &") And"
		Else
			Sql = Sql & " P.ClassID="& ClassID &" And"
		End If
		RsTemp.Close
		Set RsTemp=Nothing
	End If
	Sql = Sql & " C.ClassID=P.ClassID Order By P.Topis Desc,P.DateAndTime Desc"
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Content = Content & "<tr><td align=""center"">û���κμ�¼</td><tr>"
	Else
		i=0
		Rs.PageSize = 5
		Rs.AbsolutePage = Page
		Do While Not Rs.Eof And i < Rs.PageSize
			Content = Content & "<tr>"
			Content = Content & "<td width=""30%"" align=""center""><table border=""0"" cellpadding=""3"" cellspacing=""2"" style=""border:4px #F2F2F2 solid""><tr><td align=""center"" style=""border:1px #CCCCCC solid""><a href="""&ShowUrl&"?Action=Product&ID="& Rs("ID") &""">"
			If Rs("Pic")="" Then
				Content = Content & "<img src=""Images/m_nopic.gif"" border=""0"">"
			Else
				Content = Content & "<img src="""& Rs("Pic") &""" width=""150"" height=""150"" border=""0"">"
			End If
			Content = Content & "</a></td></tr></table></td><td><table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""3"">"
			Content = Content & "<tr><td width=""15%""><strong>����:</strong></td><td><a href="""&ShowUrl&"?Action=Product&ID="& Rs("ID") &""">"& Rs("Title") &"</a></td></tr>"
			Content = Content & "<tr><td><strong>����:</strong></td><td>"& Rs("ClassName") &"</td></tr>"
			Content = Content & "<tr><td valign=""top""><strong>����:</strong></td><td style=""line-height:180%"">"& Left(delHtml(Rs("Content")),160)&"...<div align=""right""><a href="""&ShowUrl&"?Action=Product&ID="& Rs("ID") &""" style=""color:#666666""><img src=""Images/more_01.gif"" width=""36"" height=""7"" border=""0""></a></div></td></tr>"
'			Content = Content & "<tr><td>���߶�����</td><td><a href=""javascript:Shop("& Rs("ID") &",'"& Rs("Title") &"')"">����</a></td></tr>"
			Content = Content & "</table></td></tr><tr><td colspan=""3""><hr size=""1""></td></tr>"
			i=i+1
			Rs.Movenext
		Loop
	End If
	Content = Content & "<tr><td height=""5""></td></tr>"
	Content = Content & "</table><center>"
	ShowPage(3)
	Response.Write(Content&"<br><br>")
	Rs.Close
End Sub

Sub ShowContent()
	Dim Title,Author,Url,Price,Pic,Picture,Pictures,Topis,Hits,Red,TitB,TempContent,DateAndTime
	If Action="" Or Not(Action="News" Or Action="Product" Or Action="Picture" Or Action="Faq") Then
		Response.Write("<script>alert(""��ָ��Ҫ�鿴������"");history.back();</script>")
		Response.End
	End If
	Sql="Select * From "& Action &" Where ID="& ID
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Response.Write("<script>alert(""��¼������"");history.back();</script>")
		Response.End
	Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		Title=Rs("Title")
		If Action="News" Then
			Author=Rs("Author")
			Url=Rs("Url")
			Red=Rs("Red")
			TitB=Rs("TitB")
		ElseIf Action="Product" Then
			Price=Rs("Price")
		End If
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		TempContent=Rs("Content")
		DateAndTime=Rs("DateAndTime")
	End If
	Rs.Close
	Content = "<table width=""92%"" align=""center"" border=""0"" cellpadding=""3"" cellspacing=""0"">"
	Content = Content & "<tr><td class=""content_t"""
	If Red=1 Then
		Content = Content & " style=""color:#FF0000;"""
	End If
	Content = Content & ">"& Title &"</td></tr>"
	Content = Content & "<tr><td align=""center"">��&nbsp;&nbsp;�������ࣺ"& ShowClassName(ClassID) &"&nbsp;&nbsp;���������"& Hits &"��&nbsp;&nbsp;����ʱ�䣺"& DateAndTime &"&nbsp;&nbsp;��<hr weight=""90%"" size=""1"">"
	If Action="Product" Then
		Content = Content & "<div align=""center""><img src="""&Pic&"""></div><br>"
	End IF
	Content = Content & "</td></tr><tr><td class=""content_c"">"& Replace(Replace(TempContent,Chr(10),"<br>"),Chr(13),"&nbsp;")
	If Len(Picture)>1 And Action="Product" Then
		Content = Content & "<div><br><img src=""Images/arr.gif"" align=""absmiddle"">&nbsp;<a href="""&Picture&""" style=""font-weight:bold; color:#FF6633;"">˵��������</a></div><br>"
	End IF
'		Content = Content & "<br /><br />���빺����<br />"
	Content = Content & "</td></tr>"
'	Content = Content & "<hr><span class=""content_c"">"& ShowPNRecord(ID,ClassID,Action) &"</span><br>"
	Content = Content & "<tr><td align=""right"">��<a href=""javascript:window.print()"">��ӡ</a>��&nbsp;��<a href=""javascript:history.back()"">����</a>��&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table>"
	Response.Write(Content)
End Sub
''**************************************************
''��������ShowPNRecord
''��  �ã��õ���һҳ,��һҳ��¼
''**************************************************
'Function ShowPNRecord(ID,ClassID,Action,Url)
'	ShowPNRecord = ""
'	Sql="ID < "& ID &" Order By ID Desc"
'	ShowPNRecord = ShowPNRecord & "<b>��һƪ</b>��"
'	Sql="Select ID,Title,ClassID From "& Action &" Where ClassID="& ClassID &" And "& Sql
'	Rs.Open Sql,Conn,1,2
'	If Rs.Eof Then
'		ShowPNRecord = ShowPNRecord & "û�м�¼"
'	Else
'		ShowPNRecord = ShowPNRecord & "<a href="""& Url &"="& Rs("ID") &""">" & Rs("Title") &"</a>"
'	End If
'	Rs.Close
'	ShowPNRecord = ShowPNRecord & "&nbsp;|&nbsp;"
'
'	Sql="ID > "& ID &" Order By ID Asc"
'	ShowPNRecord = ShowPNRecord & "<b>��һƪ</b>��"
'	Sql="Select ID,Title,ClassID From "& Action &" Where ClassID="& ClassID &" And "& Sql
'	Rs.Open Sql,Conn,1,2
'	If Rs.Eof Then
'		ShowPNRecord = ShowPNRecord & "û�м�¼"
'	Else
'		ShowPNRecord = ShowPNRecord & "<a href="""& Url &"="& Rs("ID") &""">" & Rs("Title") &"</a>"
'	End If
'	Rs.Close
'	ShowPNRecord = ShowPNRecord & ""
'End Function

Sub Guest()
	Content = ""
	Sql="Select * From Guest Order By DateAndTime Desc"
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Content = Content & "<div align=""center"">û���κμ�¼</div>"
	Else
		i=1
		Rs.PageSize = Cint(m_WebSet(6))
		Rs.AbsolutePage = Page
		Do While Not Rs.Eof And i < Rs.PageSize+1
			Content = ""
			Content = Content & "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""guest_t"">"
			Content = Content & "<tr><td class=""guest_s"">&nbsp;<strong>"& Rs("Title") &"</strong>&nbsp;("& Rs("Sender") &")</td></tr>"
			Content = Content & "<tr><td class=""guest_c""><table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""0"">"
			Content = Content & "<tr><td width=""5""><img src=""Images/space.gif"" width=""5"" height=""5""></td><td colspan=""2"">��������:<br><textarea rows=""3"" style=""width:100%"" readonly=""true"">"& Rs("Content") &"</textarea></td><td width=""5""><img src=""Images/space.gif"" width=""5"" height=""5""></td></tr>"
			If Rs("Revert")<>"" And Not IsNull(Rs("Revert")) Then
				Content = Content & "<tr><td width=""5""><img src=""Images/space.gif"" width=""5"" height=""5""></td><td colspan=""2"">�ظ�����:<br><textarea rows=""2"" style=""width:100%"" readonly=""true"">"& Rs("Revert") &"</textarea></td><td width=""5""><img src=""Images/space.gif"" width=""5"" height=""5""></td></tr>"
			End If
			Content = Content & "<tr><td width=""5""><img src=""Images/space.gif"" width=""5"" height=""5""></td><td>����ʱ��:"& Rs("DateAndTime") &"</td><td>IP:"& Rs("IP") &"</td><td width=""5""><img src=""Images/space.gif"" width=""5"" height=""5""></td></tr>"
			Content = Content & "</table></td></tr></table>"
			Response.Write(Content)
			i=i+1
			Rs.Movenext
		Loop
		Content = "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		ShowPage(1)
		Content = Content & "</table>"
	Response.Write(Content)
	End If
	Rs.Close
End Sub

Sub GuestAdd()
	Randomize timer
	Session("m_Code") = Int(Rnd*8998)+1000
	Content = ""
	Content = Content & "<script language=""javascript"">"
	Content = Content & "function add_guest()"
	Content = Content & "{"
	Content = Content & "	if (document.Add_G.Title.value=="""")"
	Content = Content & "	{"
	Content = Content & "		alert(""��������������!"");"
	Content = Content & "		document.Add_G.Title.focus();"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (document.Add_G.Sender.value=="""")"
	Content = Content & "	{"
	Content = Content & "		alert(""�������ύ������!"");"
	Content = Content & "		document.Add_G.Sender.focus();"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (document.Add_G.Email.value=="""")"
	Content = Content & "	{"
	Content = Content & "		alert(""����������ʼ���ַ!"");"
	Content = Content & "		document.Add_G.Email.focus();"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "	if (document.Add_G.Content.value=="""")"
	Content = Content & "	{"
	Content = Content & "		alert(""��������������!"");"
	Content = Content & "		document.Add_G.Content.focus();"
	Content = Content & "		return false;"
	Content = Content & "	}"
	Content = Content & "}"
	Content = Content & "</script>" 
	Content = Content & "<table width=""90%"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""5"">"
	Content = Content & "<form name=""Add_G"" method=""post"" action=""?Action=SendGuest"" onSubmit=""return add_guest();"">"
	Content = Content & "<tr><td height=""5"" colspan=""2""></td></tr>"
	Content = Content & "<tr><td width=""20%"" align=""right"">����:</td><td>&nbsp;<input name=""Title"" type=""text"" id=""Title"" size=""38""></td></tr>"
	Content = Content & "<tr><td align=""right"">�ύ��:</td><td>&nbsp;<input name=""Sender"" type=""text"" id=""Sender"" size=""38""></td></tr>"
	Content = Content & "<tr><td align=""right"">Email:</td><td>&nbsp;<input name=""Email"" type=""text"" id=""Email"" size=""38""></td></tr>"
	Content = Content & "<tr><td align=""right"">��������:</td><td>&nbsp;<textarea name=""Content"" rows=""5"" id=""Content"" style=""width:350px""></textarea></td></tr>"
	Content = Content & "<tr><td align=""right"">��֤��:</td><td>&nbsp;<input name=""Code"" type=""text"" id=""Code"" size=""4"">&nbsp;" & Session("m_Code") & "</td></tr>"
	Content = Content & "<tr><td>&nbsp;</td><td><input type=""submit"" value=""�ύ""></td></tr>"
	Content = Content & "<tr><td height=""5"" colspan=""2""></td></tr>"
	Content = Content & "</form>"
	Content = Content & "</table>"
	Response.Write(Content)
End Sub

Function ShowClassName(ClassID)
	If ClassID=0 Then
		ShowClassName = "��������"
	Else
		Dim RsTemp
		Set RsTemp=Conn.Execute("Select ClassID,ClassName From [Class] Where ClassID="&ClassID)
		If RsTemp.Eof Then
			ShowClassName = "�޴˷���"
		Else
			ShowClassName = RsTemp("ClassName")
		End If
		RsTemp.Close
		Set RsTemp=Nothing
	End If
End Function
'**************************************************
'��������ShowClassTitle
'��  �ã��õ���������
'**************************************************
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
				ShowClassTitle = RsShow("ClassName") & "-" & ShowClassTitle
			End If
			RsShow.MoveNext
		Loop
		RsShow.Close
		Set RsShow=Nothing
	Else
		ShowClassTitle = "��������"
	End If
	RsClass.Close
	Set RsClass=Nothing
End Function
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
	Content = Content & "	if(confirm(""�Ƿ�ɾ����ѡ�������ݣ�ɾ�����ָܻ���ȷ��Ҫɾ����""))"
	Content = Content & "		return true;"
	Content = Content & "	else"
	Content = Content & "		return false;"
	Content = Content & "}"
	Content = Content & "</script>"
	Content = Content & "<tr align=""center"" class=""t_row""><td colspan="""&Colspan&""">"
	
		Content = Content & "���У�<strong>"&Rs.RecordCount&"</strong>��&nbsp;ÿҳ��ʾ��<strong>"&Rs.PageSize&"</strong>��&nbsp;&nbsp;&nbsp;"
	if Rs.PageCount=1 or Rs.PageCount=0 then
		Content = Content & "[��ҳ]&nbsp;[��ҳ]&nbsp;[��ҳ]&nbsp;[βҳ]"
	else
		if Cint(Page)=1 then
			Content = Content & "[��ҳ]&nbsp;[��ҳ]&nbsp;"
			Content = Content & "<a href=""?"&Query_String&"Page="&Page+1&""">[��ҳ]</a>&nbsp;<a href=""?"&Query_String&"Page="&Rs.PageCount&""">[βҳ]</a>"
		elseif Cint(Page)=Rs.PageCount then
			Content = Content & "<a href=""?"&Query_String&"Page=1"">[��ҳ]</a>&nbsp;<a href=""?"&Query_String&"Page="&Page-1&""">[��ҳ]</a>&nbsp;"
			Content = Content & "[��ҳ]&nbsp;[βҳ]"
		else
			Content = Content & "<a href=""?"&Query_String&"Page=1"">[��ҳ]</a>&nbsp;<a href=""?"&Query_String&"Page="&Page-1&""">[��ҳ]</a>&nbsp;"
			Content = Content & "<a href=""?"&Query_String&"Page="&Page+1&""">[��ҳ]</a>&nbsp;<a href=""?"&Query_String&"Page="&Rs.PageCount&""">[βҳ]</a>"
		end if
	end if
	Content = Content & "&nbsp;&nbsp;&nbsp;ҳ�Σ�<strong>"&Page&"</strong>/"&Rs.PageCount&"ҳ&nbsp;ת����"
	Content = Content & "<select name=""Page"" onchange=""javascript:window.location='?"&Query_String&"Page='+this.options[this.selectedIndex].value+'';"">"
	for i=1 to Rs.PageCount
		Content = Content & "<option value="""&i&""""
		if Cint(Page) = i then Content = Content & " selected"
		Content = Content & ">��"&i&"ҳ</option>"
	next
	Content = Content & "</select>"
	Content = Content & "</td></tr>"
	Response.Write(Content)
End Sub


'<!--Sub SendGuest()
'	Dim Title,Sender,Email,Content,Code
'	Title=Trim(Request.Form("Title"))
'	Sender=Trim(Request.Form("Sender"))
'	Email=Trim(Request.Form("Email"))
'	Content=Trim(Request.Form("Content"))
'	Code=Trim(Request.Form("Code"))
'	If IsNumeric(Code)=False Then Response.Write("<script>alert(""��֤����Ϊ����"");history.back()< /script>")
'	If Sender="" Or Email="" Or Content="" Or Int(Code) <> Session("m_Code") Then
'		Response.Write("< script>alert(""��������������Ϣ."");history.back()< /script>")
'	Else
'		Sql="Select * From [Guest]"
'		Rs.Open Sql,Conn,1,2
'			Rs.AddNew
'			Rs("Title")=Title
'			Rs("Sender")=Sender
'			Rs("Email")=Email
'			Rs("Content")=Server.HTMLEncode(Content)
'			Rs("Ip")=IP
'			Rs("DateAndTime")=Now()
'			Rs.Update
'		Rs.Close
'		Response.Write("<script>alert(""�������ύ����ȴ��ظ�"");location.href="""& Http_Referer &"""< /script>")
'	End If
'	Response.End
'End Sub-->
'**************************************************
'��������SelectClass
'��  �ã��õ������»���
'**************************************************
Sub SelectClass(ClassID,CurrentID,Layout)
	Dim RsClass,SqlClass,strTemp,tmpDepth,arRshowLine(20)
	For i=0 To Ubound(arRshowLine)
		arRshowLine(i)=False
	Next
	SqlClass="Select * From [Class] Where"& ClassSQL(ClassID) &" And Layout='"&Layout&"' order by RootID,OrderID"
	Set RsClass=Server.CreateObject("Adodb.RecordSet")
	RsClass.Open SqlClass,Conn,1,1
	If RsClass.Bof And RsClass.Bof Then
		Response.Write("<option value='0'>������ӷ���</option>")
	Else
		Do While Not RsClass.Eof
			tmpDepth=RsClass("Depth")
			If RsClass("NextID")>0 Then
				arRshowLine(tmpDepth)=True
			Else
				arRshowLine(tmpDepth)=False
			End If
			strTemp="<option value="""
			If RsClass("Child")>0 Then
				strTemp=strTemp & 0
			Else
				strTemp=strTemp & RsClass("ClassID")
			End If
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
							strTemp=strTemp & "��&nbsp;"
						Else
							strTemp=strTemp & "��&nbsp;"
						End If
					Else
						If arRshowLine(i)=True Then
							strTemp=strTemp & "��"
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
'**************************************************
'��������ClassSQL
'��  �ã���ѯ����
'**************************************************
Function ClassSQL(ClassID)
	If ClassID=0 Then
		ClassSQL=" 1=1"
		Exit Function
	End If
	Dim RsTemp,ChildID
	Set RsTemp=Conn.Execute("Select ParentPath From [Class] Where ClassID="&ClassID)
	If Not RsTemp.Eof Then
		Dim ParentPath
		ParentPath = RsTemp("ParentPath")
	Else
		Content = Content & "<tr><td align=""center"">��ѡ��Ҫ��ѯ�ķ���</td></tr>"
	End If
	Set RsTemp=Conn.Execute("Select ClassID From [Class] Where ParentID=" & ClassID & " Or ParentPath Like '"& ParentPath &","& ClassID &"%'")
	If Not RsTemp.Eof Then
		Do While Not RsTemp.Eof
			If ChildID="" Then
				ChildID=RsTemp(0)
			Else
				ChildID=ChildID & "," & RsTemp(0)
			End If
			RsTemp.MoveNext
		Loop
		ClassSql = " ClassID In ("&ClassID &","& ChildID &")"
	Else
		ClassSql = " ClassID="& ClassID
	End If
	RsTemp.Close
	Set RsTemp=Nothing
End Function
'**************************************************
'��������delHtml
'��  �ã�ɾ��Html����
'**************************************************
Function delHtml(strHtml)
	Dim objRegExp, strOutput
	Set objRegExp = New Regexp
	
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "(<[a-zA-Z].*?>)|(<[\/][a-zA-Z].*?>)|(\r)"
	
	strOutput = objRegExp.Replace(strHtml,"")
	strOutput = Replace(strOutput, "<", "&lt;")
	strOutput = Replace(strOutput, ">", "&gt;") 
	strOutput = Replace(strOutput, "&nbsp;","") 
	delHtml = strOutput
	Set objRegExp = Nothing
End Function
'****************************************************
'��������LeftStr
'��  �ã���ȡ�ַ���
'��  ����Str ----�ַ��� Length��ȡ����
'����ֵ��Ҫ��ȡ���ַ���
'****************************************************
Function LeftStr(ByVal Str,ByVal Length)
	Dim X,Y,I,M,C
	Str=Trim(Str)
	If IsNull(Str) Then
		LeftStr=""
		Exit Function
	End If
	X=Len(Str)
	Y=0
	For I = 1 To X
		M=Mid(Str,I,1)
		C=Asc(M)
		If C<0 Then C=C+65536
		If C>255 Then
			Y=Y+2
		Else
			Y=Y+1
		End If
		If Y>=Length Then
			LeftStr=Left(Str,I)+".."
			Exit For
		Else
			LeftStr=Str
		End If
	Next
End Function
'**************************************************
'��������GetLocationURL
'��  �ã���ȡ��ǰҳ��ַ
'**************************************************
Function GetLocationURL() 
Dim Url 
Dim ServerPort,ServerName,ScriptName,QueryString 
ServerName = Request.ServerVariables("SERVER_NAME") 
ServerPort = Request.ServerVariables("SERVER_PORT") 
ScriptName = Request.ServerVariables("SCRIPT_NAME") 
QueryString = Request.ServerVariables("QUERY_STRING") 
Url="http://"&ServerName 
If ServerPort <> "80" Then Url = Url & ":" & ServerPort 
Url=Url&ScriptName 
If QueryString <>"" Then Url=Url&"?"& QueryString 
GetLocationURL=Url 
End Function
'**************************************************
'��������CheckStr
'��  �ã����˷Ƿ��ַ�
'**************************************************
Function CheckStr(ByVal Str)
	If Str="" Or IsNull(Str) Then
		CheckStr=""
		Exit Function
	End If
	'Str=LCase(Str)
	Str=Replace(Str,Chr(0),"")
	Str=Replace(Str,"""","&quot;")
	Str=Replace(Str,"'","''")
	Str=Replace(Str,"=","&#061")
	Str=Replace(Str,"<","&lt;")
	Str=Replace(Str,">","&gt;")
	Str=Replace(Str,"[","&#091;")
	Str=Replace(Str,"]","&#093;")
	Str=Replace(Str,"select","sel&#101;ct")
	Str=Replace(Str,"execute","&#101;xecute")
	Str=Replace(Str,"exec","&#101;xec")
	Str=Replace(Str,"join","jo&#105;n")
	Str=Replace(Str,"insert","ins&#101;rt")
	Str=Replace(Str,"delete","del&#101;te")
	Str=Replace(Str,"update","up&#100;ate")
	Str=Replace(Str,"drop","dro&#112;")
	Str=Replace(Str,"create","cr&#101;ate")
	Str=Replace(Str,"truncate","trunc&#097;te")
	Str=Replace(Str,"chr","c&#104;r")
	Str=Replace(Str,"count","co&#117;nt")
	Str=Replace(Str,"mid","m&#105;d")
	Str=Replace(Str,"char","ch&#097;r")
	Str=Replace(Str,"alter","alt&#101;r")
	Str=Replace(Str,"exists","e&#120;ists")
	Str=Replace(Str,"script","&#115;cript")
	Str=Replace(Str,"object","&#111;bject")
	Str=Replace(Str,"applet","&#097;pplet")
	CheckStr=Str
End Function
'================================================
'��������GetPrev
'�����ã���ȡ��һҳ
'�Ρ�������
'================================================
Function GetPrev(Id,Table,Pid,Url)
dim TempStr
Sql= "SELECT TOP 1 "&Id&",Title FROM "&Table&" WHERE "&Id&"<"&Pid&" ORDER BY "&Id&" DESC" 
Set Rs=Conn.Execute(Sql) 
If Rs.Bof And Rs.Eof Then 
TempStr=TempStr&"û����"&vbcrlf
Else
TempStr=TempStr&"<a href="&url&""&rs(0)&">"&rs(1)&"</a>"&vbcrlf
End If 
GetPrev =TempStr
End Function 
'================================================
'��������GetNext
'�����ã���ȡ��һҳ
'�Ρ�������
'================================================
Function GetNext(Id,Table,Pid,Url)
dim TempStr 
Sql= "SELECT TOP 1 "&Id&",Title FROM "&Table&" WHERE "&Id&">"&Pid&" ORDER BY "&Id&" DESC" 
Set Rs=Conn.Execute(Sql)
If Rs.Bof And Rs.Eof Then 
TempStr=TempStr&"û����"&vbcrlf
Else
TempStr=TempStr&"<a href="&url&""&rs(0)&">"&rs(1)&"</a>"&vbcrlf
End If 
GetNext =TempStr
End Function
'****************************************************
'��������FormatTime
'��  �ã���ʽ��ʱ��
'��  ����TimeStr-ʱ�� Types-��ȡ��ʽ
'����ֵ��ʱ��
'****************************************************
Function FormatTime(ByVal TimeStr,ByVal Types)
	Dim D_Year,D_Month,D_Day
	Dim D_Hour,D_Minute,D_Second
	If TimeStr="" Or IsNull(TimeStr) Then
		TimeStr=Now
	End If
	If Not (IsDate(TimeStr)) Then
		FormatTime=""
		Exit Function
	End If
	D_Year=Year(TimeStr)
	D_Month=Month(TimeStr)
	D_Day=Day(TimeStr)
	D_Hour=Hour(TimeStr)
	D_Minute=Minute(TimeStr)
	D_Second=Second(TimeStr)
	If Len(D_Month)<2 Then D_Month="0"&D_Month
	If Len(D_Day)<2 Then D_Day="0"&D_Day
	If Len(D_Hour)<2 Then D_Hour="0"&D_Hour
	If Len(D_Minute)<2 Then D_Minute="0"&D_Minute
	If Len(D_Second)<2 Then D_Second="0"&D_Second
	Select Case Types
		Case 1'2000-10-10 23:45:45
		FormatTime=D_Year&"-"&D_Month&"-"&D_Day&"-"&D_Hour&":"&D_Minute&":"&D_Second
		Case 2'��-��-�� ʱ:��:��
		FormatTime=D_Year&"��"&D_Month&"��"&D_Day&"��"&D_Hour&"ʱ"&D_Minute&"��"&D_Second&"��"
		Case 3'10-10
		FormatTime=D_Month&"-"&D_Day
		Case 4'2003-10-10
		FormatTime=D_Year&"-"&D_Month&"-"&D_Day
		Case 5'2003��10��10��
		FormatTime=D_Year&"��"&D_Month&"��"&D_Day&"��"
		Case 6'20031010102562
		FormatTime=D_Year&D_Month&D_Day&D_Hour&D_Minute&D_Second
		Case Else
	End Select
	FormatTime=FormatTime
End Function
%>
