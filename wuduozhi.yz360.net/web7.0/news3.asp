<!--#include file="System.asp"-->
<%
Head()
Layout="News"
%>
<script language="javascript" src="images/Admin_Js.Js"></script>

<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">信息管理</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2"><strong>操作选项</strong>：<a href="News.asp">信息管理</a>&nbsp;|&nbsp;<a href="News.asp?Action=Add">添加信息</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=News">分类管理</a></td>
  </tr>
  <tr class="t_rowh">
    <form method="get" action="News.asp">
      <td width="60%">快速搜索：&nbsp;
        <input name="Action" type="hidden" id="Action" value="Search">
        <input name="Word" type="text" id="Word" size="48"></td>
    </form>
    <td width="40%" align="right">查看分类：&nbsp;
      <select name="ClassID" id="ClassID" onchange="window.location='News.asp?ClassID='+this.options[this.selectedIndex].value+'';">
        <option value="0">跳转分类至</option>
        <%SelectClass(Cint(ClassID))%>
      </select>
      &nbsp;&nbsp;&nbsp;</td>
  </tr>
</table>
<%
Select Case Action
	Case "Add"
		Call NewsAdd()
	Case "Mod"
		Call NewsMod()
	Case "Show"
		Call NewsShow()
	Case "Del"
		Call NewsDel()
	Case "SDel"
		Call NewsSDel()
	Case "Topis"
		Call NewsTopis()
	Case "Hot"
		Call NewsHot()
	Case "Red"
		Call NewsRed()
	Case "TitB"
		Call NewsTitB()
	Case "SaveAdd"
		Call SaveAdd()
	Case "SaveMod"
		Call SaveMod()
	Case Else
		Call NewsList()
End Select

Sub NewsList()
	If Action="Search" Then
		Sql="Where Title Like '%"&Request("Word")&"%'"
		Query_String="Word="&Request("Word")&"&"
	ElseIf Cint(ClassID)>0 Then
		Dim RsTemp,ChildID
		Set RsTemp=Conn.Execute("Select ParentPath From [Class] Where ClassID="&ClassID)
		If Not RsTemp.Eof Then
			Dim ParentPath
			ParentPath = RsTemp("ParentPath")
		Else
			Msg "请选择要查询的分类","News.asp",1
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
			Sql="Where ClassID In ("&ClassID &","& ChildID &")"
		Else
			Sql="Where ClassID="& ClassID
		End If
		RsTemp.Close
		Set RsTemp=Nothing
		Query_String="ClassID="&ClassID&"&"
	Else
		Sql=""
		Query_String=""
	End If
	Rs.Open "Select * From [News] "&Sql&" Order By ID Desc",Conn,1,1
%>
<form name="myform" method="post" action="News.asp?Action=SDel">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr align="center" class="t_title">
      <td width="5%">选择</td>
      <td>标题</td>
      <td width="18%">分类</td>
      <td width="10%">点击次数</td>
      <td width="20%">操作</td>
    </tr>
    <%
	If Rs.Eof Then
		Response.Write("<tr align=""center"" class=""t_rowh""><td colspan=""5""><strong>没有任何记录</strong></td></tr>")
	Else
		i=0
		Rs.PageSize = Cint(Line)
		Rs.AbsolutePage = Cint(Page)
		Do While Not Rs.Eof And i < Rs.PageSize
%>
    <tr align="center" class="t_row" onmouseover="this.className='t_rowh'" onmouseout="this.className='t_row'">
      <td><input name="ID" type="checkbox" id="ID" value="<%=Rs("ID")%>"></td>
      <td align="left">&nbsp;<a href="?Action=Show&ID=<%=Rs("ID")%>"><%=Replace(Rs("Title"),Request("Word"),"<strong>"&Request("Word")&"</strong>")%></a></td>
      <td><%=ShowClassName(Rs("ClassID"))%></td>
      <td><%=(Rs("Hits"))%></td>
      <td><a href="News.asp?Action=Topis&ID=<%=Rs("ID")%>">
        <%If Rs("Topis")=1 Then%>
        <font color="#009933">推荐</font>
        <%Else%>
        推荐
        <%End If%>
        </a>&nbsp;&nbsp;<a href="News.asp?Action=Mod&ID=<%=Rs("ID")%>">修改</a>&nbsp;<a href="News.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">删除</a></td>
    </tr>
    <%
		i = i + 1
		Rs.MoveNext
		Loop
		ShowPage(5)
	End If
	Rs.Close
%>
  </table>
</form>
<Script language="JavaScript" type="text/JavaScript">
function Delete()
{
   if(confirm("删除内容后不能恢复！确定要删除吗？"))
     return true;
   else
     return false;
}
</Script>
<%
End Sub

Sub NewsAdd()
%>
<form name="myform" method="post" action="News.asp?Action=SaveAdd" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">添加信息</td>
    </tr>
    <tr class="t_rowh">
      <td width="20%"><u>所属分类</u></td>
      <td width="80%"><select name="ClassID" id="ClassID">
          <%SelectClass(0)%>
        </select>
        &nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
    </tr>
    <tr class="t_rowh">
      <td><u>信息标题</u></td>
      <td><input name="Title" type="text" id="Title" size="48">
        &nbsp;<font color="#FF0000">＊</font> ≤ 100个字符</td>
    </tr>
    <tr class="t_row">
      <td><u>自主SEO优化--信息描述：</u></td>
      <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="信息描述可与标题相同..." size="48">
      (也可不填)</td>
    </tr>
    <tr class="t_row">
      <td><u>自主SEO优化--搜索关键字：</u></td>
      <td><input name="Keywords" type="text" id="Keywords" value="请输入该信息搜索时的关键字..." size="48">
      (也可不填)</td>
    </tr>
    <tr class="t_row" style="display:none">
      <td><u>信息出处</u></td>
      <td><input name="Author" type="text" id="Author" size="40">
        &nbsp;&nbsp;当前为空值时，“出处链接”无效</td>
    </tr>
    <tr class="t_row" style="display:none">
      <td><u>出处链接</u></td>
      <td><input name="Url" type="text" id="Url" size="40">
        &nbsp;&nbsp;格式：“http://”</td>
    </tr>
    <tr class="t_rowh">
      <td><u>默认图片</u></td>
      <td><input name="Picture" type="hidden" id="Picture" size="48" onchange="doChange(this,Pic)" value="">
        <select name="Pic" id="Pic" size="1">
          <option value="">无</option>
        </select>
        &nbsp;&nbsp;当编辑区有插入图片时，可选择此下拉框</td>
    </tr>
    <tr class="t_rowh">
      <td><u>信息属性</u></td>
      <td><input name="Topis" type="checkbox" id="Topis" value="True">
        &nbsp;
        <label for="Topis">推荐信息</label>
        &nbsp;&nbsp;<label for="TitB"></label></td>
    </tr>
    <tr class="t_row">
      <td><u>信息内容</u>
        <li>本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度。</li>
        <li>换行请按Shift+Enter，另起一段请按Enter</li></td>
      <td><textarea name="Content" style="display:none"><%=Content%></textarea>
        <iframe id="myeditor" src="../Ok_Editor/ewebeditor.htm?id=Content&amp;style=coolblue&amp;maxlen=300&amp;savepathfilename=Picture" frameborder="0" scrolling="No" width="95%" height="405"></iframe></td>
    </tr>
    <tr class="t_row">
      <td>&nbsp;</td>
      <td><input type="submit" value="提交"></td>
    </tr>
  </table>
</form>
<%
End Sub

Sub NewsMod()
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Pictures,Topis,Hits,Red,TitB,Content,DescriptionWord,Keywords
	Sql="Select * From [News] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "信息不存在!","New.asp",1
	Else
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		DescriptionWord=Rs("DescriptionWord")
		Keywords=Rs("Keywords")
		Title=Rs("Title")
		Author=Rs("Author")
		Url=Rs("Url")
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		Red=Rs("Red")
		TitB=Rs("TitB")
		Content=Rs("Content")
	End If
	Rs.Close
%>
<form name="myform" method="post" action="News.asp?Action=SaveMod" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">修改信息</td>
    </tr>
    <tr class="t_rowh">
      <td width="20%"><u>所属分类</u></td>
      <td width="80%"><input name="ID" type="hidden" value="<%=ID%>">
        <select name="ClassID" id="ClassID">
          <%SelectClass(ClassID)%>
        </select>
        &nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
    </tr>
    <tr class="t_rowh">
      <td><u>信息标题</u></td>
      <td><input name="Title" type="text" id="Title" value="<%=Title%>" size="48">
        &nbsp;<font color="#FF0000">＊</font> ≤ 100个字符</td>
    </tr>
    <tr class="t_row">
      <td><u>自主SEO优化--信息描述：</u></td>
      <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="<%=DescriptionWord%>" size="48"></td>
    </tr>
    <tr class="t_row">
      <td><u>自主SEO优化--搜索关键字：</u></td>
      <td><input name="Keywords" type="text" id="Keywords" value="<%=Keywords%>" size="48"></td>
    </tr>
    <tr class="t_row" style="display:none">
      <td><u>信息出处</u></td>
      <td><input name="Author" type="text" id="Author" value="<%=Author%>" size="40">
        &nbsp;&nbsp;当前为空值时，“出处链接”无效</td>
    </tr>
    <tr class="t_row" style="display:none">
      <td><u>出处链接</u></td>
      <td><input name="Url" type="text" id="Url" value="<%=Url%>" size="40">
        &nbsp;&nbsp;格式：“http://”</td>
    </tr>
    <tr class="t_rowh">
      <td><u>默认图片</u></td>
      <td><input name="Picture" type="hidden" id="Picture" size="48" onchange="doChange(this,Pic)" value="<%=Picture%>">
        <select name="Pic" id="Pic" size="1">
          <option value="">无</option>
          <%SelectOption Picture,Pic%>
        </select>
        &nbsp;&nbsp;当编辑区有插入图片时，可选择此下拉框</td>
    </tr>
    <tr class="t_rowh">
      <td><u>信息属性</u></td>
      <td><input name="Topis" type="checkbox" id="Topis" value="True"<%If Topis=1 Then%> checked<%End If%>>
        &nbsp;
        <label for="Topis">推荐信息</label>
        &nbsp;&nbsp;<label for="TitB"></label></td>
    </tr>
    <tr class="t_row">
      <td><u>信息内容</u>
        <li>本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度。</li>
        <li>换行请按Shift+Enter，另起一段请按Enter</li></td>
      <td><textarea name="Content" style="display:none"><%=Content%></textarea>
        <iframe id="myeditor" src="../Ok_Editor/ewebeditor.htm?id=Content&amp;style=coolblue&amp;maxlen=300&amp;savepathfilename=Picture" frameborder="0" scrolling="No" width="95%" height="405"></iframe></td>
    </tr>
    <tr class="t_row">
      <td>&nbsp;</td>
      <td><input type="submit" value="提交"></td>
    </tr>
  </table>
</form>
<%
End Sub

Sub NewsShow()
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Pictures,Topis,Hits,Red,TitB,Content,DateAndTime,DescriptionWord,Keywords
	Sql="Select * From [News] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在!","New.asp",1
	Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		Title=Rs("Title")
		Author=Rs("Author")
		Url=Rs("Url")
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		Red=Rs("Red")
		TitB=Rs("TitB")
		Content=Rs("Content")
		DateAndTime=Rs("DateAndTime")
	End If
	Rs.Close
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title"><%=Title%></td>
  </tr>
  <tr class="t_rowh">
    <td>&nbsp;Location:&nbsp;&nbsp;
      <%Call ShowClassPath(ClassID)%></td>
    <td align="right">〖&nbsp;操作：<a href="News.asp?Action=Topis&ID=<%=ID%>">
      <%If Topis=1 Then%>
      <font color="#009933">推荐</font>
      <%Else%>
     推荐
      <%End If%>
      </a>&nbsp;<a href="News.asp?Action=Mod&ID=<%=ID%>">修改</a>&nbsp;<a href="News.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">删除</a>&nbsp;<a href="<%=Http_Referer%>">返回</a>&nbsp;〗</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,Title,ClassID,Hits,DateAndTime,Content,"News"%></td>
  </tr>
</table>
<Script language="JavaScript" type="text/JavaScript">
function Delete()
{
   if(confirm("删除内容后不能恢复！确定要删除吗？"))
     return true;
   else
     return false;
}
</Script>
<%
End Sub

Sub SaveAdd()
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Topis,Hot,Red,TitB,Content,DescriptionWord,Keywords
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"信息标题",100,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"自主SEO优化--搜索关键字：1",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"自主SEO优化--搜索关键字：2",2000,"012")
	Author=CheckChar(Trim(Request.Form("Author")),"信息出处",50,"12")
	Url=CheckChar(Trim(Request.Form("Url")),"出处链接",50,"12")
	Pic=CheckChar(Trim(Request.Form("Pic")),"默认图片",100,"12")
	Picture=CheckChar(Trim(Request.Form("Picture")),"所有图片",65535,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"信息置顶",0,"")
	Hot=CheckChar(Trim(Request.Form("Hot")),"热门信息",0,"")
	Red=CheckChar(Trim(Request.Form("Red")),"标题红色",0,"")
	TitB=CheckChar(Trim(Request.Form("TitB")),"标题加粗",0,"")
	Content=CheckChar(Request.Form("Content"),"信息内容",605535,"02")
	If ClassID<=0 Then ErrContent = ErrContent & "请先添加分类!"
	If ErrContent<>"" Then Msg ErrContent,"News.asp?Action=Add",1
	If Conn.Execute("Select Child From Class Where ClassID="&ClassID)(0)>0 Then
		Msg "所属分类含有子目录,请选择其他分类!","News.asp?Action=Add",1
	End If
	Rs.Open "Select * From [News]",conn,1,2
		Rs.Addnew
		ID=Rs("ID")
		Rs("Title")=Title
		Rs("Content")=Content
		If Author<>"" Then
			Rs("Author")=Author
			Rs("Url")=Url
		Else
			Rs("Author")=m_Config(0)
			Rs("Url")=m_Config(4)
		End If
		Rs("DescriptionWord")=DescriptionWord
		Rs("Keywords")=Keywords
		Rs("ClassID")=ClassID
		Rs("Pic")=Pic
		Rs("Picture")=Picture
		If Hot="True" Then
			Rs("Hits")=Cint(m_WebSet(3))
		Else
			Rs("Hits")=0
		End If
		If Red="True" Then
			Rs("Red")=1
		Else
			Rs("Red")=0
		End If
		If TitB="True" Then
			Rs("TitB")=1
		Else
			Rs("TitB")=0
		End If
		If Topis="True" Then
			Rs("Topis")=1
		Else
			Rs("Topis")=0
		End If
		Rs("DateAndTime")=Now()
		Rs.Update
	Rs.Close
	Msg "信息添加成功!","News.asp",0
End Sub

Sub SaveMod()
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Topis,Hot,Red,TitB,Content,DescriptionWord,Keywords
	ID=CheckChar(Trim(Request.Form("ID")),"信息编号",0,"03")
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"信息标题",100,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"自主SEO优化--搜索关键字：1",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"自主SEO优化--搜索关键字：2",2000,"012")
	Author=CheckChar(Trim(Request.Form("Author")),"信息出处",50,"12")
	Url=CheckChar(Trim(Request.Form("Url")),"出处链接",50,"12")
	Pic=CheckChar(Trim(Request.Form("Pic")),"默认图片",100,"12")
	Picture=CheckChar(Trim(Request.Form("Picture")),"所有图片",65535,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"信息置顶",0,"")
	Hot=CheckChar(Trim(Request.Form("Hot")),"热门信息",0,"")
	Red=CheckChar(Trim(Request.Form("Red")),"标题红色",0,"")
	TitB=CheckChar(Trim(Request.Form("TitB")),"标题加粗",0,"")
	Content=CheckChar(Request.Form("Content"),"信息内容",605535,"02")
	If ErrContent<>"" Then Msg ErrContent,"News.asp?Action=Mod&ID="&ID,1
	If Conn.Execute("Select Child From Class Where ClassID="&ClassID)(0)>0 Then
		Msg "所属分类含有子目录,请选择其他分类!","News.asp?Action=Mod&ID="&ID,1
	End If
	Rs.Open "Select * From [News] Where ID="&ID,conn,1,2
		Rs("Title")=Title
		Rs("Content")=Content
		If Author<>"" Then
			Rs("Author")=Author
			Rs("Url")=Url
		Else
			Rs("Author")=m_Config(0)
			Rs("Url")=m_Config(4)
		End If
		Rs("DescriptionWord")=DescriptionWord
		Rs("Keywords")=Keywords
		Rs("ClassID")=ClassID
		Rs("Pic")=Pic
		Rs("Picture")=Picture
		If Hot="True" Then
			Rs("Hits")=Cint(m_WebSet(3))
		Else
			Rs("Hits")=Rs("Hits")
		End If
		If Red="True" Then
			Rs("Red")=1
		Else
			Rs("Red")=0
		End If
		If TitB="True" Then
			Rs("TitB")=1
		Else
			Rs("TitB")=0
		End If
		If Topis="True" Then
			Rs("Topis")=1
		Else
			Rs("Topis")=0
		End If
		Rs("DateAndTime")=Now()
		Rs.Update
	Rs.Close
	Msg "信息修改成功!","News.asp",0
End Sub

Sub NewsDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"信息编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select * From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行删除操作!",Http_Referer,1
	Else
		DelFiles(Rs("Picture")&"")
		Rs.Delete
	End If
	Rs.Close
	Msg "信息ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub

Sub NewsSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"信息编号",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select * From [News] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行删除操作!",Http_Referer,1
	Else
		Do While Not Rs.Eof
		DelFiles(Rs("Picture")&"")
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "信息ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub

Sub NewsTopis()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"信息编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,Topis From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行操作!",Http_Referer,1
	Else
		If Rs("Topis")=1 Then
			Rs("Topis")=0
		Else
			Rs("Topis")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "信息ID:"&ID&"置顶设置成功!",Http_Referer,0
End Sub

Sub NewsHot()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"信息编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,Hits From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行操作!",Http_Referer,1
	Else
		If Rs("Hits")>=Cint(m_WebSet(3)) Then
			Rs("Hits")=Cint(Cint(m_WebSet(3))/2)
		Else
			Rs("Hits")=Cint(m_WebSet(3))
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "信息ID:"&ID&"热门设置成功!",Http_Referer,0
End Sub

Sub NewsRed()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"信息编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,Red From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行操作!",Http_Referer,1
	Else
		If Rs("Red")=1 Then
			Rs("Red")=0
		Else
			Rs("Red")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "信息ID:"&ID&"标题红色设置成功!",Http_Referer,0
End Sub

Sub NewsTitB()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"信息编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,TitB From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行操作!",Http_Referer,1
	Else
		If Rs("TitB")=1 Then
			Rs("TitB")=0
		Else
			Rs("TitB")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "信息ID:"&ID&"标题加粗设置成功!",Http_Referer,0
End Sub 

Bottom()
%>
