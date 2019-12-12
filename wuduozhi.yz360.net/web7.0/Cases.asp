<!--#include file="System.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台界面</title>
<style> 
a{ text-decoration:none;}  
.top_menu a{
	display:block;
	height:38px;
	 width:174px;
	 padding-left:54px;
	line-height:36px;
	font-size:14px; 
	color:#3f3f3f;
    font-family:"微软雅黑";} 
.top_menu a:hover{background:url(left_bg1.jpg);}
.sub_meuu a{ font-size:12px; color:#FFF;
	display:block;
	width:164px;
	padding-left:64px;
	height:36px;
	line-height:36px;
	font-family:"微软雅黑";
} 
.hover:hover{background:url(bg3.jpg);
display:block;
color:#dee318;}
</style> 
<script language="javascript"> 
tempj=2; 
function showed(tempi) 
{ 
if(document.getElementById("show"+tempj.toString()).style.display==''&&tempi!=tempj) 
{ 
document.getElementById("show"+tempj.toString()).style.display='none'; 
} 
if(document.getElementById("show"+tempi.toString()).style.display=='none') 
{ 
document.getElementById("show"+tempi.toString()).style.display=''; 
} 
else 
{ 
document.getElementById("show"+tempi.toString()).style.display='none'; 
} 
tempj=tempi; 
} 
</script> 
<link type="text/css" rel="stylesheet" href="css.css" />
</head>
<body>
<!--头部-->
<div class="wramp">
  <!--#include file="head.asp"-->
  <!--左侧-->

  <div class="clear"></div>
  
  
  <div id="content">
  
 <!--#include file="left.asp"-->
	
	<div class="c-right">
         <div class="xwgl">
		   
		   <div class="xw_nr">
<%

Layout="Cases"
%>
<script language="javascript" src="images/Admin_Js.Js"></script>

<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">英诺专卖</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2"><strong>操作选项</strong>：<a href="Cases.asp">英诺专卖</a>&nbsp;|&nbsp;<a href="Cases.asp?Action=Add">添加专卖</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=Cases">分类管理</a></td>
  </tr>
  <tr class="t_rowh">
    <form method="get" action="Cases.asp">
    <td width="60%">快速搜索：&nbsp;<input name="Action" type="hidden" id="Action" value="Search"><input name="Word" type="text" id="Word" size="48"></td>
    </form>
	<td width="40%" align="right">查看分类：&nbsp;<select name="ClassID" id="ClassID" onchange="window.location='Cases.asp?ClassID='+this.options[this.selectedIndex].value+'';"><option value="0">跳转分类至</option><%SelectClass(Cint(ClassID))%></select>&nbsp;&nbsp;&nbsp;</td>
  </tr>
</table>
<%
Select Case Action
	Case "Add"
		Call CasesAdd()
	Case "Mod"
		Call CasesMod()
	Case "Show"
		Call CasesShow()
	Case "Del"
		Call CasesDel()
	Case "SDel"
		Call CasesSDel()
	Case "Topis"
		Call CasesTopis()
	Case "SaveMod"
		Call SaveMod()
	Case "SaveAdd"
		Call SaveAdd()
	Case Else
		Call CasesList()
End Select

Sub CasesList()
	If Action="Search" Then
		Query_String="Action=Search&Word="&Request("Word")&"&"
		Sql="Where Title Like '%"&Request("Word")&"%'"
	ElseIf Cint(ClassID)>0 Then
		Dim RsTemp,ChildID
		Set RsTemp=Conn.Execute("Select ParentPath From [Class] Where ClassID="&ClassID)
		If Not RsTemp.Eof Then
			Dim ParentPath
			ParentPath = RsTemp("ParentPath")
		Else
			Msg "请选择要查询的分类","Cases.asp",1
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
	Sql = "Select * From [Cases] "&Sql&" Order By ID Desc"
	Rs.Open Sql,Conn,1,1
%>
<form name="myform" method="post" action="Cases.asp?Action=SDel">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr align="center" class="t_title">
    <td width="5%">选择</td>
    <td>标题</td>
    <td width="18%">分类</td>
    <td width="10%">点击次数</td>
    <td width="10%">操作</td>
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
    <td><a href="Cases.asp?Action=Topis&ID=<%=Rs("ID")%>"><%If Rs("Topis")=1 Then%><font color="#CC0000">荐</font><%Else%>荐<%End If%></a>&nbsp;<a href="Cases.asp?Action=Mod&ID=<%=Rs("ID")%>">修改</a>&nbsp;<a href="Cases.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">删除</a></td>
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

Sub CasesAdd()
%>
<script language="javascript">
// 参数说明
// s_Type : 文件类型，可用值为"image","flash","media","file"
// s_Link : 文件上传后，用于接收上传文件路径文件名的表单名
// s_Thumbnail : 文件上传后，用于接收上传图片时所产生的缩略图文件的路径文件名的表单名，当未生成缩略图时，返回空值，原图用s_Link参数接收，此参数专用于缩略图
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//以下style=coolblue,值可以依据实际需要修改为您的样式名,通过此样式的后台设置来达到控制允许上传文件类型及文件大小
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=Cases&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>

<form name="myform" method="post" action="Cases.asp?Action=SaveAdd" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">添加专卖</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>所属分类</u></td>
    <td width="80%"><select name="ClassID" id="ClassID"><%SelectClass(0)%></select>&nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
  </tr>
  <tr class="t_row">
    <td><u>专卖标题</u></td>
    <td><input name="Title" type="text" id="Title" size="48">&nbsp;<font color="#FF0000">＊</font> ≤ 100个字符&nbsp;&nbsp;&nbsp;&nbsp;<input name="Topis" type="checkbox" id="Topis" value="True"><label for="Topis">推荐专卖</label></td>
  </tr>
  <tr class="t_row">
    <td><u></u></td>
    <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="" size="48"></td>
  </tr>
  <tr class="t_row">
    <td><u>搜索关键字</u></td>
    <td><input name="Keywords" type="text" class="STYLE1" id="Keywords" value="输入专卖关键字" size="48"></td>
  </tr>
  <tr class="t_row">
    <td><u>专卖截图</u></td>
    <td><input name="Pic" type="text" id="Pic" size="48" value="/uploadfile/cases/20090806092504651.jpg">
    &nbsp;&nbsp;<input type=button value="上传图片..." onClick="showUploadDialog('image', 'myform.Pic', '')">&nbsp;&nbsp;<a href="#" target="_blank" id="yulan" onclick="this.href=document.getElementById('Pic').value">预览图片</a></td>
  </tr>
  <tr class="t_row">
    <td><u>专卖介绍</u><li>本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度。</li><li>换行请按Shift+Enter，另起一段请按Enter</li></td>
    <td><textarea name="Content1" style="display:none"><%=Content%></textarea> <iframe id="myeditor" src="../Ok_Editor/ewebeditor.htm?id=Content1&amp;style=Cases&amp;maxlen=300" frameborder="0" scrolling="No" width="95%" height="405"></iframe></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input name="submit" type="submit" value="提交" /></td>
  </tr>
</table>
</form>
<%
End Sub

Sub CasesMod()
	Dim ID,ClassID,Title,Price,Pic,Picture,Pictures,Topis,Hits,Content,TempPicture,DescriptionWord,Keywords
	Sql="Select * From Cases Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "专卖不存在!",Http_Referer,1
	Else
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		DescriptionWord=Rs("DescriptionWord")
		Keywords=Rs("Keywords")
		Title=Rs("Title")
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		Content=Rs("Content")
	End If
	Rs.Close
%>
<script language="javascript">
// 参数说明
// s_Type : 文件类型，可用值为"image","flash","media","file"
// s_Link : 文件上传后，用于接收上传文件路径文件名的表单名
// s_Thumbnail : 文件上传后，用于接收上传图片时所产生的缩略图文件的路径文件名的表单名，当未生成缩略图时，返回空值，原图用s_Link参数接收，此参数专用于缩略图
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//以下style=coolblue,值可以依据实际需要修改为您的样式名,通过此样式的后台设置来达到控制允许上传文件类型及文件大小
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=Cases&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>
<form name="myform" method="post" action="Cases.asp?Action=SaveMod" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">修改专卖</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>所属分类</u></td>
    <td width="80%"><input name="ID" type="hidden" value="<%=ID%>"><select name="ClassID" id="ClassID"><%SelectClass(ClassID)%></select>&nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
  </tr>
  <tr class="t_row">
    <td><u>专卖标题</u></td>
    <td><input name="Title" type="text" id="Title" value="<%=Title%>" size="48">&nbsp;<font color="#FF0000">＊</font> ≤ 100个字符&nbsp;&nbsp;&nbsp;&nbsp;<input name="Topis" type="checkbox" id="Topis" value="True"<%If Topis=1 Then%> checked<%End If%>>&nbsp;<label for="Topis">推荐专卖</label></td>
  </tr>
  <tr class="t_row">
    <td><u></u></td>
    <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="<%=DescriptionWord%>" size="48"></td>
  </tr>
  <tr class="t_row">
    <td><u>搜索关键字</u></td>
    <td><input name="Keywords" type="text" id="Keywords" value="<%=Keywords%>" size="48"></td>
  </tr>
  <tr class="t_row">
    <td><u>专卖截图</u></td>
    <td><input name="Pic" type="text" id="Pic" size="48" value="<%=Pic%>">&nbsp;&nbsp;<input type=button value="上传图片..." onClick="showUploadDialog('image', 'myform.Pic', '')">&nbsp;&nbsp;<a href="#" target="_blank" id="yulan" onclick="this.href=document.getElementById('Pic').value">预览图片</a></td>
  </tr>
  
  <tr class="t_row">
    <td><u>专卖介绍</u><li>本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度。</li><li>换行请按Shift+Enter，另起一段请按Enter</li></td>
    <td><textarea name="Content2" style="display:none"><%=Content%></textarea> <iframe id="myeditor" src="../Ok_Editor/ewebeditor.htm?id=Content2&amp;style=Cases&amp;maxlen=300" frameborder="0" scrolling="No" width="95%" height="405"></iframe></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="提交"></td>
  </tr>
</table>
</form>
<%
End Sub

Sub CasesShow()
	Dim ID,ClassID,Title,Pic,Picture,Pictures,Topis,Hits,Content,DateAndTime,DescriptionWord,Keywords
	Sql="Select * From [Cases] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "专卖不存在!",Http_Referer,1
	Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		Title=Rs("Title")
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
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
    <td>&nbsp;Locahost:&nbsp;&nbsp;<%Call ShowClassPath(ClassID)%></td>
    <td align="right">〖&nbsp;操作：<a href="Cases.asp?Action=Topis&ID=<%=ID%>"><%If Topis=1 Then%><font color="#CC0000">荐</font><%Else%>荐<%End If%></a>&nbsp;<a href="Cases.asp?Action=Mod&ID=<%=ID%>">修改</a>&nbsp;<a href="Cases.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">删除</a>&nbsp;<a href="<%=Http_Referer%>">返回</a>&nbsp;〗</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,Title,ClassID,Hits,DateAndTime,"<div align=""center""><img src="""&Pic&"""></div><br>"&Replace(Replace(Content,Chr(10),"<br>"),Chr(13),"&nbsp;")&"<br><br><strong>说明书下载</strong>:"&Picture,"Cases"%></td>
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
	Dim ID,ClassID,Title,Price,Pic,Picture,Topis,Content,DescriptionWord,Keywords
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"专卖名称",100,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"搜索关键字1",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"搜索关键字2",2000,"012")
	Picture=CheckChar(Trim(Request.Form("Picture")),"说明书",100,"12")
	Pic=CheckChar(Trim(Request.Form("Pic")),"图片",255,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"推荐专卖",0,"")
	Content=CheckChar(Request.Form("Content1"),"专卖介绍",65535,"02")
	If ClassID<=0 Then ErrContent = ErrContent & "请先添加分类!"
	If ErrContent<>"" Then Msg ErrContent,"Cases.asp?Action=Add",1
	Rs.Open "Select * From [Cases]",conn,1,2
		Rs.Addnew
		ID=Rs("ID")
		Rs("Title")=Title
		Rs("Content")=Content
		Rs("ClassID")=ClassID
		Rs("Pic")=Pic
		Rs("DescriptionWord")=DescriptionWord
		Rs("Keywords")=Keywords
		Rs("Picture")=Picture
		If Topis="True" Then
			Rs("Topis")=1
		Else
			Rs("Topis")=0
		End If
		Rs("Hits")=0
		Rs("DateAndTime")=Now()
		Rs.Update
	Rs.Close
	Msg "专卖添加成功!","Cases.asp",0
End Sub

Sub SaveMod()
	Dim ID,ClassID,Title,Price,Pic,Picture,Topis,Content,DescriptionWord,Keywords
	ID=CheckChar(Trim(Request.Form("ID")),"专卖编号",0,"03")
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"专卖名称",100,"012")
	Picture=CheckChar(Trim(Request.Form("Picture")),"说明书",100,"12")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"搜索关键字1",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"搜索关键字2",2000,"012")
	Pic=CheckChar(Trim(Request.Form("Pic")),"图片",255,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"推荐专卖",0,"")
	Content=CheckChar(Request.Form("Content2"),"专卖介绍",65535,"02")
	If ErrContent<>"" Then Msg ErrContent,"Cases.asp?Action=Mod&ID="&ID,1
	Rs.Open "Select * From [Cases] Where ID="&ID,conn,1,2
		Rs("Title")=Title
		Rs("Content")=Content
		Rs("ClassID")=ClassID
		Rs("DescriptionWord")=DescriptionWord
		Rs("Keywords")=Keywords
		Rs("Pic")=Pic
		Rs("Picture")=Picture
		If Topis="True" Then
			Rs("Topis")=1
		Else
			Rs("Topis")=0
		End If
		Rs("DateAndTime")=Now()
		Rs.Update
	Rs.Close
	Msg "专卖修改成功!","Cases.asp",0
End Sub

Sub CasesDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"专卖编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Cases.asp",1
	Rs.Open "Select * From [Cases] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "专卖不存在,不可以执行删除操作!",Http_Referer,1
	Else
'		DelFiles(Rs("Picture")&"")
		Rs.Delete
	End If
	Rs.Close
	Msg "专卖ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub

Sub CasesSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"专卖编号",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"Cases.asp",1
	Rs.Open "Select * From [Cases] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "专卖不存在,不可以执行删除操作!",Http_Referer,1
	Else
		Do While Not Rs.Eof
'		DelFiles(Rs("Picture")&"")
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "专卖ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub

Sub CasesTopis()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"专卖编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Cases.asp",1
	Rs.Open "Select ID,Topis From [Cases] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "专卖不存在,不可以执行操作!",Http_Referer,1
	Else
		If Rs("Topis")=1 Then
			Rs("Topis")=0
		Else
			Rs("Topis")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "专卖ID:"&ID&"置顶设置成功!",Http_Referer,0
End Sub


%>



		   </div>
		 </div>
		<div class="bottom">
		<p>版权所有:南阳永正信息技术有限公司 地址:南阳市滨河西路88号滨河鑫苑4-4-5F 邮政编码:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
</html>
