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
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">留言管理</td>
  </tr>
  <tr>
    <td colspan="2" class="t_rowh"><strong>说明</strong>：<br>当然提交人选择使用Email提示是，在回复留言后系统将自己把留言回复提示信发到提交人的信箱中。此功能需要服务器的支持。</td>
  </tr>
  <tr>
    <td colspan="2" class="t_rowh"><strong>可操作选项</strong>&nbsp;&nbsp;<a href="Guest.asp?Action=New">已回复留言</a>&nbsp;|&nbsp;<a href="Guest.asp">未回复的留言</a></td>
  </tr>
</table>
<%
Select Case Action
	Case "Revert"
		Call GuestRevert()
	Case "Del"
		Call GuestDel()
	Case "New"
		Call GuestList()
	Case Else
		Call GuestList()
End Select

Sub GuestList()
	If Action="New" Then
		Query_String="Action=New&"
		Sql="Select * From Guest Where Not(Revert Is Null) Order by ID Desc"
	Else
		Sql="Select * From Guest Where Revert Is Null Order by ID Desc"
	End If
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td class="t_title"><strong>留言列表</strong></td>
  </tr>
  <tr align="center" class="t_rowh">
    <td>没有未回复的留言，点击查看<a href="Guest.asp?Action=New" style="color:#990000"> [已回复留言]</a></td>
  </tr>
</table>
<%
	Else
		i=0
		Rs.PageSize = 5
		Rs.AbsolutePage = Cint(Page)
		Do While Not Rs.Eof And i < Rs.PageSize
%>
<form name="myform" method="post" action="Guest.asp?Action=Revert">
<table  border="0" align="center"cellpadding="0" cellspacing="1" width="680">
  
  <tr class="t_row">
    <td width="23%">留言人：<%=Rs("Sender")%></td>
    <td width="25%">邮箱：<a href="mailto:<%=Rs("Email")%>"><%=Rs("Email")%></a></td>
    <td width="52%">联系电话：<%=Rs("Tel")%></td>
  </tr>
  <tr class="t_row">
    <td colspan="3">具体要求：</td>
  </tr>
  <tr class="t_row">
    <td colspan="3">
	<textarea name="Content" rows="5" cols="50" id="Content" ><%=Rs("Content")%></textarea></td>
  </tr>


  <tr class="t_row">
    <td>留言时间：<%=Rs("DateAndTime")%>&nbsp;&nbsp;IP：<%=Rs("IP")%></td>
    <td colspan="2" >操作：&nbsp; <a href="Guest.asp?Action=Del&amp;ID=<%=rs("id")%>" style="color:#FF0000">删除留言</a> &nbsp;&nbsp;</td>
  </tr>
</table>
</form>
<%
		i = i + 1
		Rs.MoveNext
		Loop
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
<%
	ShowPage(0)
%>
</table>
<%
	End If
	Rs.Close
%>
<Script language="JavaScript" type="text/JavaScript">
function Delete(ID)
{
   if(confirm("删除留言后不能恢复！确定要删除吗？"))
   {
     location.href = 'Guest.asp?Action=Del&ID='+ID;
     return true;
   }
   else
     return false;
}
</Script>
<%
End Sub

Sub GuestRevert()
	Dim ID,Revert
	ID=CheckChar(Trim(Request.Form("ID")),"留言编号",0,"03")
	Revert=CheckChar(Trim(Request.Form("Revert")),"留言内容",65536,"012")
	If ErrContent<>"" Then Msg ErrContent,Http_Referer,1
	Sql="Select * From Guest Where ID="&ID
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "留言不存在!",Http_Referer,1
	Else
		Rs("Revert")=Revert
		Rs("RevertTime")=Now()
		Rs.Update
	End If
	Rs.Close
	Msg "留言回复成功!",Http_Referer,0
End Sub

Sub GuestDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"留言编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,Http_Referer,1
	Rs.Open "Select * From Guest Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "留言不存在,不可以执行删除操作!",Http_Referer,1
	Else
		Rs.Delete
	End If
	Rs.Close
	Msg "留言ID:"&ID&"已删除,不可以恢复!","Guest.asp",0
End Sub

Sub GuestSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"新闻编号",0,"0")
	If ErrContent<>"" Then Msg ErrContent,Http_Referer,1
	Conn.Execute("Delete From Guest Where ID In ("&ID&")")
	Msg "留言ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub
%>		 
		 
		<div class="bottom">
		<p>版权所有:南阳永正信息技术有限公司 地址:南阳市滨河西路88号滨河鑫苑4-4-5F 邮政编码:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
</html>
