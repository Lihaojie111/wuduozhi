<!--#include file="System.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����</title>
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
    font-family:"΢���ź�";} 
.top_menu a:hover{background:url(left_bg1.jpg);}
.sub_meuu a{ font-size:12px; color:#FFF;
	display:block;
	width:164px;
	padding-left:64px;
	height:36px;
	line-height:36px;
	font-family:"΢���ź�";
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
<!--ͷ��-->
<div class="wramp">
  <!--#include file="head.asp"-->
  <!--���-->

  <div class="clear"></div>
  
  
  <div id="content">
  
 <!--#include file="left.asp"-->
	
	<div class="c-right">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">���Թ���</td>
  </tr>
  <tr>
    <td colspan="2" class="t_rowh"><strong>˵��</strong>��<br>��Ȼ�ύ��ѡ��ʹ��Email��ʾ�ǣ��ڻظ����Ժ�ϵͳ���Լ������Իظ���ʾ�ŷ����ύ�˵������С��˹�����Ҫ��������֧�֡�</td>
  </tr>
  <tr>
    <td colspan="2" class="t_rowh"><strong>�ɲ���ѡ��</strong>&nbsp;&nbsp;<a href="Guest.asp?Action=New">�ѻظ�����</a>&nbsp;|&nbsp;<a href="Guest.asp">δ�ظ�������</a></td>
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
    <td class="t_title"><strong>�����б�</strong></td>
  </tr>
  <tr align="center" class="t_rowh">
    <td>û��δ�ظ������ԣ�����鿴<a href="Guest.asp?Action=New" style="color:#990000"> [�ѻظ�����]</a></td>
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
    <td width="23%">�����ˣ�<%=Rs("Sender")%></td>
    <td width="25%">���䣺<a href="mailto:<%=Rs("Email")%>"><%=Rs("Email")%></a></td>
    <td width="52%">��ϵ�绰��<%=Rs("Tel")%></td>
  </tr>
  <tr class="t_row">
    <td colspan="3">����Ҫ��</td>
  </tr>
  <tr class="t_row">
    <td colspan="3">
	<textarea name="Content" rows="5" cols="50" id="Content" ><%=Rs("Content")%></textarea></td>
  </tr>


  <tr class="t_row">
    <td>����ʱ�䣺<%=Rs("DateAndTime")%>&nbsp;&nbsp;IP��<%=Rs("IP")%></td>
    <td colspan="2" >������&nbsp; <a href="Guest.asp?Action=Del&amp;ID=<%=rs("id")%>" style="color:#FF0000">ɾ������</a> &nbsp;&nbsp;</td>
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
   if(confirm("ɾ�����Ժ��ָܻ���ȷ��Ҫɾ����"))
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
	ID=CheckChar(Trim(Request.Form("ID")),"���Ա��",0,"03")
	Revert=CheckChar(Trim(Request.Form("Revert")),"��������",65536,"012")
	If ErrContent<>"" Then Msg ErrContent,Http_Referer,1
	Sql="Select * From Guest Where ID="&ID
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "���Բ�����!",Http_Referer,1
	Else
		Rs("Revert")=Revert
		Rs("RevertTime")=Now()
		Rs.Update
	End If
	Rs.Close
	Msg "���Իظ��ɹ�!",Http_Referer,0
End Sub

Sub GuestDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"���Ա��",0,"03")
	If ErrContent<>"" Then Msg ErrContent,Http_Referer,1
	Rs.Open "Select * From Guest Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "���Բ�����,������ִ��ɾ������!",Http_Referer,1
	Else
		Rs.Delete
	End If
	Rs.Close
	Msg "����ID:"&ID&"��ɾ��,�����Իָ�!","Guest.asp",0
End Sub

Sub GuestSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"���ű��",0,"0")
	If ErrContent<>"" Then Msg ErrContent,Http_Referer,1
	Conn.Execute("Delete From Guest Where ID In ("&ID&")")
	Msg "����ID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub
%>		 
		 
		<div class="bottom">
		<p>��Ȩ����:����������Ϣ�������޹�˾ ��ַ:�����б�����·88�ű�����Է4-4-5F ��������:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
</html>
