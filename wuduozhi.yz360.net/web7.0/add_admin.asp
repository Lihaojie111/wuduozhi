<!--#include file="System.asp"-->
<!--#include file="../inc/Md5.asp"-->
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
.xiu_bt{
padding-left:20px;
font:"����",22px;
color:#66aa42;}
.xiu2{padding-left:150px;
padding-top:50px;
	}
.wramp{background-color:#f1f1f1;}
.red{color:red;}
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
<%

Select Case Action
	Case "ModPsw"
		Call ModPsw()
	Case "SaveModPsw"
		Call SaveModPsw()
	Case "Addadmin"
	    Call Addadmin()
	Case "SaveAddadmin"
	    Call SaveAddadmin()	
	Case "list"
	    Call UsersList()
	Case "Del"
	    Call UsersDel()	
	Case "SDel"
		Call UsersSDel()
	Case Else
		Call Main()
End Select

Sub Main()
%>
<body>
<!--ͷ��-->
<div class="wramp">
  <!--#include file="head.asp"-->
  <!--���-->
  <div id="content">
  
 <!--#include file="left.asp"-->
 <div class="c-right">
 <!--#include file="baidu.asp"-->
 <form name="form1" action="?Action=SaveAddadmin" method="post">
<table width="93%" border="0" cellspacing="0" cellpadding="0" class="t_border" align="center">
  <tr>
    <td colspan="2"   class="rgt03-bt"> �����ͨ����Ա</td>
   
  </tr>
  <tr  class="t_row">
    <td width="43%"  align="right">�û�����</td>
    <td width="57%"><input name="Admin" type="text" size="31"  id="Admin" /></td>
  </tr>
  <tr class="t_row">
    <td align="right">���룺</td>
    <td><input name="Password1" type="password" size="31"  id="Password1" /></td>
  </tr>
  <tr class="t_row">
    <td align="right">ȷ�����룺</td>
    <td><input name="Password2" type="password" size="31"  id="Password2"  /></td>
  </tr>
    <tr class="t_row">
    <td></td>
    <td><input name="Submit3" type="submit"  value="�ύ" id="Submit3"/></td>
  </tr>
</table>
</form>


<!--#include file="gongju.asp"-->
 </div> 
 
 <div class="clear"></div>
  </div>
</div>
</body>


<%
end sub
Sub SaveModPsw()
	Dim OldPsw,NewPsw,NewPsw2
	OldPsw=CheckChar(Trim(Request.Form("OldPsw")),"ԭʼ����",0,"0")
	NewPsw=CheckChar(Trim(Request.Form("NewPsw")),"������",0,"0")
	NewPsw2=CheckChar(Trim(Request.Form("NewPsw2")),"ȷ������",0,"0")
	If NewPsw<>NewPsw2 Then ErrContent = ErrContent & "�����������벻��ͬ!"
	If ErrContent<>"" then Msg ErrContent,"Main.asp?Action=ModPsw",1
	
	Rs.Open "Select Psw From Admin Where Admin='"& Session("m_Admin") &"'",Conn,1,2
	If Rs("Psw")<> Md5(OldPsw) Then
		Msg "ԭʼ�������!","",1
	Else
		Rs("Psw")=Md5(NewPsw)
		Rs.Update
	End If
	Rs.Close
	Msg "�����޸ĳɹ�,��ʹ��������!","Main.asp",0
End Sub

Sub SaveAddadmin()
    Dim Admin,Password1,Password2
	Admin=request("Admin")
	Password1=request("Password1")
	Password2=request("Password2")
	if Password1="" then
	Msg "���벻��Ϊ�գ�����������!","Main.asp?Action=Addadmin",1
		end if 
	if Password1<>Password2 then 
	Msg "�������벻һ�£�����������!","Main.asp?Action=Addadmin",1
	end if
		
if Admin="" then
Msg "�û�������Ϊ��!","Main.asp?Action=Addadmin",1
end if 
	sql="select * from [Admin] where Admin='"&trim(request("Admin"))&"'"
rs.open sql,conn,1,1
if not rs.eof then
%>
<script language="javascript">
alert("���û����ѱ�ע��!");
history.back();
</script>
<%
end if
rs.close

	Rs.Open "Select * From [Admin]",conn,1,2
		Rs.Addnew
		ID=Rs("ID")		
		Rs("Admin")=Admin
		Rs("Psw")=Md5(Password1)
		Rs("Grade")="0"
		Rs.Update
	Rs.Close
	Msg "����Ա��ӳɹ�!",Http_Referer,0
End Sub


Sub UsersDel()
	Dim ID
	ID=Request.QueryString("ID")
	If ErrContent<>"" Then Msg ErrContent,"Users.asp",1
	Rs.Open "Select * From [Admin] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ��ɾ������!",Http_Referer,1
	Else

		Rs.Delete
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub


Sub UsersSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"���ű��",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"adminlist.asp",1
	Rs.Open "Select * From [Admin] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ��ɾ������!",Http_Referer,1
	Else
		Do While Not Rs.Eof
		DelFiles(Rs("Picture")&"")
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub
%>
</html>
