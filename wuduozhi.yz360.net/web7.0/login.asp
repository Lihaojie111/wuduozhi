<!--#include file="System.asp"-->
<!--#include file="../inc/md5.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������WEB7.0ϵͳ</title>
<link href="css/mystyle.css" type="text/css" rel="stylesheet" />
</head>

<body>
<div class="login">
<%
Select Case Request.QueryString("Action")
	Case "ChkLogin"
		Call ChkLogin()
	Case "Logout"
		Call Logout()
	Case "chongzhi"
		Call chongzhi()	
	Case Else
		Call Login()
End Select

Sub Login()
	Randomize timer
	Session("m_Code") = Int(Rnd*8998)+1000
	Session("m_Admin") = ""
	CheckData "Login","Admin,�û���,,0|Psw,����,,0|Code,��֤��,==4,012"
%>	
<div class="tp">
<form method="post" action="Login.asp?Action=ChkLogin"  name="Login" onSubmit="return CheckData()">
<table  width="235" height="125" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="77" align="center">�û�����</td>
    <td width="158"><input name="Admin" type="text" id="Admin" size="24" maxlength="30" style="BORDER: #7187b8 1px solid; background:#f5fafe; width:160px; height:21px;"></td>
  </tr>
  <tr>
    <td align="center" >��&nbsp;&nbsp;�룺</td>
    <td><input  name="Psw" type="password" id="Psw" size="24" maxlength="30" style="BORDER: #7187b8 1px solid;background:#f5fafe; width:160px; height:21px;"></td>
  </tr>
  <tr>
    <td align="center">��֤�룺</td>
    <td height="30">
					    <input  name="Code" type="text" id="Code" size="4" maxlength="4" style="BORDER: #7187b8 1px solid; background:#f5fafe; width:61px; height:21px;">
						<input type="text" value="<%=Session("m_Code")%>" style="BORDER: #7187b8 1px solid; background:#f5fafe; width:49px; height:21px; text-align:center; font-size:18px; font-weight:bold; color:#0e4f93">					 </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input type="image"  width="73" height="24" src="images/bt1.jpg">
	<input  type="image" width="73" height="24" src="images/bt2.jpg" onClick="javascript:document.Login.reset();return  false;" >
	</td>
  </tr>
</table>
</form>
</div>
<%
	Bottom()
End Sub

Sub ChkLogin()
	Dim Admin,Psw,Code
	Admin=CheckChar(Trim(Request.Form("Admin")),"�û���",20,"012")
	Psw=CheckChar(Trim(Request.Form("Psw")),"����",30,"02")
	Code=CheckChar(Trim(Request.Form("Code")),"������",0,"03")
	If ErrContent<>"" then Msg ErrContent,"Index.asp?Action=Login",1
	If Cint(Code) <> Session("m_Code") Then Msg "�����벻��ȷ!","",1
	
	Rs.Open "Select * From Admin Where Admin='"& Admin &"'",Conn,1,1
	If Rs("Admin") <> Admin Then
		Msg "����Ա�ʺŴ���!","",1
	ElseIf Rs("Psw") <> MD5(Psw) Then
		Msg "����Ա�������!","",1
	Else
		Session("m_Admin")=Rs("Admin")
		Session("Grade")=Rs("Grade")
	
	End If
	Rs.Close
	Response.Redirect("Index.asp")
End Sub



Sub Logout()
	Session("m_Admin")=""
	Session.Abandon()
	Response.Redirect("Index.asp")
	Response.End
End Sub
%>
<div class="foot">�ͷ����ߣ�0377-66076580</div>
</div>

<div class="bot"> COPYRIGHT&copy;1999-2012 ������������Ϣ�������޹�˾</div>
</body>
</html>
