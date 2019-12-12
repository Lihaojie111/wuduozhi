<!--#include file="System.asp"-->
<!--#include file="../inc/Md5.asp"-->
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
.xiu_bt{
padding-left:20px;
font:"宋体",22px;
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
<!--头部-->
<div class="wramp">
    <!--#include file="head.asp"-->
  <!--左侧-->
  <div id="content">
  
 <!--#include file="left.asp"-->
 <div class="c-right">
<div class="c-rgt01" style="margin-bottom:50px;">
		<div style="float:left; padding-left:20px; padding-top:10px">您现在的位置：后台首台-修改密码</div>
		
		<div style="float:right; padding-right:100px;"> <form action="http://www.baidu.com/baidu" target="_blank">
<table bgcolor="#f1f1f1"><tr><td>
<input name=tn type=hidden value=baidu>
<img src="http://img.baidu.com/img/logo-80px.gif" alt="Baidu" align="bottom" border="0" />
<input type=text name=word size=30>
<input type="submit" value="百度搜索">
</td></tr></table>
</form></div>

		</div>
 <form name="myform" method="post" action="password.asp?Action=SaveModPsw">
<table width="93%"  align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="rgt03-bt">修改管理员密码</td>
  </tr>
  
  <tr class="t_row">
    <td width="40%" align="right">原始密码:</td>
    <td width="60%"><input name="OldPsw" type="password" id="OldPsw" size="31"></td>
  </tr>
  <tr class="t_row">
    <td align="right">新 密 码:</td>
    <td><input name="NewPsw" type="password" id="NewPsw" size="31"></td>
  </tr>
  <tr class="t_row">
    <td align="right">确认密码:</td>
    <td><input name="NewPsw2" type="password" id="NewPsw2" size="31"></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" name="Submit2" value="提交"></td>
  </tr>
</table>
</form>


<div class="c-rgt04" style="margin-top:100px;">
			<p class="rgt03-bt">常用工具下载</p>
			<div class="img">
			<ul>
			<li><img src="images/QQ.jpg" alt="qq" />
			<p class="padd2">&nbsp;QQ</p>
			<a href="http://dl_dir.qq.com/qqfile/qq/QQ2011/QQ2011(QQProtect2.1).exe" class="link">点击下载</a></li>
			<li><img src="images/迅雷.jpg" alt="xunlei" /><p class="padd2">迅雷</p><a  href="http://down.sandai.net/thunder7/Thunder7.2.5.3364.exe" class="link">点击下载</a></li>
			<li><img src="images/酷我音乐.jpg" alt="music" /><p class="padd5">酷我音乐盒</p><a  href="http://down.kuwo.cn/mbox/kwmusic2012.exe" class="link">点击下载</a></li>
			<li><img src="images/美图秀秀.jpg" alt="meituxiuxiu" /><p class="padd4">美图秀秀</p><a href="http://xiuxiu.dl.meitu.com/XiuXiu_360Setup_3.0.9.exe" class="link">点击下载</a></li>
			<li><img src="images/feixin.jpg" alt="feixin" /><p class="padd2">飞信</p><a href="http://fetion.chinacache.net/alading/Fetion2011December.exe" class="link">点击下载</a></li>
			</ul>
			</div>
		</div>
 </div> 
 
 <div class="clear"></div>
  </div>
</div>
</body>


<%
end sub
Sub SaveModPsw()
	Dim OldPsw,NewPsw,NewPsw2
	OldPsw=CheckChar(Trim(Request.Form("OldPsw")),"原始密码",0,"0")
	NewPsw=CheckChar(Trim(Request.Form("NewPsw")),"新密码",0,"0")
	NewPsw2=CheckChar(Trim(Request.Form("NewPsw2")),"确认密码",0,"0")
	If NewPsw<>NewPsw2 Then ErrContent = ErrContent & "两次密码输入不相同!"
	If ErrContent<>"" then Msg ErrContent,"Main.asp?Action=ModPsw",1
	
	Rs.Open "Select Psw From Admin Where Admin='"& Session("m_Admin") &"'",Conn,1,2
	If Rs("Psw")<> Md5(OldPsw) Then
		Msg "原始密码错误!","",1
	Else
		Rs("Psw")=Md5(NewPsw)
		Rs.Update
	End If
	Rs.Close
	Msg "密码修改成功,请使用新密码!","Main.asp",0
End Sub

Sub SaveAddadmin()
    Dim Admin,Password1,Password2
	Admin=request("Admin")
	Password1=request("Password1")
	Password2=request("Password2")
	if Password1="" then
	Msg "密码不能为空，请重新输入!","Main.asp?Action=Addadmin",1
		end if 
	if Password1<>Password2 then 
	Msg "两次密码不一致，请重新输入!","Main.asp?Action=Addadmin",1
	end if
		
if Admin="" then
Msg "用户名不能为空!","Main.asp?Action=Addadmin",1
end if 
	sql="select * from [Admin] where Admin='"&trim(request("Admin"))&"'"
rs.open sql,conn,1,1
if not rs.eof then
%>
<script language="javascript">
alert("该用户名已被注册!");
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
	Msg "管理员添加成功!","Main.asp?Action=Addadmin",0
End Sub


Sub UsersDel()
	Dim ID
	ID=Request.QueryString("ID")
	If ErrContent<>"" Then Msg ErrContent,"Users.asp",1
	Rs.Open "Select * From [Admin] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行删除操作!",Http_Referer,1
	Else

		Rs.Delete
	End If
	Rs.Close
	Msg "信息ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub


Sub UsersSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"新闻编号",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"Users.asp",1
	Rs.Open "Select * From [Admin] Where ID In ("&ID&")",Conn,1,2
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
%>
</html>
