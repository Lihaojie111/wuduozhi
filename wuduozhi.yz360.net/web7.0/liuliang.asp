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
		Call mainadd()
End Select
%>
<%	
Sub mainadd()
%>	
<body>
<!--头部-->
<div class="wramp">
  <!--#include file="head.asp"-->
 
  <!--左侧-->

  <div id="content">
  
 <!--#include file="left.asp"-->
	
	<div class="c-right" style="height:600px">
         
 <iframe  src="http://www.51.la" frameborder="0" scrolling="auto" width="100%"  height="100%"></iframe>
		 
		 
		<div class="bottom">
		<p>版权所有:南阳永正信息技术有限公司 地址:南阳市滨河西路88号滨河鑫苑4-4-5F 邮政编码:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
<%
end sub
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
%>

</html>
