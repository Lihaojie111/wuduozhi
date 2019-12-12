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
.top_menu a:hover{background:url( left_bg1.jpg);}
.sub_meuu a{ font-size:12px; color:#FFF;
	display:block;
	width:164px;
	padding-left:64px;
	height:36px;
	line-height:36px;
	font-family:"微软雅黑";
} 
.hover:hover{background:url( bg3.jpg);
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
<link type="text/css" rel="stylesheet" href=" css.css" />
</head>
<%

Select Case Action
	Case "addseo"
		Call addseo()
	Case "Modseo"
		Call Modseo()

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
		<!--#include file="baidu.asp"-->
<%
	
	Sql="Select * from seo "
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "信息不存在!",Http_Referer,1
	Else
		ID=Rs("ID")
		DescriptionWord=Rs("DescriptionWord")
		Keywords1=Rs("Keywords1")
		Keywords2=Rs("Keywords2")
		Keywords3=Rs("Keywords3")
		Keywords4=Rs("Keywords4")
		Title=Rs("Title")
		Content=Rs("Content")
	End If
	Rs.Close
%>	
 <form name="form1" action="seo.asp?Action=addseo" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:50px; margin-bottom:50px" >
  <tr>
    <td height="32"  colspan="6"  > <div  class="rgt03-bt" style="width:90%; margin-left:26px; "> 首页SEO优化</div></td>
    </tr>
  <tr>
    <td width="128" height="33" align="right">*网站首页标题：</td>
    <td colspan="4"><input name="title" type="text" id="title" size="43"  value="<%=title%>"></td>
    <td width="319"><font color="#999999">&nbsp;一般不超过30个汉字。</font></td>
  </tr>
  <tr>
    <td height="37" align="right" width="128">*优化关键词：</td>
    <td width="113"><input name="Keywords1" type="text" id="Keywords1"  size="13"   value="<%=Keywords1%>">-</td>
    <td width="110"><input name="Keywords2" type="text" id="Keywords2"  size="13"  value="<%=Keywords2%>">-</td>
    <td width="115"><input name="Keywords3" type="text" id="Keywords3"  size="13"   value="<%=Keywords3%>">-</td>
    <td width="93"><input name="Keywords4" type="text" id="Keywords4"  size="13"  value="<%=Keywords4%>"></td>
	<td><font color="#999999">&nbsp;一般不超过30个汉字。</font></td>
  </tr>
  <tr>
    <td height="29" align="right">*网站优化描述：</td>
    <td colspan="5"><textarea name="DescriptionWord" cols="62" rows="5"   wrap="physical" id="DescriptionWord" onkeyup="textLimit(this, 200);"><%=DescriptionWord%></textarea></td>
  </tr>
  <tr>
    <td height="36">&nbsp;</td>
    <td><input type="submit" name="button" id="button" value="提交" /></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
  </tr>
</table>

</form>		
		<!--#include file="gongju.asp"-->
		
		<div class="bottom">
		<p>版权所有:南阳永正信息技术有限公司 地址:南阳市滨河西路88号滨河鑫苑4-4-5F 邮政编码:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
<%
end sub
Sub addseo()
	Dim title,Keywords1,Keywords2,Keywords3,Keywords4,txtContent
	title=Request.Form("title")
	Keywords1=Request.Form("Keywords1")
	Keywords2=Request.Form("Keywords2")
	Keywords3=Request.Form("Keywords3")
	Keywords4=Request.Form("Keywords4")
	DescriptionWord=Request.Form("DescriptionWord")
	
	Rs.Open "Select * From seo",conn,1,2
		Rs("Title")=Title
		Rs("Keywords1")=Keywords1
		Rs("Keywords2")=Keywords2
		Rs("Keywords3")=Keywords3		
		Rs("Keywords4")=Keywords4
		Rs("DescriptionWord")=DescriptionWord
	
		Rs("DateAndTime")=Now()
		Rs.Update
	Rs.Close
	Msg "提交成功!","seo.asp",0

end sub
%>


</html>
