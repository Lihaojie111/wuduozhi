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




   <script charset="utf-8" src="../kindeditor/kindeditor.js"></script>
<script charset="utf-8" src="../kindeditor/lang/zh_CN.js"></script>
<script charset="utf-8" src="../kindeditor/plugins/code/prettify.js"></script>
<script>     KindEditor.ready(function (K) {
		 var editor1 = K.create('#content1', {             
		 cssPath: '../kindeditor/plugins/code/prettify.css',             
		 uploadJson: '../kindeditor/asp/upload_json.asp',            
		  fileManagerJson: '../kindeditor/aspt/file_manager_json.asp',            
		   allowFileManager: true,             
		   afterCreate: function () {                 
		   var self = this;                 
		   K.ctrl(document, 13, function () {                    
			self.sync();                     
			K('form[name=example]')[0].submit();                 
			});                 
			K.ctrl(self.edit.doc, 13, function () {                    
			 self.sync();                    
			  K('form[name=example]')[0].submit();                
			   });             
			   }         
			   });        
				prettyPrint();     
				}); 
				
				
				KindEditor.ready(function(K) {
				var editor = K.editor({
					allowFileManager : true
				});
				
						K('#image1').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							imageUrl : K('#url1').val(),
							clickFn : function(url, title, width, height, border, align) {
								K('#url1').val(url);
								editor.hideDialog();
							}
						});
					});
				});	
				});
          </script>



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
<div class="c-rgt01">
		<div style="float:left; padding-left:20px; padding-top:10px">您现在的位置：后台首台-单页SEO优化管理</div>
		
		<div style="float:right; padding-right:100px;"> <form action="http://www.baidu.com/baidu" target="_blank">
<table bgcolor="#f1f1f1"><tr><td>
<input name=tn type=hidden value=baidu>
<img src="http://img.baidu.com/img/logo-80px.gif" alt="Baidu" align="bottom" border="0" />
<input type=text name=word size=30>
<input type="submit" value="百度搜索">
</td></tr></table>
</form></div>

		</div>
       <%
Select Case Action
	Case "Add"
		Call aboutAdd()
	Case "Mod"
		Call aboutMod()
	Case "Show"
		Call aboutShow()
	Case "SaveAdd"
		Call SaveAdd()
	Case "SaveMod"
		Call SaveMod()
	Case "Del"
		Call SaveDel()
	Case "SDel"
		Call SaveSDel()
	Case Else
		Call aboutlist()
End Select

Sub aboutlist()
%>
<form name="myform" method="post" action="about.asp?Action=SDel">
<table border="0" align="center" width="100%" cellpadding="0" cellspacing="1" class="t_border">
  <tr align="center" style="background-color:#009933">
    <td width="10%">选择</td>
    <td  width="10%">编号</td>
    <td>标题</td>
    <td width="10%">点击次数</td>
    <td width="40%">添加时间</td>
    <td width="20%">操作</td>
  </tr>
<%
	Sql ="Select * From [about] Order By ID Desc"
	Rs.Open Sql,Conn,1,1
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
    <td ><%=Rs("ID")%></td>
    <td align="center"><a href="?Action=Show&ID=<%=Rs("ID")%>"><%=Rs("Title")%></a></td>
    <td><%=Rs("Hits")%></td>
    <td><%=Rs("DateAndTime")%></td>
    <td><a href="about.asp?Action=Mod&ID=<%=Rs("ID")%>">修改</a>&nbsp;<a href="about.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">删除</a></td>
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

Sub aboutAdd()
%>
<script language="javascript">
// 参数说明
// s_Type : 文件类型，可用值为"image","flash","media","file"
// s_Link : 文件上传后，用于接收上传文件路径文件名的表单名
// s_Thumbnail : 文件上传后，用于接收上传图片时所产生的缩略图文件的路径文件名的表单名，当未生成缩略图时，返回空值，原图用s_Link参数接收，此参数专用于缩略图
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//以下style=coolblue,值可以依据实际需要修改为您的样式名,通过此样式的后台设置来达到控制允许上传文件类型及文件大小
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=about&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>

<form name="myform" method="post" action="about.asp?Action=SaveAdd" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">单页自主SEO优化管理</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>单页标题：</u></td>
    <td width="80%"><input name="Title" type="text" id="Title" size="48">&nbsp;&nbsp;<font color="#FF5353">＊</font>&nbsp;<font color="#999999">≤100个字符（多个关键词请用“|”隔开）</font></td>
  </tr>
    <tr class="t_row">
    <td width="20%"><u>单页关键词：</u></td>
    <td width="80%"><input name="Keywords" type="text" id="Keywords" size="48">
    &nbsp;&nbsp;<font color="#FF5353">＊</font>&nbsp;<font color="#999999">≤100个字符（多个关键词请用“|”隔开）</font></td>
  </tr>
      <tr class="t_row">
    <td width="20%"><u>单页描述：</u></td>
    <td width="80%"><input name="DescriptionWord" type="text" id="DescriptionWord" size="48">&nbsp;&nbsp;<font color="#FF5353">＊</font>&nbsp;<font color="#999999">≤100个字符（多个关键词请用“|”隔开）</font></td>
  </tr>
  
<tr class="t_row">
    <td><u>默认图片</u></td>
    <td><input type="text" id="url1" value="" name="pic" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）</p></td>
  </tr>
<tr class="t_row">
    <td><u>单页内容</u><li>本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度。</li><li>换行请按Shift+Enter，另起一段请按Enter</li></td>
    <td><textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content1"></textarea></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="提交"></td>
  </tr>
</table>
</form>
<%
End Sub
Sub aboutMod()
	Dim ID,Title,Keywords,Content,pic,DescriptionWord
	Sql="Select * From [about] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "单页不存在!",Http_Referer,1
	Else
		ID=Rs("ID")
		Title=Rs("Title")
		pic=Rs("pic")
		Keywords=Rs("Keywords")
		DescriptionWord=Rs("DescriptionWord")
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
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=about&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>

<form name="myform" method="post" action="about.asp?Action=SaveMod" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">修改单页自主SEO优化管理</td>
  </tr>
  <tr class="t_rowh">
    <td width="20%"><u>单页标题:</u></td>
    <td width="80%"><input name="ID" type="hidden" value="<%=ID%>"><input name="Title" type="text" id="Title" value="<%=Title%>" size="48">
      &nbsp;<font color="#FF5353">＊</font>&nbsp;<font color="#999999">≤100个字符（多个关键词请用“|”隔开）</font></td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>单页关键词：</u></td>
    <td width="80%"><input name="Keywords" type="text" id="Keywords" value="<%=Keywords%>" size="48">
    &nbsp;&nbsp;<font color="#FF5353">＊</font>&nbsp;<font color="#999999">≤100个字符（多个关键词请用“|”隔开）</font></td>
  </tr>
      <tr class="t_row">
    <td width="20%"><u>单页描述：</u></td>
    <td width="80%"><input name="DescriptionWord" type="text" id="DescriptionWord" value="<%=DescriptionWord%>" size="48">&nbsp;&nbsp;<font color="#FF5353">＊</font>&nbsp;<font color="#999999">≤100个字符（多个关键词请用“|”隔开）</font></td>
  </tr>
  
 <tr class="t_row">
    <td><u>默认图片</u></td>
    <td><input type="text" id="url1" value="<%=pic%>" name="pic" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）</p></td>
  </tr>
<tr class="t_row">
    <td><u>单页内容</u><li>本系统将会把图片复制到本站服务器上，系统会因复制图片的大小而影响速度。</li><li>换行请按Shift+Enter，另起一段请按Enter</li></td>
    <td><textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content1"> <%=Content%></textarea></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="提交"></td>
  </tr>
</table>
</form>
<%
End Sub
Sub aboutShow()
	Dim ID,Title,pic,Keywords,DescriptionWord,Content,Hits,DateAndTime
	Sql="Select * From [about] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "单页不存在!",Http_Referer,1
	Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ID=Rs("ID")
		Title=Rs("Title")
		pic=Rs("pic")
		Keywords=Rs("Keywords")
		DescriptionWord=Rs("DescriptionWord")
		Content=Rs("Content")
		Hits=Rs("Hits")
		DateAndTime=Rs("DateAndTime")
	End If
	Rs.Close
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title"><%=Title%></td>
  </tr>
  <tr class="t_rowh">
    <td>&nbsp;</td>
    <td align="right">〖&nbsp;操作：<a href="about.asp?Action=Mod&ID=<%=ID%>">修改</a>&nbsp;<a href="about.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">删除</a>&nbsp;<a href="<%=Http_Referer%>">返回</a>&nbsp;〗</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,Title,0,Hits,DateAndTime,Content,"[about]"%></td>
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
	Dim Title,pic,Content,Keywords,DescriptionWord
	Title=CheckChar(Trim(Request.Form("Title")),"标题",200,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"标题",200,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"标题",200,"012")
	pic=CheckChar(Trim(Request.Form("pic")),"图片",225,"2")
	Content=CheckChar(Trim(Request.Form("Content1")),"单页内容",10000,"2")
	If ErrContent<>"" then Msg ErrContent,Http_Referer,1
	Rs.Open "Select * From [about]",Conn,1,2
		Rs.AddNew
			Rs("Title")=Title
			Rs("pic")=pic
			Rs("Keywords")=Keywords
			Rs("DescriptionWord")=DescriptionWord
			Rs("Content")=Content
			Rs("Hits")=0
			Rs("DateAndTime")=Now()
		Rs.Update
	Rs.Close
	Msg "单页添加成功!","about.asp",0
End Sub
Sub SaveMod()
	Dim ID,Title,pic,Content,Keywords,DescriptionWord
	ID=CheckChar(Trim(Request.Form("ID")),"单页编号",200,"012")
	Title=CheckChar(Trim(Request.Form("Title")),"标题",200,"012")
	pic=CheckChar(Trim(Request.Form("pic")),"标题",225,"2")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"标题",200,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"标题",200,"012")
	Content=CheckChar(Trim(Request.Form("Content1")),"单页内容",10000,"2")
	If ErrContent<>"" then Msg ErrContent,Http_Referer,1
	Rs.Open "Select * From [about] Where ID="&ID,Conn,1,2
	        Rs("Title")=Title
			Rs("pic")=pic
			Rs("Keywords")=Keywords
			Rs("DescriptionWord")=DescriptionWord
			Rs("Content")=Content
		Rs.Update
	Rs.Close
	Msg "单页修改成功!","about.asp",0
End Sub
Sub SaveDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"单页编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"about.asp",1
	Rs.Open "Select * From [about] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "单页不存在,不可以执行删除操作!",Http_Referer,1
	Else
		Rs.Delete
	End If
	Rs.Close
	Msg "单页ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub
Sub SaveSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"单页编号",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"about.asp",1
	Rs.Open "Select * From [about] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "单页不存在,不可以执行删除操作!",Http_Referer,1
	Else
		Do While Not Rs.Eof
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "单页ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
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
