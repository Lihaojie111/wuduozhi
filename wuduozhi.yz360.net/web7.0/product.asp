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
         <div class="xwgl">
		   
		   <div class="xw_nr">
<%
Layout="Product"
%>
<script language="javascript" src="images/Admin_Js.Js"></script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">形象管理</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2"><strong>操作选项</strong>：<a href="Product.asp">形象管理</a>&nbsp;|&nbsp;<a href="Product.asp?Action=Add">添加形象</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=Product">分类管理</a></td>
  </tr>
  <tr class="t_rowh">
    <form method="get" action="Product.asp">
    <td width="60%">快速搜索：&nbsp;<input name="Action" type="hidden" id="Action" value="Search"><input name="Word" type="text" id="Word" size="48"></td>
    </form>
	<td width="40%" align="right">查看分类：&nbsp;<select name="ClassID" id="ClassID" onchange="window.location='Product.asp?ClassID='+this.options[this.selectedIndex].value+'';"><option value="0">跳转分类至</option><%SelectClass(Cint(ClassID))%></select>&nbsp;&nbsp;&nbsp;</td>
  </tr>
</table>
<%
Select Case Action
	Case "Add"
		Call ProductAdd()
	Case "Mod"
		Call ProductMod()
	Case "Show"
		Call ProductShow()
	Case "Del"
		Call ProductDel()
	Case "SDel"
		Call ProductSDel()
	Case "Topis"
		Call ProductTopis()
	Case "SaveMod"
		Call SaveMod()
	Case "SaveAdd"
		Call SaveAdd()
	Case Else
		Call ProductList()
End Select

Sub ProductList()
	If Action="Search" Then
		Query_String="Action=Search&Word="&Request("Word")&"&"
		Sql="Where Title Like '%"&Request("Word")&"%'"
	ElseIf Cint(ClassID)>0 Then
		Dim RsTemp,ChildID
		Set RsTemp=Conn.Execute("Select ParentPath From Class Where ClassID="&ClassID)
		If Not RsTemp.Eof Then
			Dim ParentPath
			ParentPath = RsTemp("ParentPath")
		Else
			Msg "请选择要查询的分类","Product.asp",1
		End If
		Set RsTemp=Conn.Execute("Select ClassID From Class Where ParentID=" & ClassID & " Or ParentPath Like '"& ParentPath &","& ClassID &"%'")
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
	Sql = "Select * From Product "&Sql&" Order By ID Desc"
	Rs.Open Sql,Conn,1,1
%>
<form name="myform" method="post" action="Product.asp?Action=SDel">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr align="center" style="background-color:#009933">
    <td width="5%">选择</td>
	<td width="5%">ID</td>
    <td width="50%">标题</td>
    <td width="20%">分类</td>
    <td width="8%">点击次数</td>
    <td width="12%">操作</td>
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
	 <td><%=Rs("ID")%></td>
    <td align="left">&nbsp;<a href="?Action=Show&ID=<%=Rs("ID")%>"><%=Replace(Rs("Title"),Request("Word"),"<strong>"&Request("Word")&"</strong>")%></a></td>
    <td><%=ShowClassName(Rs("ClassID"))%></td>
    <td><%=(Rs("Hits"))%></td>
    <td><a href="Product.asp?Action=Topis&ID=<%=Rs("ID")%>"><%If Rs("Topis")=1 Then%><font color="#CC0000">荐</font><%Else%>荐<%End If%></a>&nbsp;<a href="Product.asp?Action=Mod&ID=<%=Rs("ID")%>">修改</a>&nbsp;<a href="Product.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">删除</a></td>
  </tr>
<%
		i = i + 1
		Rs.MoveNext
		Loop
		ShowPage(6)
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

Sub ProductAdd()
%>

<form name="myform" method="post" action="Product.asp?Action=SaveAdd" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">添加形象</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>所属分类</u></td>
    <td width="80%"><select name="ClassID" id="ClassID"><%SelectClass(0)%></select>&nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
  </tr>
  <tr class="t_row">
    <td><u>形象标题</u></td>
    <td><input name="Title" type="text" id="Title" size="48">&nbsp;<font color="#FF0000">＊</font> ≤ 100个字符&nbsp;&nbsp;&nbsp;&nbsp;<input name="Topis" type="checkbox" id="Topis" value="True"><label for="Topis">推荐形象</label></td>
  </tr>
  <tr class="t_row">
    <td><u>信息内容描述：</u></td>
    <td><input name="neirongmiaoshu" type="text" id="neirongmiaoshu" value="" size="48"></td>
  </tr>
  <tr class="t_row">
    <td><u>自主SEO优化--形象描述：</u></td>
    <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="描述可与标题相同.." size="48"></td>
  </tr>
  <tr class="t_row">
    <td><u>形象搜索关键字：</u></td>
    <td><input name="Keywords" type="text" id="Keywords" value="关键字可从标题中挑选即可.." size="48">
      <font color="#999999">（多个关键词请用“|”隔开）</font></td>
  </tr>
  <tr class="t_row">
    <td><u>上传形象</u></td>
    <td>		<p><input type="text" id="url1" value="" name="pic" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）</p></td>
  </tr>
  <tr class="t_row" >
    <td><u>链接地址</u></td>
    <td><input name="Price" type="text" id="Price" size="40" value="http://"></td>
  </tr>
    <tr class="t_row">
      <td colspan="2"><u>信息内容</u></td>
    </tr>
        <tr class="t_row">
      <td colspan="2">	<textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content1"> <%=Content%></textarea></td>
    </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="提交"></td>
  </tr>
</table>
</form>
<%
End Sub

Sub ProductMod()
	Dim ID,ClassID,Title,Price,Pic,Picture,Pictures,Topis,Hits,Content,TempPicture,neirongmiaoshu,DescriptionWord,Keywords
	Sql="Select * From Product Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "产品不存在!",Http_Referer,1
	Else
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		neirongmiaoshu=Rs("neirongmiaoshu")
		DescriptionWord=Rs("DescriptionWord")
		Keywords=Rs("Keywords")
		Title=Rs("Title")
		Price=Rs("Price")
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
// s_Thumbnail : 文件上传后，用于接收上传形象时所产生的缩略图文件的路径文件名的表单名，当未生成缩略图时，返回空值，原图用s_Link参数接收，此参数专用于缩略图
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//以下style=coolblue,值可以依据实际需要修改为您的样式名,通过此样式的后台设置来达到控制允许上传文件类型及文件大小
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=product&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>
<form name="myform" method="post" action="Product.asp?Action=SaveMod" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">修改形象</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>所属分类</u></td>
    <td width="80%"><input name="ID" type="hidden" value="<%=ID%>"><select name="ClassID" id="ClassID"><%SelectClass(ClassID)%></select>&nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
  </tr>
  <tr class="t_row">
    <td><u>形象标题</u></td>
    <td><input name="Title" type="text" id="Title" value="<%=Title%>" size="48">&nbsp;<font color="#FF0000">＊</font> ≤ 100个字符&nbsp;&nbsp;&nbsp;&nbsp;<input name="Topis" type="checkbox" id="Topis" value="True"<%If Topis=1 Then%> checked<%End If%>>&nbsp;<label for="Topis">推荐形象</label></td>
  </tr>
  neirongmiaoshu
  <tr class="t_row">
    <td><u>信息内容描述：</u></td>
    <td><input name="neirongmiaoshu" type="text" id="neirongmiaoshu" value="<%=neirongmiaoshu%>" size="48"></td>
  </tr>
  
  <tr class="t_row">
    <td><u>自主SEO优化--形象描述：</u></td>
    <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="<%=DescriptionWord%>" size="48">
（也可不填）</td>
  </tr>
  <tr class="t_row">
    <td><u>形象搜索关键字：</u></td>
    <td><input name="Keywords" type="text" id="Keywords" value="<%=Keywords%>" size="48">
（也可不填）</td>
  </tr>
 <tr class="t_row">
    <td><u>默认形象</u></td>
    <td>		<p><input type="text" id="url1" value="<%=pic%>" name="pic" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）</p></td>
  </tr>
  
  <tr class="t_row" style="display:none;">
    <td><u>链接地址：</u></td>
    <td><input name="Price" type="text" id="Price" value="<%=Price%>" size="40" ></td>
  </tr>
    <tr class="t_row">
      <td colspan="2"><u>信息内容</u></td>
    </tr>
        <tr class="t_row">
      <td colspan="2">	<textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content1"> <%=Content%></textarea></td>
    </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="提交"></td>
  </tr>
</table>
</form>
<%
End Sub

Sub ProductShow()
	Dim ID,ClassID,Title,Price,Pic,Picture,Pictures,Topis,Hits,Content,DateAndTime,neirongmiaoshu,DescriptionWord,Keywords
	Sql="Select * From Product Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "产品不存在!",Http_Referer,1
	Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		Title=Rs("Title")
		Price=Rs("Price")
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
    <td align="right">〖&nbsp;操作：<a href="Product.asp?Action=Topis&ID=<%=ID%>"><%If Topis=1 Then%><font color="#CC0000">荐</font><%Else%>荐<%End If%></a>&nbsp;<a href="Product.asp?Action=Mod&ID=<%=ID%>">修改</a>&nbsp;<a href="Product.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">删除</a>&nbsp;<a href="<%=Http_Referer%>">返回</a>&nbsp;〗</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,Title,ClassID,Hits,DateAndTime,"<div align=""center""><img src="""&Pic&"""></div><br>"&Replace(Replace(Content,Chr(10),"<br>"),Chr(13),"&nbsp;")&"<br><br><strong>说明书下载</strong>:"&Picture,"Product"%></td>
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
	Dim ID,ClassID,Title,Price,Pic,Picture,Topis,Content,neirongmiaoshu,DescriptionWord,Keywords
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"产品名称",100,"012")
	neirongmiaoshu=CheckChar(Trim(Request.Form("neirongmiaoshu")),"信息内容描述",2000,"12")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"产品品牌",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"产品型号",2000,"012")
	Price=CheckChar(Trim(Request.Form("Price")),"产品价格",50,"2")
	Picture=CheckChar(Trim(Request.Form("Picture")),"说明书",100,"12")
	Pic=CheckChar(Trim(Request.Form("Pic")),"产品",255,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"推荐产品",0,"")
	Content=CheckChar(Request.Form("Content1"),"产品介绍",65535,"02")
	If ClassID<=0 Then ErrContent = ErrContent & "请先添加分类!"
	If ErrContent<>"" Then Msg ErrContent,"Product.asp?Action=Add",1
	Rs.Open "Select * From Product",conn,1,2
		Rs.Addnew
		ID=Rs("ID")
		Rs("Title")=Title
		Rs("Content")=Content
		Rs("ClassID")=ClassID
		Rs("Price")=Price
		Rs("Pic")=Pic
		Rs("DescriptionWord")=DescriptionWord
		Rs("neirongmiaoshu")=neirongmiaoshu
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
	Msg "产品添加成功!","Product.asp",0
End Sub

Sub SaveMod()
	Dim ID,ClassID,Title,Price,Pic,Picture,Topis,Content,neirongmiaoshu,DescriptionWord,Keywords
	ID=CheckChar(Trim(Request.Form("ID")),"产品编号",0,"03")
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"产品名称",100,"012")
	Price=CheckChar(Trim(Request.Form("Price")),"产品价格",50,"2")
	Picture=CheckChar(Trim(Request.Form("Picture")),"说明书",100,"12")
	neirongmiaoshu=CheckChar(Trim(Request.Form("neirongmiaoshu")),"信息内容描述",2000,"12")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"产品品牌",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"产品型号",2000,"012")
	Pic=CheckChar(Trim(Request.Form("Pic")),"产品",255,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"推荐产品",0,"")
	Content=CheckChar(Request.Form("Content1"),"产品介绍",65535,"02")
	If ErrContent<>"" Then Msg ErrContent,"Product.asp?Action=Mod&ID="&ID,1
	Rs.Open "Select * From Product Where ID="&ID,conn,1,2
		Rs("Title")=Title
		Rs("Content")=Content
		Rs("ClassID")=ClassID
		Rs("Price")=Price
		Rs("neirongmiaoshu")=neirongmiaoshu
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
	Msg "产品修改成功!","Product.asp",0
End Sub

Sub ProductDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"产品编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Product.asp",1
	Rs.Open "Select * From Product Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "产品不存在,不可以执行删除操作!",Http_Referer,1
	Else
'		DelFiles(Rs("Picture")&"")
		Rs.Delete
	End If
	Rs.Close
	Msg "产品ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub

Sub ProductSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"产品编号",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"Product.asp",1
	Rs.Open "Select * From Product Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "产品不存在,不可以执行删除操作!",Http_Referer,1
	Else
		Do While Not Rs.Eof
'		DelFiles(Rs("Picture")&"")
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "产品ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub

Sub ProductTopis()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"产品编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Product.asp",1
	Rs.Open "Select ID,Topis From Product Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "产品不存在,不可以执行操作!",Http_Referer,1
	Else
		If Rs("Topis")=1 Then
			Rs("Topis")=0
		Else
			Rs("Topis")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "产品ID:"&ID&"置顶设置成功!",Http_Referer,0
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
