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
		 var editor1 = K.create('#Content1', {             
		 cssPath: '../kingeditor/plugins/code/prettify.css',             
		 uploadJson: '../kindeditor/asp/upload_json.asp',            
		  fileManagerJson: '../kindeditor/asp/file_manager_json.asp',            
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
				
				K('#insertfile').click(function() {
					editor.loadPlugin('insertfile', function() {
						editor.plugin.fileDialog({
							fileUrl : K('#url').val(),
							clickFn : function(url, title) {
								K('#url').val(url);
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
Layout="zzry"
%>
<script language="javascript" src="images/Admin_Js.Js"></script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">信息管理</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2"><strong>操作选项</strong>：<a href="zzry.asp">信息管理</a>&nbsp;|&nbsp;<a href="zzry.asp?Action=Add">添加信息</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=zzry">分类管理</a></td>
  </tr>
  <tr class="t_rowh">
<form method="get" action="zzry.asp">
      <td width="60%">快速搜索：&nbsp;
        <input name="Action" type="hidden" id="Action" value="Search">
        <input name="Word" type="text" id="Word" size="48"></td>
		<td><input  type="submit"  value="搜索" /></td>
    </form>
    <td width="40%" align="right">查看分类：&nbsp;
      <select name="ClassID" id="ClassID" onchange="window.location='zzry.asp?ClassID='+this.options[this.selectedIndex].value+'';">
        <option value="0">跳转分类至</option>
        <%SelectClass(Cint(ClassID))%>
      </select>
      &nbsp;&nbsp;&nbsp;</td>
  </tr>
</table>
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
		Call NewsList()
End Select

Sub NewsList()
	If Action="Search" Then
		Sql="Where Title Like '%"&Request("Word")&"%'"
		Query_String="Word="&Request("Word")&"&"
	ElseIf Cint(ClassID)>0 Then
		Dim RsTemp,ChildID
		Set RsTemp=Conn.Execute("Select ParentPath From [Class] Where ClassID="&ClassID)
		If Not RsTemp.Eof Then
			Dim ParentPath
			ParentPath = RsTemp("ParentPath")
		Else
			Msg "请选择要查询的分类","zzry.asp",1
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

	Rs.Open "Select * From [zzry] "&Sql&" Order By ID Desc",Conn,1,1

%>
<form name="myform" method="post" action="zzry.asp?Action=SDel">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr align="center" style="background-color:#009933">
      <td width="5%">选择</td>
	  <td width="6%">ID</td>
      <td width="41%">荣誉名称</td>	
	  <td width="14%">分类</td>
      <td width="19%">添加时间</td>
      <td width="15%">操作</td>
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
	   <td><%=rs("id")%></td>
      <td align="center"><%=rs("title")%></td>
	  <td><%=ShowClassName(Rs("ClassID"))%></td>
     <td><%=rs("DateAndTime")%></td>
      <td>
        &nbsp;&nbsp;<a href="zzry.asp?Action=Mod&ID=<%=Rs("ID")%>">修改</a>&nbsp;<a href="zzry.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">删除</a></td>
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

Sub NewsAdd()
%>
<form name="myform" method="post" action="zzry.asp?Action=SaveAdd" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">添加信息</td>
    </tr>
	 <tr class="t_rowh">
      <td width="15%">所属分类</td>
      <td width="80%"><select name="ClassID" id="ClassID">
          <%SelectClass(0)%>
        </select>
        &nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
    </tr>
   
   <tr class="t_rowh">
      <td style="width:15%">荣誉名称</td>
      <td style="width:80%"><input name="title" type="text" id="title" size="48">
        &nbsp;<font color="#FF0000">＊</font> ≤ 100个字符</td>
    </tr>	
	 <tr class="t_row">
    <td><u>图片上传</u></td>
    <td>		<p><input type="text" id="url1" value="" name="pic" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）</p></td>
  </tr>
    
    <tr class="t_row">
      <td>&nbsp;</td>
      <td><input type="submit" value="提交"></td>
    </tr>
  </table>
</form>
<%
End Sub

Sub NewsMod()
	Dim ID,ClassID,title,pic
	Sql="Select * From [zzry] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "信息不存在!","zzry.asp",1
	Else
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		title=Rs("title")
		pic=Rs("pic")



	
	End If
	Rs.Close
%>
<form name="myform" method="post" action="zzry.asp?Action=SaveMod" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">修改信息</td>
    </tr>
     <tr class="t_rowh">
      <td width="15%"><u>所属分类</u></td>
      <td width="80%"><input name="ID" type="hidden" value="<%=ID%>">
        <select name="ClassID" id="ClassID">
          <%SelectClass(ClassID)%>
        </select>
        &nbsp;<font color="#FF0000">＊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>注意：不能指定为含有子分类的分类</strong></font></td>
    </tr>
    
       <tr class="t_rowh">
      <td style="width:15%">荣誉名称</td>
      <td style="width:80%"><input name="title" type="text" id="title" size="48" value="<%=title%>">
        &nbsp;<font color="#FF0000">＊</font> ≤ 100个字符</td>
    </tr>
	 <tr class="t_row">
    <td><u>图片上传</u></td>
    <td>		<p><input type="text" id="url1" value="<%=pic%>" name="pic" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）</p></td>
  </tr>
    
    
    <tr class="t_row">
      <td>&nbsp;</td>
      <td><input type="submit" value="提交"></td>
    </tr>
  </table>
</form>
<%
End Sub

Sub NewsShow()
	Dim ID,ClassID,title,pic,DateAndTime
	Sql="Select * From [zzry] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在!","zzry.asp",1
	Else
		
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		title=Rs("title")
		pic=Rs("pic")
		DateAndTime=Rs("DateAndTime")
	End If
	Rs.Close
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title"><%=title%></td>
  </tr>
  <tr class="t_rowh">
    <td>&nbsp;Location:&nbsp;&nbsp;
      <%Call ShowClassPath(ClassID)%></td>
    <td align="right">
      </a>&nbsp;<a href="zzry.asp?Action=Mod&ID=<%=ID%>">修改</a>&nbsp;<a href="zzry.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">删除</a>&nbsp;<a href="<%=Http_Referer%>">返回</a>&nbsp;〗</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,title,pic,DateAndTime,"zzry"%></td>
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
	Dim ID,ClassID,title,czxt,rjjs,pic,spdz
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	title=CheckChar(Trim(Request.Form("title")),"荣誉名称",100,"012")
	Pic=CheckChar(Trim(Request.Form("Pic")),"图片",255,"12")
If ClassID<=0 Then ErrContent = ErrContent & "请先添加分类!"
	If ErrContent<>"" Then Msg ErrContent,"zzry.asp?Action=Add",1
	If Conn.Execute("Select Child From Class Where ClassID="&ClassID)(0)>0 Then
		Msg "所属分类含有子目录,请选择其他分类!","zzry.asp?Action=Add",1
	End If
	Rs.Open "Select * From [zzry]",conn,1,2
		Rs.Addnew
		ID=Rs("ID")
		Rs("ClassID")=ClassID
		Rs("title")=title
		Rs("pic")=pic
		Rs("DateAndTime")=Now()

		Rs.Update
	Rs.Close
	Msg "信息添加成功!","zzry.asp",0
		
End Sub

Sub SaveMod()
	Dim ID,ClassID,title,czxt,rjjs,pic,spdz
	ID=CheckChar(Trim(Request.Form("ID")),"信息编号",0,"03")
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"所属分类",0,"03")
	title=CheckChar(Trim(Request.Form("title")),"荣誉名称",100,"012")
	Pic=CheckChar(Trim(Request.Form("Pic")),"图片",255,"12")
	If ErrContent<>"" Then Msg ErrContent,"zzry.asp?Action=Mod&ID="&ID,1
	If Conn.Execute("Select Child From Class Where ClassID="&ClassID)(0)>0 Then
		Msg "所属分类含有子目录,请选择其他分类!","zzry.asp?Action=Mod&ID="&ID,1
	End If
	
	Rs.Open "Select * From [zzry] Where ID="&ID,conn,1,2
	    Rs("title")=title
		Rs("ClassID")=ClassID
		Rs("pic")=pic
		Rs("DateAndTime")=Now()
		
		Rs.Update
	Rs.Close
	Msg "信息修改成功!","zzry.asp",0
End Sub

Sub NewsDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"信息编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"zzry.asp",1
	Rs.Open "Select * From [zzry] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行删除操作!",Http_Referer,1
	Else
		
		Rs.Delete
	End If
	Rs.Close
	Msg "信息ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
End Sub


Sub NewsSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"信息编号",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"zzry.asp",1
	Rs.Open "Select * From [zzry] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在,不可以执行删除操作!",Http_Referer,1
	Else
		Do While Not Rs.Eof
		
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "信息ID:"&ID&"已删除,不可以恢复!",Http_Referer,0
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
