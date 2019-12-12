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
Layout="News"
%>
<script language="javascript" src="images/Admin_Js.Js"></script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">信息管理</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2"><strong>操作选项</strong>：<a href="yqlj.asp">信息管理</a>&nbsp;|&nbsp;<a href="yqlj.asp?Action=Add">添加信息</a>&nbsp;&nbsp;</td>
  </tr>
  <tr class="t_rowh">
   
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
	Sql = "Select * From [yqlj]  Order By ID Desc"
Rs.Open Sql,Conn,1,1

%>
<form name="myform" method="post" action="yaopin.asp?Action=SDel">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr align="center" style="background-color:#009933">
      <td width="5%">选择</td>
	  <td width="6%">ID</td>
      <td width="20%">网站名称</td>
      <td width="35%">网站地址</td>
	
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
      <td align="left">&nbsp;<a href="?Action=Show&ID=<%=Rs("ID")%>"><%=rs("wzmc")%></a></td>
      <td><%=rs("wzdz")%></td>

     <td><%=rs("DateAndTime")%></td>
      <td>
        &nbsp;&nbsp;<a href="yqlj.asp?Action=Mod&ID=<%=Rs("ID")%>">修改</a>&nbsp;<a href="yqlj.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">删除</a></td>
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
<form name="myform" method="post" action="yqlj.asp?Action=SaveAdd" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">添加信息</td>
    </tr>
   
   <tr class="t_rowh">
      <td style="width:15%">　网站名称</td>
      <td style="width:80%"><input name="wzmc" type="text" id="wzmc" size="48">
        &nbsp;<font color="#FF0000">＊</font> ≤ 100个字符</td>
    </tr>
    <tr class="t_rowh">
      <td>　网站地址</td>
      <td><input name="wzdz" type="text" id="wzdz" size="48">
        &nbsp;<font color="#FF0000">＊</font> 记得添加<font color="#FF0000">http://</font></td>
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
	Dim ID,wzmc,wzdz
	Sql="Select * From [yqlj] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "信息不存在!","yqlj.asp",1
	Else
		ID=Rs("ID")
		wzmc=Rs("wzmc")
		wzdz=Rs("wzdz")

	
	End If
	Rs.Close
%>
<form name="myform" method="post" action="yqlj.asp?Action=SaveMod" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">修改信息</td>
    </tr>
    
    
       <tr class="t_rowh">
      <td style="width:15%">　网站名称</td>
      <td style="width:80%"><input name="ID" type="hidden" value="<%=ID%>"><input name="wzmc" type="text" id="wzmc" size="48" value="<%=wzmc%>">
        &nbsp;<font color="#FF0000">＊</font> ≤ 100个字符</td>
    </tr>
    <tr class="t_rowh">
      <td>　网站地址</td>
      <td><input name="wzdz" type="text" id="wzdz" size="48" value="<%=wzdz%>">
        &nbsp;<font color="#FF0000">＊</font> 记得添加<font color="#FF0000">http://</font></td>
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
	Dim ID,wzmc,wzdz,DateAndTime
	Sql="Select * From [yqlj] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "信息不存在!","yqlj.asp",1
	Else
		
		ID=Rs("ID")
		wzmc=Rs("wzmc")
		wzdz=Rs("wzdz")
		DateAndTime=Rs("DateAndTime")
	End If
	Rs.Close
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title"><%=wzmc%></td>
  </tr>
  <tr class="t_rowh">
    <td>&nbsp;Location:&nbsp;&nbsp;
      <%Call ShowClassPath(ClassID)%></td>
    <td align="right">
      </a>&nbsp;<a href="yqlj.asp?Action=Mod&ID=<%=ID%>">修改</a>&nbsp;<a href="yqlj.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">删除</a>&nbsp;<a href="<%=Http_Referer%>">返回</a>&nbsp;〗</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,wzmc,wzdz,DateAndTime,"yqlj"%></td>
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
	Dim ID,wzmc,wzdz
	
	wzmc=CheckChar(Trim(Request.Form("wzmc")),"网站名称",100,"012")
	wzdz=CheckChar(Trim(Request.Form("wzdz")),"网站地址",100,"012")

	If wzmc="" Then
	Response.Write("<script>alert(""请输入网站名称！"");history.back()</script>")
	elseif wzdz="" Then
	Response.Write("<script>alert(""请输入网站地址！"");history.back()</script>")
		else	
	Rs.Open "Select * From [yqlj]",conn,1,2
		Rs.Addnew
		ID=Rs("ID")
	    Rs("wzmc")=wzmc
		Rs("wzdz")=wzdz

		Rs("DateAndTime")=Now()

		Rs.Update
	Rs.Close
	Msg "信息添加成功!","yqlj.asp",0
		end if
End Sub

Sub SaveMod()
	Dim ID,wzmc,wzdz
	ID=CheckChar(Trim(Request.Form("ID")),"所属分类",0,"03")
	wzmc=CheckChar(Trim(Request.Form("wzmc")),"网站名称",100,"012")
	wzdz=CheckChar(Trim(Request.Form("wzdz")),"网站地址",100,"012")
	
	
	
	
	Rs.Open "Select * From [yqlj] Where ID="&ID,conn,1,2
	    Rs("wzmc")=wzmc
		Rs("wzdz")=wzdz
		
		Rs("DateAndTime")=Now()
		
		Rs.Update
	Rs.Close
	Msg "信息修改成功!","yqlj.asp",0
End Sub

Sub NewsDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"信息编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"yqlj.asp",1
	Rs.Open "Select * From [yqlj] Where ID="&ID,Conn,1,2
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
	If ErrContent<>"" Then Msg ErrContent,"yqlj.asp",1
	Rs.Open "Select * From [yqlj] Where ID In ("&ID&")",Conn,1,2
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
