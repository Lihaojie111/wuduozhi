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
<!--ͷ��-->
<div class="wramp">
  <!--#include file="head.asp"-->
  <!--���-->

  <div class="clear"></div>
  
  
  <div id="content">
  
 <!--#include file="left.asp"-->
	
	<div class="c-right">
<div class="c-rgt01">
		<div style="float:left; padding-left:20px; padding-top:10px">�����ڵ�λ�ã���̨��̨-��ҳSEO�Ż�����</div>
		
		<div style="float:right; padding-right:100px;"> <form action="http://www.baidu.com/baidu" target="_blank">
<table bgcolor="#f1f1f1"><tr><td>
<input name=tn type=hidden value=baidu>
<img src="http://img.baidu.com/img/logo-80px.gif" alt="Baidu" align="bottom" border="0" />
<input type=text name=word size=30>
<input type="submit" value="�ٶ�����">
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
    <td width="10%">ѡ��</td>
    <td  width="10%">���</td>
    <td>����</td>
    <td width="10%">�������</td>
    <td width="40%">���ʱ��</td>
    <td width="20%">����</td>
  </tr>
<%
	Sql ="Select * From [about] Order By ID Desc"
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Response.Write("<tr align=""center"" class=""t_rowh""><td colspan=""5""><strong>û���κμ�¼</strong></td></tr>")
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
    <td><a href="about.asp?Action=Mod&ID=<%=Rs("ID")%>">�޸�</a>&nbsp;<a href="about.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">ɾ��</a></td>
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
   if(confirm("ɾ�����ݺ��ָܻ���ȷ��Ҫɾ����"))
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
// ����˵��
// s_Type : �ļ����ͣ�����ֵΪ"image","flash","media","file"
// s_Link : �ļ��ϴ������ڽ����ϴ��ļ�·���ļ����ı���
// s_Thumbnail : �ļ��ϴ������ڽ����ϴ�ͼƬʱ������������ͼ�ļ���·���ļ����ı�������δ��������ͼʱ�����ؿ�ֵ��ԭͼ��s_Link�������գ��˲���ר��������ͼ
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//����style=coolblue,ֵ��������ʵ����Ҫ�޸�Ϊ������ʽ��,ͨ������ʽ�ĺ�̨�������ﵽ���������ϴ��ļ����ͼ��ļ���С
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=about&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>

<form name="myform" method="post" action="about.asp?Action=SaveAdd" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">��ҳ����SEO�Ż�����</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>��ҳ���⣺</u></td>
    <td width="80%"><input name="Title" type="text" id="Title" size="48">&nbsp;&nbsp;<font color="#FF5353">��</font>&nbsp;<font color="#999999">��100���ַ�������ؼ������á�|��������</font></td>
  </tr>
    <tr class="t_row">
    <td width="20%"><u>��ҳ�ؼ��ʣ�</u></td>
    <td width="80%"><input name="Keywords" type="text" id="Keywords" size="48">
    &nbsp;&nbsp;<font color="#FF5353">��</font>&nbsp;<font color="#999999">��100���ַ�������ؼ������á�|��������</font></td>
  </tr>
      <tr class="t_row">
    <td width="20%"><u>��ҳ������</u></td>
    <td width="80%"><input name="DescriptionWord" type="text" id="DescriptionWord" size="48">&nbsp;&nbsp;<font color="#FF5353">��</font>&nbsp;<font color="#999999">��100���ַ�������ؼ������á�|��������</font></td>
  </tr>
  
<tr class="t_row">
    <td><u>Ĭ��ͼƬ</u></td>
    <td><input type="text" id="url1" value="" name="pic" /> <input type="button" id="image1" value="ѡ��ͼƬ" />������ͼƬ + �����ϴ���</p></td>
  </tr>
<tr class="t_row">
    <td><u>��ҳ����</u><li>��ϵͳ�����ͼƬ���Ƶ���վ�������ϣ�ϵͳ������ͼƬ�Ĵ�С��Ӱ���ٶȡ�</li><li>�����밴Shift+Enter������һ���밴Enter</li></td>
    <td><textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content1"></textarea></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="�ύ"></td>
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
		Msg "��ҳ������!",Http_Referer,1
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
// ����˵��
// s_Type : �ļ����ͣ�����ֵΪ"image","flash","media","file"
// s_Link : �ļ��ϴ������ڽ����ϴ��ļ�·���ļ����ı���
// s_Thumbnail : �ļ��ϴ������ڽ����ϴ�ͼƬʱ������������ͼ�ļ���·���ļ����ı�������δ��������ͼʱ�����ؿ�ֵ��ԭͼ��s_Link�������գ��˲���ר��������ͼ
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//����style=coolblue,ֵ��������ʵ����Ҫ�޸�Ϊ������ʽ��,ͨ������ʽ�ĺ�̨�������ﵽ���������ϴ��ļ����ͼ��ļ���С
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=about&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>

<form name="myform" method="post" action="about.asp?Action=SaveMod" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">�޸ĵ�ҳ����SEO�Ż�����</td>
  </tr>
  <tr class="t_rowh">
    <td width="20%"><u>��ҳ����:</u></td>
    <td width="80%"><input name="ID" type="hidden" value="<%=ID%>"><input name="Title" type="text" id="Title" value="<%=Title%>" size="48">
      &nbsp;<font color="#FF5353">��</font>&nbsp;<font color="#999999">��100���ַ�������ؼ������á�|��������</font></td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>��ҳ�ؼ��ʣ�</u></td>
    <td width="80%"><input name="Keywords" type="text" id="Keywords" value="<%=Keywords%>" size="48">
    &nbsp;&nbsp;<font color="#FF5353">��</font>&nbsp;<font color="#999999">��100���ַ�������ؼ������á�|��������</font></td>
  </tr>
      <tr class="t_row">
    <td width="20%"><u>��ҳ������</u></td>
    <td width="80%"><input name="DescriptionWord" type="text" id="DescriptionWord" value="<%=DescriptionWord%>" size="48">&nbsp;&nbsp;<font color="#FF5353">��</font>&nbsp;<font color="#999999">��100���ַ�������ؼ������á�|��������</font></td>
  </tr>
  
 <tr class="t_row">
    <td><u>Ĭ��ͼƬ</u></td>
    <td><input type="text" id="url1" value="<%=pic%>" name="pic" /> <input type="button" id="image1" value="ѡ��ͼƬ" />������ͼƬ + �����ϴ���</p></td>
  </tr>
<tr class="t_row">
    <td><u>��ҳ����</u><li>��ϵͳ�����ͼƬ���Ƶ���վ�������ϣ�ϵͳ������ͼƬ�Ĵ�С��Ӱ���ٶȡ�</li><li>�����밴Shift+Enter������һ���밴Enter</li></td>
    <td><textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content1"> <%=Content%></textarea></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="�ύ"></td>
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
		Msg "��ҳ������!",Http_Referer,1
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
    <td align="right">��&nbsp;������<a href="about.asp?Action=Mod&ID=<%=ID%>">�޸�</a>&nbsp;<a href="about.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">ɾ��</a>&nbsp;<a href="<%=Http_Referer%>">����</a>&nbsp;��</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,Title,0,Hits,DateAndTime,Content,"[about]"%></td>
  </tr>
</table>
<Script language="JavaScript" type="text/JavaScript">
function Delete()
{
   if(confirm("ɾ�����ݺ��ָܻ���ȷ��Ҫɾ����"))
     return true;
   else
     return false;
}
</Script>
<%
End Sub
Sub SaveAdd()
	Dim Title,pic,Content,Keywords,DescriptionWord
	Title=CheckChar(Trim(Request.Form("Title")),"����",200,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"����",200,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"����",200,"012")
	pic=CheckChar(Trim(Request.Form("pic")),"ͼƬ",225,"2")
	Content=CheckChar(Trim(Request.Form("Content1")),"��ҳ����",10000,"2")
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
	Msg "��ҳ��ӳɹ�!","about.asp",0
End Sub
Sub SaveMod()
	Dim ID,Title,pic,Content,Keywords,DescriptionWord
	ID=CheckChar(Trim(Request.Form("ID")),"��ҳ���",200,"012")
	Title=CheckChar(Trim(Request.Form("Title")),"����",200,"012")
	pic=CheckChar(Trim(Request.Form("pic")),"����",225,"2")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"����",200,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"����",200,"012")
	Content=CheckChar(Trim(Request.Form("Content1")),"��ҳ����",10000,"2")
	If ErrContent<>"" then Msg ErrContent,Http_Referer,1
	Rs.Open "Select * From [about] Where ID="&ID,Conn,1,2
	        Rs("Title")=Title
			Rs("pic")=pic
			Rs("Keywords")=Keywords
			Rs("DescriptionWord")=DescriptionWord
			Rs("Content")=Content
		Rs.Update
	Rs.Close
	Msg "��ҳ�޸ĳɹ�!","about.asp",0
End Sub
Sub SaveDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��ҳ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"about.asp",1
	Rs.Open "Select * From [about] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��ҳ������,������ִ��ɾ������!",Http_Referer,1
	Else
		Rs.Delete
	End If
	Rs.Close
	Msg "��ҳID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub
Sub SaveSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"��ҳ���",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"about.asp",1
	Rs.Open "Select * From [about] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "��ҳ������,������ִ��ɾ������!",Http_Referer,1
	Else
		Do While Not Rs.Eof
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "��ҳID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
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
