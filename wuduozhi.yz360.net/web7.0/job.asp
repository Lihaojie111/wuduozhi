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

<link type="text/css" rel="stylesheet" href="css.css" />
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
	<div class="c-rgt01" style="margin-bottom:10px;">
		<div style="float:left; padding-left:20px; padding-top:10px">�����ڵ�λ�ã���̨��̨-�޸�����</div>
		
		<div style="float:right; padding-right:100px;"> <form action="http://www.baidu.com/baidu" target="_blank">
<table bgcolor="#f1f1f1"><tr><td>
<input name=tn type=hidden value=baidu>
<img src="http://img.baidu.com/img/logo-80px.gif" alt="Baidu" align="bottom" border="0" />
<input type=text name=word size=30>
<input type="submit" value="�ٶ�����">
</td></tr></table>
</form></div>

		</div>
<table border="0" align="center"  cellpadding="0" cellspacing="1" style="padding-left:50px;">
  <tr>
    <td class="t_title">&nbsp;</td>
  </tr>
  <tr>
    <td ><strong>����ѡ��</strong>��&nbsp;&nbsp;<a href="job.asp?Action=Add">����ְλ</a>&nbsp;|&nbsp;<a href="job.asp">ְλ���� </a></td>
  </tr>
 </table>
 <br />
<%
Select Case Action
    Case "chakan"
	    Call chakan()
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
	Case"CKDel"
		Call CKDel()
	Case Else
		Call aboutlist()
End Select
Sub aboutlist()
%>


<form name="myform" method="post" action="job.asp?Action=SDel">
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr align="center" bgcolor="#339933">
    <td width="9%">ѡ��</td>
    <td width="15%">��Ƹְλ</td>
	<td width="15%">��������</td>
    <td width="15%">���ʱ��</td>
    <td width="15%">����</td>
  </tr>
<%
	Sql ="Select * From [join]  Order By ID Desc"
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Response.Write("<tr align=""center"" class=""t_rowh""><td colspan=""5""><strong>û���κμ�¼!</strong></td></tr>")
	Else
		i=0
		Rs.PageSize = Cint(Line)
		Rs.AbsolutePage = Cint(Page)
		Do While Not Rs.Eof And i < Rs.PageSize
%>
  <tr align="center" class="t_row" onmouseover="this.className='t_rowh'" onmouseout="this.className='t_row'">
    <td><input name="ID" type="checkbox" id="ID" value="<%=Rs("ID")%>"></td>
    <td align="center"><%=Rs("zhiwei")%></td>
	<td><%=Rs("jingyan")%></td>
    <td><%=Rs("DateAndTime")%></td>
    <td><a href="job.asp?Action=Mod&ID=<%=Rs("ID")%>">�޸�</a>&nbsp;<a href="job.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">ɾ��</a></td>
  </tr>
<%
		i = i + 1
		Rs.MoveNext
		Loop
		ShowPage(8)
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

Sub chakan()
%>
<form name="myform" method="post" action="job.asp?Action=chakan">
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr align="center" bgcolor="#339933">
    <td width="7%">ѡ��</td>
    <td width="12%">ӦƸְλ</td>
    <td width="10%">����</td>
	<td width="13%">����</td>
	<td width="11%">�绰</td>
    <td width="15%">�鿴����</td>
    <td width="21%">ӦƸʱ��</td>
    <td width="11%">����</td>
  </tr>
<%  
	Sql ="Select * From [job]  Order By ID Desc "
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Response.Write("<tr align=""center"" class=""t_rowh""><td colspan=""5""><strong>û���κμ�¼!</strong></td></tr>")
	Else
		i=0
		Rs.PageSize = Cint(Line)
		Rs.AbsolutePage = Cint(Page)
		Do While Not Rs.Eof And i < Rs.PageSize
%>
  <tr align="center" class="t_row" onmouseover="this.className='t_rowh'" onmouseout="this.className='t_row'">
    <td><input name="ID" type="checkbox" id="ID" value="<%=Rs("ID")%>"></td>
    <td align="center">&nbsp;<%=Rs("title")%></td>
    <td align="center"><%=Rs("sender")%></td>
	<td><%=Rs("email")%></td>
	<td><%=Rs("Tel")%></td>
    <td><a href="?Action=Show&ID=<%=Rs("ID")%>">�鿴����</a></td>
    <td><%=Rs("DateAndTime")%></td>
    <td><a href="about.asp?Action=CKMod&ID=<%=Rs("ID")%>"></a>&nbsp;<a href="job.asp?Action=CKDel&ID=<%=Rs("ID")%>" onClick="return Delete();">ɾ��</a></td>
  </tr>
<%
		i = i + 1
		Rs.MoveNext
		Loop
		ShowPage(8)
	End If
	Rs.Close
%>
</table>
</form>
<%
end Sub


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
<br />
<form name="myform" method="post" action="job.asp?Action=SaveAdd" onsubmit="return doCheck();">
<table border="0" width="95%" align="center" cellpadding="0" cellspacing="1" >
  <tr>
    <td colspan="2"  class="rgt03-bt">������Ƹ</td>
  </tr>
  <tr class="t_row">
    <td width="30%"><u>��Ƹְλ</u></td>
    <td width="70%"><input name="zhiwei" type="text" id="zhiwei" size="20"> 
      &nbsp;&nbsp; <span class="hui12">�������飺
      <input name="jingyan" type="text" id="jingyan" size="20" />
      </span></td>
  </tr>
  <style>
  .ke-container{ width:675px !important;}
  </style>
   <tr class="t_row">
    <td><u>�ϴ�����</u></td>
    <td>		<p><input type="text" id="url1" value="" name="pic" /> <input type="button" id="image1" value="ѡ��ͼƬ" />������ͼƬ + �����ϴ���</p></td>
  </tr>
 
  <tr class="t_row">
    <td valign="top"><u>����Ҫ��</u>:</td>
    <td>
	<textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content1"> <%=Content%></textarea></td>
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
	Dim ID,Title,Pic,Hits,Content
	Sql="Select * From [join] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "ְλ������!",Http_Referer,1
	Else
		ID=Rs("ID")
		zhiwei=Rs("zhiwei")
		jingyan=Rs("jingyan")
		address=Rs("address")
		gongzi=Rs("gongzi")
		yriqi=Rs("yriqi")
		pic=rs("pic")
		renshu=Rs("renshu")
		Content=Rs("Content")
		DateAndTime=Rs("DateAndTime")

		
		
	End If
	Rs.Close
%>


<form name="myform" method="post" action="job.asp?Action=SaveMod" onsubmit="return doCheck();">
<input  name="ID" type="hidden" value="<%=ID%>"/>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td  class="t_title">�޸�ְλ</td>         
     <td><input type="text"  name="zhiwei" id="zhiwei" value="<%=zhiwei%>" size="48" /></td>
  </tr>
  <tr class="t_row">
    <td><u>��������</u></td>
    <td><input type="text" name="jingyan" id="jingyan" value="<%=jingyan%>" size="48"/></td>
  </tr>
    <tr class="t_row">
    <td><u>�ϴ�����</u></td>
    <td>		<p><input type="text" id="url1" value="<%=pic%>" name="pic" /> <input type="button" id="image1" value="ѡ��ͼƬ" />������ͼƬ + �����ϴ���</p></td>
  </tr>
  <tr class="t_row">
    <td>ְλ���� 
      <li>�����밴Shift+Enter������һ���밴Enter</li></td>
    <td>
		  <textarea id="content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server" name="content2"> <%=Content%></textarea>
	</td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="�ύ"></td>
  </tr>
</table>
</form>
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
<%
End Sub

Sub aboutShow()
	Dim ID,Title,Pic,Content,Hits,DateAndTime
	Sql="Select * From [job] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "ְλ������!",Http_Referer,1
	Else
		
		Rs.Update
		ID=Rs("ID")
		Title=Rs("Title")
		Pic=Rs("Sender")
		Content=Rs("Content")
		Hits=Rs("Tel")
		Picture=Rs("Email")
		DateAndTime=Rs("DateAndTime")
	End If
	Rs.Close
%>
<table border="1" align="center" cellpadding="1" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title"><%=Title%></td>
  </tr>
  <tr class="t_rowh">
    <td>&nbsp;</td>
    <td align="right">��&nbsp;������<a href="job.asp?Action=Mod&ID=<%=ID%>">�޸�</a>&nbsp;<a href="job.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">ɾ��</a>&nbsp;<a href="<%=Http_Referer%>">����</a>&nbsp;��</td>
  </tr>
  
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px">
	<table width="600" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th colspan="8" scope="row" style="font-size:16px">ӦƸ������Ϣ</th>
    </tr>
  <tr>
    <th  scope="row">&nbsp;</th>
    <td width="123">&nbsp;</td>

  </tr>
  <tr>
    <th  scope="row">ӦƸְλ:</th>
    <td width="123"><%=Title%></td>

  </tr>
  <tr>
    <th scope="row">&nbsp;��&nbsp;&nbsp;&nbsp; ����</th>
    <td>&nbsp;<%=Pic%></td>

  </tr>
    <tr>
    <th scope="row">&nbsp;ӦƸʱ��</th>
    <td>&nbsp;<%=DateAndTime%></td>
 
  </tr>
  <tr>
    <th scope="row">��������</th>
    <td>&nbsp;<%=Picture%></td>

  </tr>
  <tr>
    <th scope="row">&nbsp;��ϵ�绰</th>
    <td>&nbsp;<%=Hits%></td>
 
  </tr>
    <tr>
    <th scope="row">&nbsp;�������</th>
    <td>&nbsp;<%=Content%></td>
 
  </tr>

</table>

	
	</td>
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
	Dim zhiwei,address,xueli,pic,jingyan,gongzi,renshu,yriqi,content
	zhiwei=Request.Form("zhiwei")
	jingyan=Request.Form("jingyan")
	xueli=Request.Form("xueli")
	pic=Request.Form("pic")
	address=Request.Form("address")
	gongzi=Request.Form("gongzi")
	renshu=Request.Form("renshu")			
	yriqi=Request.Form("yriqi")
	bumen=Request.Form("bumen")
		
	Content=CheckChar(Trim(Request.Form("Content1")),"ְλ����",10000,"2")
	If ErrContent<>"" then Msg ErrContent,Http_Referer,1
	Rs.Open "Select * From [join]",Conn,1,2
		Rs.AddNew
		Rs("zhiwei")=zhiwei
		Rs("pic")=pic
		Rs("jingyan")=jingyan
		Rs("xueli")=xueli
		Rs("address")=address
		Rs("gongzi")=gongzi
		Rs("renshu")=renshu
		Rs("yriqi")=yriqi


		
		Rs("Content")=Content
		Rs("DateAndTime")=Now()
	    Rs("type")=1
		Rs.Update
	Rs.Close
	Msg "ְλ��ӳɹ�!","job.asp",0
End Sub
Sub SaveMod()
	Dim ID, title,xinghao,dunwei,adress,pic,gongzi,yriqi,renshu,Content
	ID=Request.Form("ID")
	zhiwei=Request.Form("zhiwei")
	jingyan=Request.Form("jingyan")
	xueli=Request.Form("xueli")
	pic=Request.Form("pic")
	address=Request.Form("address")
	gongzi=Request.Form("gongzi")
	renshu=Request.Form("renshu")			
	yriqi=Request.Form("yriqi")
	bumen=Request.Form("bumen")
	
	Content=Request.Form("Content2")
	If ErrContent<>"" then Msg ErrContent,Http_Referer,1
	Rs.Open "Select * From [join] Where ID="&ID,Conn,1,2
		Rs("zhiwei")=zhiwei
		Rs("jingyan")=jingyan
		Rs("xueli")=xueli
		Rs("pic")=pic
		Rs("address")=address
		Rs("gongzi")=gongzi
		Rs("renshu")=renshu
		Rs("yriqi")=yriqi
		
	
		Rs("Content")=Content
		Rs("DateAndTime")=Now()
	    Rs("type")=1
		Rs.Update
	Rs.Close
	Msg "ְλ�޸ĳɹ�!","job.asp",0
End Sub
Sub SaveDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"ְλ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"job.asp",1
	Rs.Open "Select * From [join] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "ְλ������,������ִ��ɾ������!",Http_Referer,1
	Else
		Rs.Delete
	End If
	Rs.Close
	Msg "ְλID:"&ID&"�ѳɹ�ɾ��,�����Իָ�!",Http_Referer,0
End Sub

Sub CKDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"ְλ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"job.asp",1
	Rs.Open "Select * From [job] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "ְλ������,������ִ��ɾ������!",Http_Referer,1
	Else
		Rs.Delete
	End If
	Rs.Close
	Msg "ְλID:"&ID&"�ѳɹ�ɾ��,�����Իָ�!",Http_Referer,0
End Sub

%>
<br />
<div  style="clear:both"></div>
	<div class="bottom">
		<p>��Ȩ����:����������Ϣ�������޹�˾ ��ַ:�����б�����·88�ű�����Է4-4-5F ��������:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
</html>
