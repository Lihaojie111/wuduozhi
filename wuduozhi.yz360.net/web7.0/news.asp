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
<!--ͷ��-->
<div class="wramp">
  <!--#include file="head.asp"-->
  <!--���-->

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
    <td colspan="2" class="t_title">��Ϣ����</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2"><strong>����ѡ��</strong>��<a href="News.asp">��Ϣ����</a>&nbsp;|&nbsp;<a href="News.asp?Action=Add">�����Ϣ</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=News">�������</a></td>
  </tr>
  <tr class="t_rowh">
    <form method="get" action="News.asp">
      <td width="60%">����������&nbsp;
        <input name="Action" type="hidden" id="Action" value="Search">
        <input name="Word" type="text" id="Word" size="48"></td>
    </form>
    <td width="40%" align="right">�鿴���ࣺ&nbsp;
      <select name="ClassID" id="ClassID" onchange="window.location='News.asp?ClassID='+this.options[this.selectedIndex].value+'';">
        <option value="0">��ת������</option>
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
			Msg "��ѡ��Ҫ��ѯ�ķ���","News.asp",1
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
	Rs.Open "Select * From [News] "&Sql&" Order By ID Desc",Conn,1,1
%>
<form name="myform" method="post" action="News.asp?Action=SDel">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr align="center" style="background-color:#009933">
      <td width="5%">ѡ��</td>
	  <td width="6%">ID</td>
      <td width="45%">����</td>
      <td width="14%">����</td>
      <td width="10%">�������</td>
      <td width="20%">����</td>
    </tr>
    <%
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
	   <td><%=rs("id")%></td>
      <td align="left">&nbsp;<a href="?Action=Show&ID=<%=Rs("ID")%>"><%=Replace(Rs("Title"),Request("Word"),"<strong>"&Request("Word")&"</strong>")%></a></td>
      <td><%=ShowClassName(Rs("ClassID"))%></td>
      <td><%=(Rs("Hits"))%></td>
      <td><a href="News.asp?Action=Topis&ID=<%=Rs("ID")%>">
        <%If Rs("Topis")=1 Then%>
        <font color="#009933">�Ƽ�</font>
        <%Else%>
        �Ƽ�
        <%End If%>
        </a>&nbsp;&nbsp;<a href="News.asp?Action=Mod&ID=<%=Rs("ID")%>">�޸�</a>&nbsp;<a href="News.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">ɾ��</a></td>
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
   if(confirm("ɾ�����ݺ��ָܻ���ȷ��Ҫɾ����"))
     return true;
   else
     return false;
}
</Script>
<%
End Sub

Sub NewsAdd()
%>
<form name="myform" method="post" action="News.asp?Action=SaveAdd" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">�����Ϣ</td>
    </tr>
    <tr class="t_rowh">
      <td width="20%">��������</td>
      <td width="80%"><select name="ClassID" id="ClassID">
          <%SelectClass(0)%>
        </select>
        &nbsp;<font color="#FF0000">��</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>ע�⣺����ָ��Ϊ�����ӷ���ķ���</strong></font></td>
    </tr>
    <tr class="t_rowh">
      <td>��Ϣ����</td>
      <td><input name="Title" type="text" id="Title" size="48">
        &nbsp;<font color="#FF0000">��</font> �� 100���ַ�</td>
    </tr>
    <tr class="t_row">
      <td>��Ϣ����������</td>
      <td><input name="neirongmiaoshu" type="text" id="neirongmiaoshu" value="" size="48">
      (Ҳ�ɲ���)</td>
    </tr>
    <tr class="t_row">
      <td>����SEO�Ż�--��Ϣ������</td>
      <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="��Ϣ�������������ͬ..." size="48">
      (Ҳ�ɲ���)</td>
    </tr>
    <tr class="t_row">
      <td>����SEO�Ż�--�����ؼ��֣�</td>
      <td><input name="Keywords" type="text" id="Keywords" value="���������Ϣ����ʱ�Ĺؼ���..." size="48">
      (Ҳ�ɲ���)</td>
    </tr>
    <tr class="t_row" style="display:none">
      <td>��Ϣ����</td>
      <td><input name="Author" type="text" id="Author" size="40">
        &nbsp;&nbsp;��ǰΪ��ֵʱ�����������ӡ���Ч</td>
    </tr>
    <tr class="t_row" style="display:none">
      <td><u>��������</u></td>
      <td><input name="Url" type="text" id="Url" size="40">
        &nbsp;&nbsp;��ʽ����http://��</td>
    </tr>
    <tr class="t_rowh">
      <td><u>��Ƶ�ļ�</u></td>
      <td><input type="text" id="url" value="" name="files" /> <input type="button" id="insertfile" value="ѡ���ļ�" />
        &nbsp;&nbsp;���༭���в���ͼƬʱ����ѡ���������</td>
    </tr>
    <tr class="t_rowh">
      <td><u>��Ϣ����</u></td>
      <td><input name="Topis" type="checkbox" id="Topis" value="True">
        &nbsp;
        <label for="Topis">�Ƽ���Ϣ</label>
        &nbsp;&nbsp;<label for="TitB"></label></td>
    </tr>
    <tr class="t_row">
      <td colspan="2"><u>��Ϣ����</u></td>
    </tr>
        <tr class="t_row">
      <td colspan="2">	<textarea id="Content1" name="Content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server"> <%=Content%></textarea></td>
    </tr>
    <tr class="t_row">
      <td>&nbsp;</td>
      <td><input type="submit" value="�ύ"></td>
    </tr>
  </table>
</form>
<%
End Sub

Sub NewsMod()
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Pictures,Topis,Hits,Red,TitB,Content,DescriptionWord,neirongmiaoshu,Keywords
	Sql="Select * From [News] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "��Ϣ������!","New.asp",1
	Else
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		DescriptionWord=Rs("DescriptionWord")
		neirongmiaoshu=Rs("neirongmiaoshu")
		Keywords=Rs("Keywords")
		Title=Rs("Title")
		Author=Rs("Author")
	
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		Red=Rs("Red")
		TitB=Rs("TitB")
		Content=Rs("Content")
	End If
	Rs.Close
%>
<form name="myform" method="post" action="News.asp?Action=SaveMod" onsubmit="return doCheck();">
  <table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
    <tr>
      <td colspan="2" class="t_title">�޸���Ϣ</td>
    </tr>
    <tr class="t_rowh">
      <td width="20%"><u>��������</u></td>
      <td width="80%"><input name="ID" type="hidden" value="<%=ID%>">
        <select name="ClassID" id="ClassID">
          <%SelectClass(ClassID)%>
        </select>
        &nbsp;<font color="#FF0000">��</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>ע�⣺����ָ��Ϊ�����ӷ���ķ���</strong></font></td>
    </tr>
    <tr class="t_rowh">
      <td><u>��Ϣ����</u></td>
      <td><input name="Title" type="text" id="Title" value="<%=Title%>" size="48">
        &nbsp;<font color="#FF0000">��</font> �� 100���ַ�</td>
    </tr>
    
    <tr class="t_row">
      <td><u>��Ϣ����������</u></td>
      <td><input name="neirongmiaoshu" type="text" id="neirongmiaoshu" value="<%=neirongmiaoshu%>" size="48"></td>
    </tr>
    <tr class="t_row">
      <td><u>����SEO�Ż�--��Ϣ������</u></td>
      <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="<%=DescriptionWord%>" size="48"></td>
    </tr>
    <tr class="t_row">
      <td><u>����SEO�Ż�--�����ؼ��֣�</u></td>
      <td><input name="Keywords" type="text" id="Keywords" value="<%=Keywords%>" size="48"></td>
    </tr>
    <tr class="t_row" style="display:none">
      <td><u>��Ϣ����</u></td>
      <td><input name="Author" type="text" id="Author" value="<%=Author%>" size="40">
        &nbsp;&nbsp;��ǰΪ��ֵʱ�����������ӡ���Ч</td>
    </tr>
    <tr class="t_row" style="display:none">
      <td><u>��������</u></td>
      <td><input name="Url" type="text" id="Url" value="<%=Url%>" size="40">
        &nbsp;&nbsp;��ʽ����http://��</td>
    </tr>
    <tr class="t_rowh">
      <td><u>��Ƶ�ļ�</u></td>
      <td><input type="text" id="url" value="<%=url%>" name="files" /> <input type="button" id="insertfile" value="ѡ���ļ�" />�ϴ���Ƶ�ļ���100M���£�</td>
    </tr>
    <tr class="t_rowh">
      <td><u>��Ϣ����</u></td>
      <td><input name="Topis" type="checkbox" id="Topis" value="True"<%If Topis=1 Then%> checked<%End If%>>
        &nbsp;
        <label for="Topis">�Ƽ���Ϣ</label>
        &nbsp;&nbsp;<label for="TitB"></label></td>
    </tr>
    <tr class="t_row">
      <td colspan="2"><u>��Ϣ����</u></td>
    </tr>
        <tr class="t_row">
      <td colspan="2">	<textarea id="Content1" name="Content1" cols="70" rows="8" style="width:700px;height:400px;visibility:hidden;" runat="server"> <%=Content%></textarea></td>
    </tr>
    <tr class="t_row">
      <td>&nbsp;</td>
      <td><input type="submit" value="�ύ"></td>
    </tr>
  </table>
</form>
<%
End Sub

Sub NewsShow()
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Pictures,Topis,Hits,Red,TitB,Content,DateAndTime,neirongmiaoshu,DescriptionWord,Keywords
	Sql="Select * From [News] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������!","New.asp",1
	Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		Title=Rs("Title")
		Author=Rs("Author")
		Url=Rs("url")
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		Red=Rs("Red")
		TitB=Rs("TitB")
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
    <td>&nbsp;Location:&nbsp;&nbsp;
      <%Call ShowClassPath(ClassID)%></td>
    <td align="right">��&nbsp;������<a href="News.asp?Action=Topis&ID=<%=ID%>">
      <%If Topis=1 Then%>
      <font color="#009933">�Ƽ�</font>
      <%Else%>
     �Ƽ�
      <%End If%>
      </a>&nbsp;<a href="News.asp?Action=Mod&ID=<%=ID%>">�޸�</a>&nbsp;<a href="News.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">ɾ��</a>&nbsp;<a href="<%=Http_Referer%>">����</a>&nbsp;��</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,Title,ClassID,Hits,DateAndTime,Content,"News"%></td>
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
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Topis,Hot,Red,TitB,Content,neirongmiaoshu,DescriptionWord,Keywords
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"��������",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"��Ϣ����",100,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"����SEO�Ż�--�����ؼ��֣�1",2000,"012")
	neirongmiaoshu=CheckChar(Trim(Request.Form("neirongmiaoshu")),"��Ϣ��������",2000,"12")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"����SEO�Ż�--�����ؼ��֣�2",2000,"012")
	Author=CheckChar(Trim(Request.Form("Author")),"��Ϣ����",50,"12")
	Url=CheckChar(Trim(Request.Form("files")),"��Ƶ�ļ�",100,"12")
	Pic=CheckChar(Trim(Request.Form("Pic")),"Ĭ��ͼƬ",100,"12")
	Picture=CheckChar(Trim(Request.Form("Picture")),"����ͼƬ",65535,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"��Ϣ�ö�",0,"")
	Hot=CheckChar(Trim(Request.Form("Hot")),"������Ϣ",0,"")
	Red=CheckChar(Trim(Request.Form("Red")),"�����ɫ",0,"")
	TitB=CheckChar(Trim(Request.Form("TitB")),"����Ӵ�",0,"")
	Content=CheckChar(Request.Form("Content1"),"��Ϣ����",605535,"02")
	If ClassID<=0 Then ErrContent = ErrContent & "������ӷ���!"
	If ErrContent<>"" Then Msg ErrContent,"News.asp?Action=Add",1
	If Conn.Execute("Select Child From Class Where ClassID="&ClassID)(0)>0 Then
		Msg "�������ຬ����Ŀ¼,��ѡ����������!","News.asp?Action=Add",1
	End If
	Rs.Open "Select * From [News]",conn,1,2
		Rs.Addnew
		ID=Rs("ID")
		Rs("Title")=Title
		Rs("Content")=Content

			Rs("Url")=Url
	     Rs("neirongmiaoshu")= neirongmiaoshu
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
	Msg "��Ϣ��ӳɹ�!","News.asp",0
End Sub

Sub SaveMod()
	Dim ID,ClassID,Title,Author,Url,Pic,Picture,Topis,Hot,Red,TitB,Content,neirongmiaoshu,DescriptionWord,Keywords
	ID=CheckChar(Trim(Request.Form("ID")),"��Ϣ���",0,"03")
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"��������",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"��Ϣ����",100,"012")
	neirongmiaoshu=CheckChar(Trim(Request.Form("neirongmiaoshu")),"��Ϣ��������",2000,"12")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"����SEO�Ż�--�����ؼ��֣�1",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"����SEO�Ż�--�����ؼ��֣�2",2000,"012")
	Author=CheckChar(Trim(Request.Form("Author")),"��Ϣ����",50,"12")
	Url=CheckChar(Trim(Request.Form("files")),"��������",100,"12")
	Pic=CheckChar(Trim(Request.Form("Pic")),"Ĭ��ͼƬ",100,"12")
	Picture=CheckChar(Trim(Request.Form("Picture")),"����ͼƬ",65535,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"��Ϣ�ö�",0,"")
	Hot=CheckChar(Trim(Request.Form("Hot")),"������Ϣ",0,"")
	Red=CheckChar(Trim(Request.Form("Red")),"�����ɫ",0,"")
	TitB=CheckChar(Trim(Request.Form("TitB")),"����Ӵ�",0,"")
	Content=CheckChar(Request.Form("Content1"),"��Ϣ����",605535,"02")
	If ErrContent<>"" Then Msg ErrContent,"News.asp?Action=Mod&ID="&ID,1
	If Conn.Execute("Select Child From Class Where ClassID="&ClassID)(0)>0 Then
		Msg "�������ຬ����Ŀ¼,��ѡ����������!","News.asp?Action=Mod&ID="&ID,1
	End If
	Rs.Open "Select * From [News] Where ID="&ID,conn,1,2
		Rs("Title")=Title
		Rs("Content")=Content
 
		Rs("Url")=Url
 			Rs("neirongmiaoshu")=neirongmiaoshu
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
	Msg "��Ϣ�޸ĳɹ�!","News.asp",0
End Sub

Sub NewsDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��Ϣ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select * From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ��ɾ������!",Http_Referer,1
	Else
		DelFiles(Rs("Picture")&"")
		Rs.Delete
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub

Sub NewsSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"��Ϣ���",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select * From [News] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ��ɾ������!",Http_Referer,1
	Else
		Do While Not Rs.Eof
		DelFiles(Rs("Picture")&"")
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub

Sub NewsTopis()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��Ϣ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,Topis From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ�в���!",Http_Referer,1
	Else
		If Rs("Topis")=1 Then
			Rs("Topis")=0
		Else
			Rs("Topis")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"�ö����óɹ�!",Http_Referer,0
End Sub

Sub NewsHot()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��Ϣ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,Hits From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ�в���!",Http_Referer,1
	Else
		If Rs("Hits")>=Cint(m_WebSet(3)) Then
			Rs("Hits")=Cint(Cint(m_WebSet(3))/2)
		Else
			Rs("Hits")=Cint(m_WebSet(3))
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"�������óɹ�!",Http_Referer,0
End Sub

Sub NewsRed()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��Ϣ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,Red From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ�в���!",Http_Referer,1
	Else
		If Rs("Red")=1 Then
			Rs("Red")=0
		Else
			Rs("Red")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"�����ɫ���óɹ�!",Http_Referer,0
End Sub

Sub NewsTitB()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��Ϣ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"News.asp",1
	Rs.Open "Select ID,TitB From [News] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ�в���!",Http_Referer,1
	Else
		If Rs("TitB")=1 Then
			Rs("TitB")=0
		Else
			Rs("TitB")=1
		End If
		Rs.Update
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"����Ӵ����óɹ�!",Http_Referer,0
End Sub 

%>


		   </div>
		 </div>
		<div class="bottom">
		<p>��Ȩ����:����������Ϣ�������޹�˾ ��ַ:�����б�����·88�ű�����Է4-4-5F ��������:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
</html>
