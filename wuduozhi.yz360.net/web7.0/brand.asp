<!--#include file="System.asp"-->
<script language="javascript" src="Calendar.js" >
</script>
<%
Head()
Layout="brand"
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">��Ϣ����</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2"><strong>����ѡ��</strong>��<a href="brand.asp">��Ϣ����</a>&nbsp;|&nbsp;<a href="brand.asp?Action=Add">�����Ϣ</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=brand">�������</a></td>
  </tr>
  <tr class="t_rowh">
    <form method="get" action="brand.asp">
    <td width="60%">����������&nbsp;<input name="Action" type="hidden" id="Action" value="Search"><input name="Word" type="text" id="Word" size="48"></td>
    </form>
	<td width="40%" align="right">�鿴���ࣺ&nbsp;<select name="ClassID" id="ClassID" onchange="window.location='brand.asp?ClassID='+this.options[this.selectedIndex].value+'';"><option value="0">��ת������</option><%SelectClass(Cint(ClassID))%></select>&nbsp;&nbsp;&nbsp;</td>
  </tr>
</table>
<%
Select Case Action
	Case "Add"
		Call brandAdd()
	Case "Mod"
		Call brandMod()
	Case "Show"
		Call brandShow()
	Case "Del"
		Call brandDel()
	Case "SDel"
		Call brandSDel()
	Case "Topis"
		Call brandTopis()
	Case "SaveMod"
		Call SaveMod()
	Case "SaveAdd"
		Call SaveAdd()
	Case Else
		Call brandList()
End Select

Sub brandList()
	If Action="Search" Then
		Query_String="Action=Search&Word="&Request("Word")&"&"
		Sql="Where Title Like '%"&Request("Word")&"%'"
	ElseIf Cint(ClassID)>0 Then
		Dim RsTemp,ChildID
		Set RsTemp=Conn.Execute("Select ParentPath From [Class] Where ClassID="&ClassID)
		If Not RsTemp.Eof Then
			Dim ParentPath
			ParentPath = RsTemp("ParentPath")
		Else
			Msg "��ѡ��Ҫ��ѯ�ķ���","brand.asp",1
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
	Sql = "Select * From [brand] "&Sql&" Order By ID Desc"
	Rs.Open Sql,Conn,1,1
%>
<form name="myform" method="post" action="brand.asp?Action=SDel">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr align="center" class="t_title">
    <td width="5%">ѡ��</td>
    <td>����</td>
    <td width="18%">����</td>
    <td width="10%">�������</td>
    <td width="10%">����</td>
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
    <td align="left">&nbsp;<a href="?Action=Show&ID=<%=Rs("ID")%>"><%=Replace(Rs("Title"),Request("Word"),"<strong>"&Request("Word")&"</strong>")%></a></td>
    <td><%=ShowClassName(Rs("ClassID"))%></td>
    <td><%=(Rs("Hits"))%></td>
    <td><a href="brand.asp?Action=Topis&ID=<%=Rs("ID")%>"><%If Rs("Topis")=1 Then%><font color="#CC0000">��</font><%Else%>��<%End If%></a>&nbsp;<a href="brand.asp?Action=Mod&ID=<%=Rs("ID")%>">�޸�</a>&nbsp;<a href="brand.asp?Action=Del&ID=<%=Rs("ID")%>" onClick="return Delete();">ɾ��</a></td>
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

Sub brandAdd()
%>
<script language="javascript">
// ����˵��
// s_Type : �ļ����ͣ�����ֵΪ"image","flash","media","file"
// s_Link : �ļ��ϴ������ڽ����ϴ��ļ�·���ļ����ı���
// s_Thumbnail : �ļ��ϴ������ڽ����ϴ�ͼƬʱ������������ͼ�ļ���·���ļ����ı�������δ��������ͼʱ�����ؿ�ֵ��ԭͼ��s_Link�������գ��˲���ר��������ͼ
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//����style=coolblue,ֵ��������ʵ����Ҫ�޸�Ϊ������ʽ��,ͨ������ʽ�ĺ�̨�������ﵽ���������ϴ��ļ����ͼ��ļ���С
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=brand&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>

<form name="myform" method="post" action="brand.asp?Action=SaveAdd" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">�����Ϣ</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>��������</u></td>
    <td width="80%"><select name="ClassID" id="ClassID"><%SelectClass(0)%></select>&nbsp;<font color="#FF0000">��</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>ע�⣺����ָ��Ϊ�����ӷ���ķ���</strong></font></td>
  </tr>
  <tr class="t_row">
    <td><u>��Ϣ����</u></td>
    <td><input name="Title" type="text" id="Title" size="60">&nbsp;<font color="#FF0000">��</font> �� 100���ַ�&nbsp;&nbsp;&nbsp;&nbsp;<input name="Topis" type="checkbox" id="Topis" value="True"><label for="Topis">�Ƽ���Ϣ</label></td>
  </tr>
  <tr class="t_row">
    <td><u>����SEO�Ż�--��Ϣ������</u></td>
    <td><input name="DescriptionWord" type="text" id="DescriptionWord" size="60" />      
      ��Ҳ�ɲ��</td>
  </tr>
  <tr class="t_row">
    <td><u>����SEO�Ż�--�����ؼ��֣�</u></td>
    <td><input name="Keywords" type="text" id="Keywords" size="60">
      ��Ҳ�ɲ��</td>
  </tr>
  <tr class="t_row">
    <td>��Ϣ�������ڣ�</td>
    <td><input name="DateAndTime" type="text" id="DateAndTime" value="<%=now()%>" size="30" />    
      ��Ĭ��Ϊ��ǰʱ�䣩</td>
  </tr>
   <tr  class="t_row">
      <td valign="middle">ѡ����Ϣ���ڣ�</td>
     <td align="left" valign="middle"><input name="rz" id= "rz" type="text" class="input3" size=20  onfocus="setday(this)" readonly />
     �����и������ڣ�</td>
	 </tr>
	<!--<tr class="t_row">
    <td>��ѡ���Ա�鿴Ȩ�ޣ�</td>
    <td><label><input type="checkbox" name="mianfei" value="1" />��ѻ�Ա</label><input type="checkbox" name="putong" value="2" />��ͨ��Ա</label><input type="checkbox" name="gaoji" value="3" /> �߼���Ա
      </label> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#FF0000"> ע��ѡ�б�ʾ�û�Ա���Բ鿴������Ϣ��</span></td>
  </tr> -->
	 
<!--  <tr class="t_row">
    <td><u>Ĭ��ͼƬ</u></td>
    <td><input name="Pic" type="text" id="Pic" size="48" value="/images/td_img1.jpg">&nbsp;&nbsp;<input type=button value="�ϴ�ͼƬ..." onClick="showUploadDialog('image', 'myform.Pic', '')">&nbsp;&nbsp;<a href="#" target="_blank" id="yulan" onclick="this.href=document.getElementById('Pic').value">Ԥ��ͼƬ</a></td>
  </tr>-->
  <tr class="t_row">
    <td><u>��Ϣ����</u><li>��ϵͳ�����ͼƬ���Ƶ���վ�������ϣ�ϵͳ������ͼƬ�Ĵ�С��Ӱ���ٶȡ�</li><li>�����밴Shift+Enter������һ���밴Enter</li></td>
    <td><textarea name="Content" style="display:none"><%=Content%></textarea> <iframe id="myeditor" src="../Ok_Editor/ewebeditor.htm?id=Content&amp;style=brand&amp;maxlen=300" frameborder="0" scrolling="No" width="95%" height="405"></iframe></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="�ύ"></td>
  </tr>
</table>
</form>
<%
End Sub

Sub brandMod()
	Dim ID,ClassID,Title,Price,Pic,Picture,Pictures,Topis,Hits,Content,TempPicture,DescriptionWord,Keywords,mianfei,putong,gaoji
	Sql="Select * From brand Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "��Ϣ������!",Http_Referer,1
	Else
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		DescriptionWord=Rs("DescriptionWord")
		Keywords=Rs("Keywords")
		Title=Rs("Title")
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		Content=Rs("Content")
		DateAndTime=Rs("DateAndTime")
		mianfei=Rs("mianfei")
		putong=Rs("putong")
		gaoji=Rs("gaoji")
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
	var arr = showModalDialog("../Ok_Editor/dialog/i_upload.htm?style=brand&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>
<form name="myform" method="post" action="brand.asp?Action=SaveMod" onsubmit="return doCheck();">
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">�޸���Ϣ</td>
  </tr>
  <tr class="t_row">
    <td width="20%"><u>��������</u></td>
    <td width="80%"><input name="ID" type="hidden" value="<%=ID%>"><select name="ClassID" id="ClassID"><%SelectClass(ClassID)%></select>
    &nbsp;<font color="#FF0000">��</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>ע�⣺����ָ��Ϊ�����ӷ���ķ��ࡣ</strong></font></td>
  </tr>
  <tr class="t_row">
    <td><u>��Ϣ����</u></td>
    <td><input name="Title" type="text" id="Title" value="<%=Title%>" size="60">&nbsp;<font color="#FF0000">��</font> �� 100���ַ�&nbsp;&nbsp;&nbsp;&nbsp;<input name="Topis" type="checkbox" id="Topis" value="True"<%If Topis=1 Then%> checked<%End If%>>&nbsp;<label for="Topis">�Ƽ���Ϣ</label></td>
  </tr>
  <tr class="t_row">
    <td>����SEO�Ż�--��Ϣ������</td>
    <td><input name="DescriptionWord" type="text" id="DescriptionWord" value="<%=DescriptionWord%>" size="60"></td>
  </tr>
  <tr class="t_row">
    <td>����SEO�Ż�--�����ؼ��֣�</td>
    <td><input name="Keywords" type="text" id="Keywords" value="<%=Keywords%>" size="60"></td>
  </tr>
  <tr class="t_row">
    <td>��Ϣ�������ڣ�</td>
    <td><input name="DateAndTime" type="text" id="DateAndTime" value="<%=DateAndTime%>" size="30">
    ��Ĭ�����ڣ�</td>
  </tr>
  <tr  class="t_row">
      <td >������Ϣ���ڣ�</td>
      <td ><input name="rz" id= "rz" type="text" size=20  onfocus="setday(this)" readonly />
      �����и������ڣ�</td>
	 </tr>
<!--	<tr class="t_row">
    <td>��ѡ���Ա�鿴Ȩ�ޣ�</td>
    <td><input type="checkbox" name="mianfei" value="1" <%If mianfei=1 Then%> checked<%End If%> /><label>��ѻ�Ա</label><input type="checkbox" name="putong" value="2"<%If putong=2 Then%> checked<%End If%> />��ͨ��Ա<input type="checkbox" name="gaoji" value="3"<%If gaoji=3 Then%> checked<%End If%> />
    �߼���Ա</label> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#FF0000"> ע��ѡ�б�ʾ�û�Ա���Բ鿴������Ϣ��</span></td>
  </tr> -->
  <!--<tr class="t_row">
    <td><u>Ĭ��ͼƬ</u></td>
    <td><input name="Pic" type="text" id="Pic" size="48" value="<%=Pic%>">&nbsp;&nbsp;<input type=button value="�ϴ�ͼƬ..." onClick="showUploadDialog('image', 'myform.Pic', '')">&nbsp;&nbsp;<a href="#" target="_blank" id="yulan" onclick="this.href=document.getElementById('Pic').value">Ԥ��ͼƬ</a></td>
  </tr>-->
  
  <tr class="t_row">
    <td><u>��Ϣ����</u><li>��ϵͳ�����ͼƬ���Ƶ���վ�������ϣ�ϵͳ������ͼƬ�Ĵ�С��Ӱ���ٶȡ�</li><li>�����밴Shift+Enter������һ���밴Enter</li></td>
    <td><textarea name="Content" style="display:none"><%=Content%></textarea> <iframe id="myeditor" src="../Ok_Editor/ewebeditor.htm?id=Content&amp;style=brand&amp;maxlen=300" frameborder="0" scrolling="No" width="95%" height="405"></iframe></td>
  </tr>
  
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="�ύ"></td>
  </tr>
</table>
</form>
<%
End Sub

Sub brandShow()
	Dim ID,ClassID,Title,Pic,Picture,Pictures,Topis,Hits,Content,DateAndTime,DescriptionWord,Keywords
	Sql="Select * From [brand] Where ID="&Request.QueryString("ID")
	Rs.Open Sql,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������!",Http_Referer,1
	Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ID=Rs("ID")
		ClassID=Rs("ClassID")
		Title=Rs("Title")
		Pic=Rs("Pic")
		Picture=Rs("Picture")
		Topis=Rs("Topis")
		Hits=Rs("Hits")
		Content=Rs("Content")
		DateAndTime=Rs("DateAndTime")
		mianfei=Rs("mianfei")
		putong=Rs("putong")
		gaoji=Rs("gaoji")
	End If
	Rs.Close
%>
<table border="0" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title"><%=Title%></td>
  </tr>
  <tr class="t_rowh">
    <td>&nbsp;Locahost:&nbsp;&nbsp;<%Call ShowClassPath(ClassID)%></td>
    <td align="right">��&nbsp;������<a href="brand.asp?Action=Topis&ID=<%=ID%>"><%If Topis=1 Then%><font color="#CC0000">��</font><%Else%>��</a><a href="brand.asp?Action=Topis&ID=<%=ID%>">
      <%End If%>
    </a>&nbsp;<a href="brand.asp?Action=Mod&ID=<%=ID%>">�޸�</a>&nbsp;<a href="brand.asp?Action=Del&ID=<%=ID%>" onClick="return Delete();">ɾ��</a>&nbsp;<a href="<%=Http_Referer%>">����</a>&nbsp;��</td>
  </tr>
  <tr class="t_row">
    <td colspan="2" valign="top" style="padding:20px"><%ShowContent ID,Title,ClassID,Hits,DateAndTime,"<div align=""center""><img src="""&Pic&"""></div><br>"&Replace(Replace(Content,Chr(10),"<br>"),Chr(13),"&nbsp;")&"<br><br><strong>˵��������</strong>:"&Picture,"brand"%></td>
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
	Dim ID,ClassID,Title,Price,Pic,Picture,Topis,Content,DescriptionWord,Keywords
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"��������",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"��Ϣ����",100,"012")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"�����ؼ���1",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"�����ؼ���2",2000,"012")
	Picture=CheckChar(Trim(Request.Form("Picture")),"˵����",100,"12")
	Pic=CheckChar(Trim(Request.Form("Pic")),"ͼƬ",255,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"�Ƽ���Ϣ",0,"")
	Content=CheckChar(Request.Form("Content"),"��Ϣ����",65535,"02")
	DateAndTime=Trim(Request.Form("DateAndTime"))
	rz=Request.Form("rz")
	mianfei=Trim(Request.Form("mianfei"))
	putong=Trim(Request.Form("putong"))
	gaoji=Trim(Request.Form("gaoji"))
	If ClassID<=0 Then ErrContent = ErrContent & "������ӷ���!"
	If ErrContent<>"" Then Msg ErrContent,"brand.asp?Action=Add",1
	Rs.Open "Select * From [brand]",conn,1,2
		Rs.Addnew
		ID=Rs("ID")
		Rs("Title")=Title
		Rs("Content")=Content
		Rs("ClassID")=ClassID
		Rs("Pic")=Pic
		Rs("DescriptionWord")=DescriptionWord
		Rs("Keywords")=Keywords
		Rs("Picture")=Picture
	
	
		If Topis="True" Then
			Rs("Topis")=1
		Else
			Rs("Topis")=0
		End If
		If mianfei="1" Then
			Rs("mianfei")=1
		Else
			Rs("mianfei")=0
		End If
		If putong="2" Then
			Rs("putong")=2
		Else
			Rs("putong")=0
		End If
		If gaoji="3" Then
			Rs("gaoji")=3
		Else
			Rs("gaoji")=0
		End If
		
		if rz="" then
		Rs("DateAndTime")=DateAndTime
		else
		   Rs("DateAndTime")=rz	
		end if 	
		Rs("Hits")=0
		
		Rs.Update
	Rs.Close
	Msg "��Ϣ��ӳɹ�!","brand.asp",0
End Sub

Sub SaveMod()
	Dim ID,ClassID,Title,Price,Pic,Picture,Topis,Content,DescriptionWord,Keywords
	ID=CheckChar(Trim(Request.Form("ID")),"��Ϣ���",0,"03")
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"��������",0,"03")
	Title=CheckChar(Trim(Request.Form("Title")),"��Ϣ����",100,"012")
	Picture=CheckChar(Trim(Request.Form("Picture")),"˵����",100,"12")
	DescriptionWord=CheckChar(Trim(Request.Form("DescriptionWord")),"�����ؼ���1",2000,"012")
	Keywords=CheckChar(Trim(Request.Form("Keywords")),"�����ؼ���2",2000,"012")
	Pic=CheckChar(Trim(Request.Form("Pic")),"ͼƬ",255,"12")
	Topis=CheckChar(Trim(Request.Form("Topis")),"�Ƽ���Ϣ",0,"")
	Content=CheckChar(Request.Form("Content"),"��Ϣ����",65535,"02")
	DateAndTime=Trim(Request.Form("DateAndTime"))
	rz=Request.Form("rz")
	mianfei=Trim(Request.Form("mianfei"))
	putong=Trim(Request.Form("putong"))
	gaoji=Trim(Request.Form("gaoji"))
	If ErrContent<>"" Then Msg ErrContent,"brand.asp?Action=Mod&ID="&ID,1
	Rs.Open "Select * From [brand] Where ID="&ID,conn,1,2
		Rs("Title")=Title
		Rs("Content")=Content
		Rs("ClassID")=ClassID
		Rs("DescriptionWord")=DescriptionWord
		Rs("Keywords")=Keywords
		Rs("Pic")=Pic
		Rs("Picture")=Picture
		
		
		If Topis="True" Then
			Rs("Topis")=1
		Else
			Rs("Topis")=0
		End If
		
		If mianfei="1" Then
			Rs("mianfei")=1
		Else
			Rs("mianfei")=0
		End If
		If putong="2" Then
			Rs("putong")=2
		Else
			Rs("putong")=0
		End If
		If gaoji="3" Then
			Rs("gaoji")=3
		Else
			Rs("gaoji")=0
		End If
		
		
		if rz="" then
		Rs("DateAndTime")=DateAndTime
		else
		   Rs("DateAndTime")=rz	
		end if 	
		Rs.Update
	Rs.Close
	Msg "��Ϣ�޸ĳɹ�!","brand.asp",0
End Sub

Sub brandDel()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��Ϣ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"brand.asp",1
	Rs.Open "Select * From [brand] Where ID="&ID,Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ��ɾ������!",Http_Referer,1
	Else
'		DelFiles(Rs("Picture")&"")
		Rs.Delete
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub

Sub brandSDel()
	Dim ID
	ID=CheckChar(Trim(Request.Form("ID")),"��Ϣ���",0,"0")
	If ErrContent<>"" Then Msg ErrContent,"brand.asp",1
	Rs.Open "Select * From [brand] Where ID In ("&ID&")",Conn,1,2
	If Rs.Eof Then
		Msg "��Ϣ������,������ִ��ɾ������!",Http_Referer,1
	Else
		Do While Not Rs.Eof
'		DelFiles(Rs("Picture")&"")
		Rs.Delete
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "��ϢID:"&ID&"��ɾ��,�����Իָ�!",Http_Referer,0
End Sub

Sub brandTopis()
	Dim ID
	ID=CheckChar(Trim(Request.QueryString("ID")),"��Ϣ���",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"brand.asp",1
	Rs.Open "Select ID,Topis From [brand] Where ID="&ID,Conn,1,2
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

Bottom()
%>
