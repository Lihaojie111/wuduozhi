<!--#include file="System.asp"-->
<%Layout=Request.QueryString("Layout")
%>
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

<body>
<!--头部-->
<div class="wramp">
  <!--#include file="head.asp"-->
 
  <!--左侧-->

  <div id="content">
  
 <!--#include file="left.asp"-->
	
	<div class="c-right">
         <div class="xwgl">
		   <div class="xw_k">
		     
		
		   </div>

		   <div class="xw_nr">
	        <div  class="rgt03-bt">
			分类管理
			 </div>
		    <table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <!--<tr>
    <td class="t_title">分类管理(<%=Layout%>)</td>
  </tr>-->
  <tr>
    <td class="t_rowh"><strong>说明：</strong><br>分类管理可执行的操作有：添加、删除、修改、排序<br>1、不可删除包含有子分类的分类；<br>2、删除包含有内容的分类会把分类下所有的内容删除</td>
  </tr>
  <tr>
    <td class="t_rowh"><strong>操作选项</strong>&nbsp;&nbsp;<a href="Class.asp?Layout=<%=Layout%>">分类管理首页</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Add">新建分类</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Unite">分类内容合并</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=DelAll" onClick="return DeleteAll();">删除所有分类</a></td>
  </tr>
</table>
<Script language="JavaScript" type="text/JavaScript">
function DeleteAll()
{
   if(confirm("删除分类将同时删除所有分类中的所有内容，并且不能恢复！确定要删除所有分类吗？"))
     return true;
   else
     return false;
}
</Script>
<%
Select Case Action
	Case "Add"
		Call ClassAdd()
	Case "Mod"
		Call ClassMod()
	Case "Move"
		Call ClassMove()
	Case "Clear"
		Call ClassClear()
	Case "Del"
		Call ClassDel()
	Case "Up"
		Call ClassUp()
	Case "Down"
		Call ClassDown()
	Case "UpN"
		Call ClassUpN()
	Case "DownN"
		Call ClassDownN()
	Case "Unite"
		Call ClassUnite()
	Case "DelAll"
		Call ClassDelAll()
	Case "BatchMove"
		Call BatchMove()
	Case "SaveBatchMove"
		Call SaveBatchMove()
	Case "SaveBatchDateMove"
		Call SaveBatchDateMove()
	Case "SaveBatchDateDel"
		Call SaveBatchDateDel()
	Case "BatchDel"
		Call BatchDel()
	Case "SaveAdd"
		Call SaveAdd()
	Case "SaveMove"
		Call SaveMove()
	Case "SaveMod"
		Call SaveMod()
	Case "SaveUnite"
		Call SaveUnite()
	Case Else
		Call ClassList()
End Select

Sub ClassList()
	For i=0 To ubound(arrShowLine)
		arrShowLine(i)=False
	Next
	Dim SqlClass,RsClass,i,iDepth
	SqlClass="select * From Class Where Layout='"&Layout&"' order by RootID,OrderID"
	Set RsClass=Server.CreateObject("Adodb.Recordset")
	RsClass.open SqlClass,Conn,1,1
%>
<table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr class="t_title">
    <td>分类标题</td>
    <td width="380">操作</td>
  </tr>
<%
	If RsClass.Eof Then
		Response.Write("<tr align=""center"" class=""t_rowh""><td colspan=""5""><strong>没有任何分类</strong></td></tr>")
	Else
		Do While Not RsClass.Eof
		iDepth=rsClass("Depth")
		If rsClass("NextID")>0 Then
			arrShowLine(iDepth)=True
		Else
			arrShowLine(iDepth)=False
		End If
%>
  <tr class="t_row" onmouseover="this.className='t_rowh'" onmouseout="this.className='t_row'">
    <td>&nbsp;
<%
		If iDepth>0 Then
			For i=1 To iDepth
				If i=iDepth Then
					If rsClass("NextID")>0 Then
						Response.Write("<img src=""images/tree1.gif"" width=""17"" height=""16"" align=""absmiddle"">&nbsp;")
					Else
						Response.Write("<img src=""images/tree2.gif"" width=""17"" height=""16"" align=""absmiddle"">&nbsp;")
					End If
				Else
					If arrShowLine(i)=True Then
						Response.Write("<img src=""images/tree3.gif"" width=""17"" height=""16"" align=""absmiddle"">&nbsp;")
					Else
						Response.Write("<img src=""images/tree4.gif"" width=""17"" height=""16"" align=""absmiddle"">&nbsp;")
					End If
				End If
			Next
		End If
		If RsClass("Child")>0 Then
			Response.Write("<img src=""Images/+.gif"" width=""9"" height=""9"" border=""0"" align=""absmiddle"">")
		Else
			Response.Write("<img src=""Images/-.gif"" width=""9"" height=""9"" border=""0"" align=""absmiddle"">")
		End If
		If RsClass("Depth")=0 Then
			Response.Write("<strong>")
		End If
		Response.Write("&nbsp;<a href="""& Layout &".asp?ClassID="& RsClass("ClassID") &""">"&RsClass("ClassName")&"</a>")
'		Response.Write("</strong>&nbsp;P:"&RsClass("PrevID")&"&nbsp;N:"&RsClass("NextID")&"&nbsp;D:"&RsClass("Depth")&"&nbsp;Max:"&Ubound(arrShowLine)&"<strong>")
		If RsClass("Child")>0 Then
			Response.Write("("&RsClass("Child")&")")
		End If
		If RsClass("Depth")=0 Then
			Response.Write("</strong>")
		End If
%>
	</td>
    <td align="right">ID：<%=RsClass("ClassID")%>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Add&ParentID=<%=RsClass("ClassID")%>">添加子分类</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Mod&ClassID=<%=RsClass("ClassID")%>">修改</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Up<%If RsClass("Depth")<>0 Then Response.Write("N")%>&ClassID=<%=RsClass("ClassID")%>">上移</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Down<%If RsClass("Depth")<>0 Then Response.Write("N")%>&ClassID=<%=RsClass("ClassID")%>">下移</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Move&ClassID=<%=RsClass("ClassID")%>">移动</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Clear&ClassID=<%=RsClass("ClassID")%>" onClick="return ConfirmDel3();">清空</a>&nbsp;|&nbsp;<a href="Class.asp?Layout=<%=Layout%>&Action=Del&ClassID=<%=RsClass("ClassID")%>" onClick="<%if rsClass("Child")>0 then%>return ConfirmDel1();<%else%>return ConfirmDel2();<%end if%>">删除</a>&nbsp;</td>
  </tr>
<%
		RsClass.MoveNext
		Loop
%>
</table>
<Script language="JavaScript" type="text/JavaScript">
function ConfirmDel1()
{
   alert("此分类下还有子分类，必须先删除下属子分类后才能删除此分类！");
   return false;
}

function ConfirmDel2()
{
   if(confirm("删除分类将同时删除此分类中的所有内容，并且不能恢复！确定要删除此分类吗？"))
     return true;
   else
     return false;
	 
}
function ConfirmDel3()
{
   if(confirm("清空分类将把分类（包括子分类）的所有内容删除！确定要清空此分类吗？"))
     return true;
   else
     return false;
	 
}
</Script>
<%
		RsClass.Close
		Set RsClass=Nothing
	End If
End Sub

Sub ClassAdd()
	Dim ParentID
	ParentID=CheckChar(Trim(Request.QueryString("ParentID")),"父级分类编号",0,"")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	If ParentID <> "" Then
		ParentID = Cint(ParentID)
	Else
		ParentID = 0
	End If
%>
<form name="myform" method="post" action="Class.asp?Layout=<%=Layout%>&Action=SaveAdd">
<table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">添加分类</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2">说明：当所属分类是“请选择分类”时，则建立一级分类。</td>
  </tr>
  <tr class="t_row">
    <td width="40%">所属分类</td>
    <td width="60%"><select name="ParentID" id="ParentID"><option value="0">无（作为一级分类）</option><%SelectClass(ParentID)%></select></td>
  </tr>
  <tr class="t_row">
    <td>分类名称</td>
    <td><input name="ClassName" type="text" id="ClassName" value="" size="48"></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="Submit" value="添加分类"></td>
  </tr>
</table>
</Form>
<%
End Sub

Sub ClassMod()
	Dim ClassName,ParentID
	ClassID=CheckChar(Trim(ClassID),"分类编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	If ClassID <> "" Then
		ClassID = Cint(ClassID)
	Else
		ClassID = 0
	End If
	Sql="Select ClassID,ClassName,ParentID From Class Where ClassID="&ClassID&" And Layout='"&Layout&"'"
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "分类不存在!","Class.asp?Layout="&Layout,1
	Else
		ClassName=Rs("ClassName")
		ParentID=Rs("ParentID")
	End If
	Rs.Close
%>
<form name="myform" method="post" action="Class.asp?Layout=<%=Layout%>&Action=SaveMod">
<table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">修改分类</td>
  </tr>
  <tr class="t_rowh">
    <td colspan="2">说明：当所属分类是“请选择分类”时，则建立一级分类。</td>
  </tr>
  <tr class="t_row">
    <td width="40%">原分类名称</td>
    <td width="60%"><%=ClassName%></td>
  </tr>
  <tr class="t_row">
    <td>新分类名称</td>
    <td><input name="ClassName" type="text" id="ClassName" size="48"><input name="ClassID" type="hidden" value="<%=ClassID%>"></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="Submit" value="修改分类"></td>
  </tr>
</table>
</Form>
<%
End Sub

Sub ClassMove()
	Dim ClassName,ParentID
	ClassID=CheckChar(Trim(ClassID),"分类编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	If ClassID <> "" Then
		ClassID = Cint(ClassID)
	Else
		ClassID = 0
	End If
	Sql="Select ClassID,ClassName,ParentID From Class Where ClassID="&ClassID&" And Layout='"&Layout&"'"
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "分类不存在!","Class.asp?Layout="&Layout,1
	Else
		ClassName=Rs("ClassName")
		ParentID=Rs("ParentID")
	End If
	Rs.Close
%>
<form name="myform" method="post" action="Class.asp?Layout=<%=Layout%>&Action=SaveMove">
<table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">移动分类</td>
  </tr>
  <tr class="t_row">
    <td width="38%"><u>分类名称</u></td>
    <td width="62%"><%=ClassName%><input name="ClassID" type="hidden" value="<%=ClassID%>"></td>
  </tr>
  <tr class="t_row">
    <td><u>当前所属分类</u></td>
    <td><%=ShowClassName(ParentID)%></td>
  </tr>
  <tr class="t_row">
    <td><u>移动到</u><br>不能指定为当前分类的下属子分类<br>不能指定为外部分类<br><font color="#CC3300">当目标分类有内容时<br>目标分类的内容将移至当前分类中</font></td>
    <td><select name="ParentID" size="2" style="height:300px;width:300px;"><option value="0">无（作为一级分类）</option><%SelectClass(ParentID)%></select></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" name="Submit2" value="移动分类"></td>
  </tr>
</table>
</form>
<%
End Sub

Sub ClassUnite()
%>
<form name="myform" method="post" action="Class.asp?Layout=<%=Layout%>&Action=SaveUnite">
<table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">分类内容合并</td>
  </tr>
  <tr>
    <td colspan="2" class="t_rowh"><strong>注意事项：</strong><br>所有操作不可逆，请慎重操作！！！<br>不能在同一个分类内进行操作，不能将一个分类合并到其下属分类中。目标分类中不能含有子分类。<br>合并后您所指定的分类（或者包括其下属分类）将被删除，所有文章将转移到目标分类中。</td>
  </tr>
  <tr class="t_row">
    <td width="40%" height="60%">将分类：</td>
    <td><select name="ClassID" id="ClassID"><%SelectClass(0)%></select></td>
  </tr>
  <tr class="t_row">
    <td>合并到：</td>
    <td><select name="TargetClassID" id="TargetClassID"><%SelectClass(0)%></select></td>
  </tr>
  <tr class="t_row">
    <td>&nbsp;</td>
    <td><input type="submit" value="合并分类"></td>
  </tr>
</table>
</form>
<%
End Sub

Sub BatchMove()
%>
<table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <form name="myform" method="post" action="Class.asp?Layout=<%=Layout%>&Action=SaveBatchMove">
  <tr>
    <td colspan="2" class="t_title">按分类移动</td>
  </tr>
  <tr class="t_row">
    <td width="40%">原分类</td>
    <td width="60%"><select name="OldClassID" id="OldClassID"><%SelectClass(0)%></select>&nbsp;&nbsp;<input type="submit" value="提交"></td>
  </tr>
  <tr class="t_row">
    <td>目标分类</td>
    <td><select name="NewClassID" id="NewClassID"><%SelectClass(0)%></select></td>
  </tr>
  </form>
  <tr>
    <td colspan="2" class="t_title">按日期移动</td>
  </tr>
  <form name="myform" method="post" action="Class.asp?Layout=<%=Layout%>&Action=SaveBatchDateMove">
  <tr class="t_row">
    <td>移动多少天前的(填写数字)</td>
    <td><input name="DateWord" type="text" id="DateWord" value="0" size="38">&nbsp;&nbsp;<input type="submit" name="submit2" value="提交"></td>
  </tr>
  <tr class="t_row">
    <td>原分类</td>
    <td><select name="OldClassID" id="OldClassID"><%SelectClass(0)%></select></td>
  </tr>
  <tr class="t_row">
    <td>目标分类</td>
    <td><select name="NewClassID" id="NewClassID"><%SelectClass(0)%></select></td>
  </tr>
  </form>
</table>
<%
End Sub

Sub BatchDel()
%>
<form name="myform" method="post" action="Class.asp?Layout=<%=Layout%>&Action=SaveBatchDateDel">
<table border="1" align="center" cellpadding="0" cellspacing="1" class="t_border">
  <tr>
    <td colspan="2" class="t_title">删除分类指定日期前的内容</td>
  </tr>
  <tr class="t_row">
    <td width="40%">删除多少天前的(填写数字)</td>
    <td width="60%"><input name="DateWord" type="text" id="DateWord" value="0" size="38">&nbsp;&nbsp;<input type="submit" value="提交"></td>
  </tr>
  <tr class="t_row">
    <td>指定分类</td>
    <td><select name="ClassID" id="ClassID"><%SelectClass(0)%></select></td>
  </tr>
</table>
</form>
<%
End Sub

Sub SaveAdd()
	Dim ClassName,PrevOrderID,FoundErr,ParentID
	Dim Sql,Rs,tRs
	Dim RootID,ParentDepth,ParentPath,ParentStr,ParentName,MaxClassID,MaxRootID
	Dim PrevID,NextID,Child

	ClassName=CheckChar(Trim(Request.Form("ClassName")),"分类名称",50,"012")
	ParentID=CheckChar(Trim(Request.Form("ParentID")),"父分类ID",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout&"&Action=Add",1

	Set Rs = Conn.Execute("Select Max(ClassID) From Class")
	MaxClassID=Rs(0)
	If Isnull(MaxClassID) Then
		MaxClassID=0
	End If
	Rs.Close
	ClassID=MaxClassID+1
	Set Rs=Conn.Execute("Select Max(RootID) From Class Where Layout='"&Layout&"'")
	MaxRootID=Rs(0)
	If isnull(MaxRootID) Then
		MaxRootID=0
	End If
	Rs.Close
	RootID=MaxRootID+1
	
	If ParentID>0 Then
		Sql="Select * From Class where ClassID=" & ParentID & " And Layout='"&Layout&"'"
		Rs.Open Sql,Conn,1,1
		If Rs.Bof And Rs.Eof Then
			Msg "所属分类已经被删除！","Class.asp?Layout="&Layout&"&Action=Add",1
		End If
		RootID=Rs("RootID")
		ParentName=Rs("ClassName")
		ParentDepth=Rs("Depth")
		ParentPath=Rs("ParentPath")
		Child=Rs("Child")
		ParentPath=ParentPath & "," & ParentID     '得到此分类的父级分类路径
		PrevOrderID=Rs("OrderID")
		If Child>0 Then		
			dim RsPrevOrderID
			'得到与本分类同级的最后一个分类的OrderID
			Set RsPrevOrderID=Conn.Execute("Select Max(OrderID) From Class where ParentID=" & ParentID &" And Layout='"&Layout&"'")
			PrevOrderID=RsPrevOrderID(0)
			Set tRs=Conn.Execute("Select ClassID from Class where ParentID=" & ParentID & " And OrderID=" & PrevOrderID &" And Layout='"&Layout&"'")
			PrevID=tRs(0)
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			Set RsPrevOrderID=Conn.Execute("Select Max(OrderID) From Class where ParentPath like '%" & ParentPath & "%' And Layout='"&Layout&"'")
			If (Not(RsPrevOrderID.Bof And RsPrevOrderID.Eof)) Then
				If Not IsNull(RsPrevOrderID(0))  Then
			 		If RsPrevOrderID(0)>PrevOrderID Then
						PrevOrderID=RsPrevOrderID(0)
					End If
				End If
			End If
		Else
			PrevID=0
		End If
		Rs.Close
	Else
		If MaxRootID>0 Then
			Set tRs=Conn.Execute("Select ClassID from Class where RootID=" & MaxRootID & " And Depth=0 And Layout='"&Layout&"'")
			PrevID=tRs(0)
			tRs.Close
		Else
			PrevID=0
		End If
		PrevOrderID=0
		ParentPath="0"
	End If

	Sql="Select * From Class Where ParentID=" & ParentID & " And ClassName='" & ClassName & "' And Layout='"&Layout&"'"
	Set Rs=Server.CreateObject("adodb.recordSet")
	Rs.Open Sql,Conn,1,1
	If Not(Rs.Bof And Rs.Eof) Then
		Rs.Close
		Set Rs=Nothing
		If ParentID=0 Then
			Msg "已经存在一级分类：" & ClassName &" !","Class.asp?Layout="&Layout&"&Action=Add",1
		Else
			Msg """" & ParentName & """中已经存在子分类""" & ClassName & """!","Class.asp?Layout="&Layout&"&Action=Add",1
		End If
	End If
	Rs.Close

	Sql="Select * From Class Where Layout='"&Layout&"'"
	Rs.Open Sql,Conn,1,3
    Rs.AddNew
	Rs("ClassID")=ClassID
   	Rs("ClassName")=ClassName
	Rs("RootID")=RootID
	Rs("ParentID")=ParentID
	If ParentID>0 Then
		Rs("Depth")=ParentDepth+1
	Else
		Rs("Depth")=0
	End If
	Rs("ParentPath")=ParentPath
	Rs("OrderID")=PrevOrderID
	Rs("Child")=0
	Rs("PrevID")=PrevID
	Rs("NextID")=0
	Rs("Layout")=Layout
	Rs.Update
	Rs.Close
    Set Rs=Nothing
	'将父类的内容移动至此分类
	Conn.Execute("Update "&Layout&" Set ClassID="&ClassID&" Where ClassID="&ParentID)
	
	'更新与本分类同一父分类的上一个分类的“NextID”字段值
	If PrevID>0 Then
		Conn.Execute("Update Class Set NextID=" & ClassID & " where ClassID=" & PrevID &" And Layout='"&Layout&"'")
	End If
	
	If ParentID>0 Then
		'更新其父类的子分类数
		Conn.Execute("Update Class Set Child=Child+1 where ClassID="&ParentID &" And Layout='"&Layout&"'")
		
		'更新该分类排序以及大于本需要和同在本分类下的分类排序序号
		Conn.Execute("Update Class Set OrderID=OrderID+1 where RootID=" & rootid & " And OrderID>" & PrevOrderID &" And Layout='"&Layout&"'")
		Conn.Execute("Update Class Set OrderID=" & PrevOrderID & "+1 where ClassID=" & ClassID &" And Layout='"&Layout&"'")
	End If
	Msg "分类添加成功!","Class.asp?Layout="&Layout,0
End Sub

Sub SaveMod()
	Dim ClassName,rsClass

	ClassID=CheckChar(Trim(Request.Form("ClassID")),"分类编号",0,"03")
	ClassName=CheckChar(Trim(Request.Form("ClassName")),"分类名称",50,"012")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout&"&Action=Mod",1
	
	Sql="Select * From Class Where ClassID=" & ClassID & " And Layout='"&Layout&"'"
	Set rsClass=server.CreateObject ("Adodb.Recordset")
	rsClass.open sql,conn,1,3
	If rsClass.bof And rsClass.eof Then
		Msg "找不到指定的分类!","Class.asp?Layout="&Layout&"&Action=Mod&ClassID="&ClassID,1
	End If
	
   	rsClass("ClassName")=ClassName
	rsClass.Update
	rsClass.Close
	Set rsClass=Nothing

	Msg "分类修改修改成功!","Class.asp?Layout="&Layout,0
End Sub

Sub SaveMove()
	dim sql,rsClass,i
	dim rParentID
	dim trs,rs
	dim ParentID,RootID,Depth,Child,ParentPath,ParentName,iParentID,iParentPath,PrevOrderID,PrevID,NextID

	ClassID=CheckChar(Trim(Request.Form("ClassID")),"分类编号",0,"03")
	rParentID=CheckChar(Trim(Request.Form("ParentID")),"移动至分类编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout&"&Action=Move&ClassID="&ClassID,1
	
	sql="select * From Class where ClassID=" & ClassID & " And Layout='"&Layout&"'"
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		rsClass.close
		set rsClass=nothing
		Msg "找不到指定的分类!","Class.asp?Layout="&Layout&"&Action=Move&ClassID="&ClassID,1
	End If
	
	if rsClass("ParentID")<>rParentID then   '更改了所属分类，则要做一系列检查
		if rParentID=rsClass("ClassID") then
			Msg "所属分类不能为自己!","Class.asp?Layout="&Layout&"&Action=Move&ClassID="&ClassID,1
		end if
		'判断所指定的分类是否为外部分类或本分类的下属分类
		if rsClass("ParentID")=0 then
			if rParentID>0 then
				set trs=conn.execute("select rootid From Class where ClassID="&rParentID & " And Layout='"&Layout&"'")
				if trs.bof and trs.eof then
					Msg "不能指定外部分类为所属分类!","Class.asp?Layout="&Layout&"&Action=Move&ClassID="&ClassID,1
				else
					if rsClass("rootid")=trs(0) then
						Msg "不能指定该分类的下属分类作为所属分类!","Class.asp?Layout="&Layout&"&Action=Move&ClassID="&ClassID,1
					end if
				end if
				trs.close
				set trs=nothing
			end if
		else
			set trs=conn.execute("select ClassID From Class where ParentPath like '%"&rsClass("ParentPath")&"," & rsClass("ClassID") & "%' and ClassID="&rParentID & " And Layout='"&Layout&"'")
			if not (trs.eof and trs.bof) then
				Msg "您不能指定该分类的下属分类作为所属分类!","Class.asp?Layout="&Layout&"&Action=Move&ClassID="&ClassID,1
			end if
			trs.close
			set trs=nothing
		end if
		
	end if
	
	if rsClass("ParentID")=0 then
		ParentID=rsClass("ClassID")
		iParentID=0
	else
		ParentID=rsClass("ParentID")
		iParentID=rsClass("ParentID")
	end if
	Depth=rsClass("Depth")
	Child=rsClass("Child")
	RootID=rsClass("RootID")
	ParentPath=rsClass("ParentPath")
	PrevID=rsClass("PrevID")
	NextID=rsClass("NextID")
	rsClass.close
	set rsClass=nothing
	
	
  '假如更改了所属分类
  '需要更新其原来所属分类信息，包括深度、父级ID、分类数、排序、继承版主等数据
  '需要更新当前所属分类信息
  '继承版主数据需要另写函数进行更新--取消，在前台可用ClassID in ParentPath来获得
  dim mrs,MaxRootID
  set mrs=conn.execute("select max(rootid) From Class Where Layout='"&Layout&"'")
  MaxRootID=mrs(0)
  set mrs=nothing
  if isnull(MaxRootID) then
	MaxRootID=0
  end if
  dim k,nParentPath,mParentPath
  dim ParentSql,ClassCount
  dim rsPrevOrderID
  if clng(parentid)<>rParentID and not (iParentID=0 and rParentID=0) then  '假如更改了所属分类
	'更新原来同一父分类的上一个分类的NextID和下一个分类的PrevID
	if PrevID>0 then
		conn.execute "update Class set NextID=" & NextID & " where ClassID=" & PrevID & " And Layout='"&Layout&"'"
	end if
	if NextID>0 then
		conn.execute "update Class set PrevID=" & PrevID & " where ClassID=" & NextID & " And Layout='"&Layout&"'"
	end if
	
	if iParentID>0 and rParentID=0 then  	'如果原来不是一级分类改成一级分类
		'得到上一个一级分类分类
		sql="select ClassID,NextID from Class where RootID=" & MaxRootID & " and Depth=0 And Layout='"&Layout&"'"
		set rs=server.CreateObject("Adodb.recordset")
		rs.open sql,conn,1,3
		PrevID=rs(0)      '得到新的PrevID
		rs(1)=ClassID     '更新上一个一级分类分类的NextID的值
		rs.update
		rs.close
		set rs=nothing
		
		MaxRootID=MaxRootID+1
		'更新当前分类数据
		conn.execute("update Class set depth=0,OrderID=0,rootid="&maxrootid&",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0 where ClassID="&ClassID & " And Layout='"&Layout&"'")
		'如果有下属分类，则更新其下属分类数据。下属分类的排序不需考虑，只需更新下属分类深度和一级排序ID(rootid)数据
		if child>0 then
			i=0
			ParentPath=ParentPath & ","
			set rs=conn.execute("select * From Class where ParentPath like '%"&ParentPath & ClassID&"%' And Layout='"&Layout&"'")
			do while not rs.eof
				i=i+1
				mParentPath=replace(rs("ParentPath"),ParentPath,"")
				conn.execute("update Class set depth=depth-"&depth&",rootid="&maxrootid&",ParentPath='"&mParentPath&"' where ClassID="&rs("ClassID") & " And Layout='"&Layout&"'")
				rs.movenext
			loop
			rs.close
			set rs=nothing
		end if
		

		'更新其原来所属分类的分类数，排序相当于剪枝而不需考虑
		conn.execute("update Class set child=child-1 where ClassID="&iParentID & " And Layout='"&Layout&"'")
		
	elseif iParentID>0 and rParentID>0 then    '如果是将一个分分类移动到其他分分类下
		'得到当前分类的下属子分类数
		ParentPath=ParentPath & ","
		set rs=conn.execute("select count(*) From Class where ParentPath like '%"&ParentPath & ClassID&"%' And Layout='"&Layout&"'")
		ClassCount=rs(0)
		if isnull(ClassCount) then
			ClassCount=1
		end if
		rs.close
		set rs=nothing
		
		'获得目标分类的相关信息		
		set trs=conn.execute("select * From Class where ClassID="&rParentID & " And Layout='"&Layout&"'")
		if trs("Child")>0 then		
			'得到与本分类同级的最后一个分类的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From Class where ParentID=" & trs("ClassID") & " And Layout='"&Layout&"'")
			PrevOrderID=rsPrevOrderID(0)
			'得到与本分类同级的最后一个分类的ClassID
			sql="select ClassID,NextID from Class where ParentID=" & trs("ClassID") & " and OrderID=" & PrevOrderID & " And Layout='"&Layout&"'"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)    '得到新的PrevID
			rs(1)=ClassID     '更新上一个分类的NextID的值
			rs.update
			rs.close
			set rs=nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From Class where ParentPath like '%" & trs("ParentPath") & "," & trs("ClassID") & "%' And Layout='"&Layout&"'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
		
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.execute("update Class set OrderID=OrderID+" & ClassCount & "+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID & " And Layout='"&Layout&"'")
		
		'更新当前分类数据
		conn.execute("update Class set depth="&trs("depth")&"+1,OrderID="&PrevOrderID&"+1,rootid="&trs("rootid")&",ParentID="&rParentID&",ParentPath='" & trs("ParentPath") & "," & trs("ClassID") & "',PrevID=" & PrevID & ",NextID=0 where ClassID="&ClassID & " And Layout='"&Layout&"'")
		
		'如果有子分类则更新子分类数据，深度为原来的相对深度加上当前所属分类的深度
		set rs=conn.execute("select * From Class where ParentPath like '%"&ParentPath&ClassID&"%' And Layout='"&Layout&"' order by OrderID")
		i=1
		do while not rs.eof
			i=i+1
			iParentPath=trs("ParentPath") & "," & trs("ClassID") & "," & replace(rs("ParentPath"),ParentPath,"")
			conn.execute("update Class set depth=depth-"&depth&"+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&iParentPath&"' where ClassID="&rs("ClassID") & " And Layout='"&Layout&"'")
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		
		'更新所指向的上级分类的子分类数
		conn.execute("update Class set child=child+1 where ClassID="&rParentID & " And Layout='"&Layout&"'")
		
		'更新其原父类的子分类数			
		conn.execute("update Class set child=child-1 where ClassID="&iParentID & " And Layout='"&Layout&"'")
	else    '如果原来是一级分类改成其他分类的下属分类
		'得到移动的分类总数
		set rs=conn.execute("select count(*) From Class where rootid="&rootid & " And Layout='"&Layout&"'")
		ClassCount=rs(0)
		rs.close
		set rs=nothing
		
		'获得目标分类的相关信息		
		set trs=conn.execute("select * From Class where ClassID="&rParentID & " And Layout='"&Layout&"'")
		if trs("Child")>0 then		
			'得到与本分类同级的最后一个分类的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From Class where ParentID=" & trs("ClassID") & " And Layout='"&Layout&"'")
			PrevOrderID=rsPrevOrderID(0)
			sql="select ClassID,NextID from Class where ParentID=" & trs("ClassID") & " and OrderID=" & PrevOrderID & " And Layout='"&Layout&"'"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)
			rs(1)=ClassID
			rs.update
			set rs=nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From Class where ParentPath like '%" & trs("ParentPath") & "," & trs("ClassID") & "%' And Layout='"&Layout&"'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
	
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.execute("update Class set OrderID=OrderID+" & ClassCount &"+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID & " And Layout='"&Layout&"'")
		
		conn.execute("update Class set PrevID=" & PrevID & ",NextID=0 where ClassID=" & ClassID & " And Layout='"&Layout&"'")
		set rs=conn.execute("select * From Class where rootid="&rootid&" And Layout='"&Layout&"' order by OrderID")
		i=0
		do while not rs.eof
			i=i+1
			if rs("parentid")=0 then
				ParentPath=trs("ParentPath") & "," & trs("ClassID")
				conn.execute("update Class set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"',parentid="&rParentID&" where ClassID="&rs("ClassID") & " And Layout='"&Layout&"'")
			else
				ParentPath=trs("ParentPath") & "," & trs("ClassID") & "," & replace(rs("ParentPath"),"0,","")
				conn.execute("update Class set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"' where ClassID="&rs("ClassID") & " And Layout='"&Layout&"'")
			end if
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		'更新所指向的上级分类分类数		
		Conn.Execute("Update Class Set child=child+1 Where ClassID="&rParentID & " And Layout='"&Layout&"'")
	end if
  end if
  If rParentID>0 Then
	  Conn.Execute("Update "&Layout&" Set ClassID="&ClassID&" Where ClassID="&rParentID)
  End If
  Msg "分类移动成功!","Class.asp?Layout="&Layout,0
End Sub

Sub ClassClear()
	dim strClassID,rs,trs,SuccessMsg
	ClassID=CheckChar(Trim(ClassID),"分类编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1

	strClassID=cstr(ClassID)
	set rs=conn.execute("select ClassID,Child,ParentPath from Class where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	if rs.bof and rs.eof then
		Msg "分类不存在，或者已经被删除!","Class.asp?Layout="&Layout,1
	end if
	if rs(1)>0 then
		set trs=conn.execute("select ClassID from Class where ParentID=" & rs(0) & " And Layout='"&Layout&"'")
		do while not trs.eof
			strClassID=strClassID & "," & trs(0)
			trs.movenext
		loop
		trs.close
		set trs=conn.execute("select ClassID from Class where ParentPath like '%" & rs(2) & "," & rs(0) & "%' And Layout='"&Layout&"'")
		do while not trs.eof
			strClassID=strClassID & "," & trs(0)
			trs.movenext
		loop
		trs.close
		set trs=nothing
	end if
	rs.close
	set rs=nothing
	conn.execute("delete from "&Layout&" where ClassID in (" & strClassID & ")")	
	Msg "此分类（包括子分类）的所有内容已经被删除!","Class.asp?Layout="&Layout,0
End Sub

Sub ClassDel()
	Dim sql,rs,PrevID,NextID
	ClassID=CheckChar(Trim(ClassID),"分类编号",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	
	sql="select ClassID,RootID,Depth,ParentID,Child,PrevID,NextID From Class where ClassID="&ClassID & " And Layout='"&Layout&"'"
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		Msg "分类不存在,或者已经被删除!","Class.asp?Layout="&Layout,1
	else
		if rs("Child")>0 then
			Msg "该分类含有子分类,请删除其子分类后再进行删除本分类的操作!","Class.asp?Layout="&Layout,1
		end if
	end if
	PrevID=rs("PrevID")
	NextID=rs("NextID")
	if rs("Depth")>0 then
		conn.execute("update Class set Child=Child-1 where ClassID=" & rs("ParentID") & " And Layout='"&Layout&"'")
	end if
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	'删除本分类的所有内容
	conn.execute("delete from "&Layout&" where ClassID=" & ClassID)
	
	'修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update Class set NextID=" & NextID & " where ClassID=" & PrevID & " And Layout='"&Layout&"'"
	end if
	if NextID>0 then
		conn.execute "update Class set PrevID=" & PrevID & " where ClassID=" & NextID & " And Layout='"&Layout&"'"
	end if
'	Msg "分类已删除!","Class.asp?Layout="&Layout,0
	Response.Redirect "Class.asp?Layout="&Layout
End Sub

Sub ClassUp()
	Dim sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	ClassID=CheckChar(Trim(ClassID),"分类ID",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	MoveNum=1
	'得到本分类的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID,RootID from Class where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	PrevID=rs(0)
	NextID=rs(1)
	cRootID=rs(2)
	rs.close
	set rs=nothing
	If PrevID=0 Then Msg "当前分类在同级分类中已经在最上面!","Class.asp?Layout="&Layout,1
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update Class set NextID=" & NextID & " where ClassID=" & PrevID & " And Layout='"&Layout&"'"
	end if
	if NextID>0 then
		conn.execute "update Class set PrevID=" & PrevID & " where ClassID=" & NextID & " And Layout='"&Layout&"'"
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From Class Where Layout='"&Layout&"'")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.execute("update Class set RootID=" & MaxRootID & " where RootID=" & cRootID) & " And Layout='"&Layout&"'"
	
	'然后将位于当前分类以上的分类的RootID依次加一，范围为要提升的数字
	sqlOrder="select * From Class where ParentID=0 and RootID<" & cRootID & " And Layout='"&Layout&"' order by RootID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前分类已经在最上面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		conn.execute("update Class set RootID=RootID+1 where RootID=" & tRootID & " And Layout='"&Layout&"'")
		i=i+1
		if i>MoveNum then
			rsOrder("PrevID")=ClassID
			rsOrder.update
			conn.execute("update Class set NextID=" & rsOrder("ClassID") & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update Class set PrevID=0 where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	else
		rsOrder("NextID")=ClassID
		rsOrder.update
		conn.execute("update Class set PrevID=" & rsOrder("ClassID") & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.execute("update Class set RootID=" & tRootID & " where RootID=" & MaxRootID & " And Layout='"&Layout&"'")
'	Msg "分类上移成功!","Class.asp?Layout="&Layout,0
	Response.Redirect "Class.asp?Layout="&Layout
End Sub

Sub ClassDown()
	dim sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	ClassID=CheckChar(Trim(ClassID),"分类ID",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	MoveNum=1
	'得到本分类的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID,RootID from Class where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	PrevID=rs(0)
	NextID=rs(1)
	cRootID=rs(2)
	rs.close
	set rs=nothing
	If NextID=0 Then Msg "当前分类在同级分类中已经在最下面!","Class.asp?Layout="&Layout,1
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update Class set NextID=" & NextID & " where ClassID=" & PrevID & " And Layout='"&Layout&"'"
	end if
	if NextID>0 then
		conn.execute "update Class set PrevID=" & PrevID & " where ClassID=" & NextID & " And Layout='"&Layout&"'"
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From Class Where Layout='"&Layout&"'")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.execute("update Class set RootID=" & MaxRootID & " where RootID=" & cRootID & " And Layout='"&Layout&"'")
	
	'然后将位于当前分类以下的分类的RootID依次减一，范围为要下降的数字
	sqlOrder="select * From Class where ParentID=0 and RootID>" & cRootID & " And Layout='"&Layout&"' order by RootID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前分类已经在最下面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		conn.execute("update Class set RootID=RootID-1 where RootID=" & tRootID & " And Layout='"&Layout&"'")
		i=i+1
		if i>MoveNum then
			rsOrder("NextID")=ClassID
			rsOrder.update
			conn.execute("update Class set PrevID=" & rsOrder("ClassID") & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")
			exit do
		end if
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update Class set NextID=0 where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	else
		rsOrder("PrevID")=ClassID
		rsOrder.update
		conn.execute("update Class set NextID=" & rsOrder("ClassID") & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.execute("update Class set RootID=" & tRootID & " where RootID=" & MaxRootID & " And Layout='"&Layout&"'")
'	Msg "分类下移成功!","Class.asp?Layout="&Layout,0
	Response.Redirect "Class.asp?Layout="&Layout
End Sub

Sub ClassUpN()
	dim sqlOrder,rsOrder,MoveNum,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	ClassID=CheckChar(Trim(ClassID),"分类ID",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	MoveNum=1

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From Class where ClassID="&ClassID & " And Layout='"&Layout&"'")
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & ClassID
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing
	If PrevID=0 Then Msg "当前分类在同级分类中已经在最上面!","Class.asp?Layout="&Layout,1
	if child>0 then
		set rs=conn.execute("select count(*) From Class where ParentPath like '%"&ParentPath&"%' And Layout='"&Layout&"'")
		oldorders=rs(0)
		rs.close
		set rs=nothing
	else
		oldorders=0
	end if
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update Class set NextID=" & NextID & " where ClassID=" & PrevID & " And Layout='"&Layout&"'"
	end if
	if NextID>0 then
		conn.execute "update Class set PrevID=" & PrevID & " where ClassID=" & NextID & " And Layout='"&Layout&"'"
	end if
	
	'和该分类同级且排序在其之上的分类------更新其排序，范围为要提升的数字
	sql="select ClassID,OrderID,child,ParentPath,PrevID,NextID From Class where ParentID="&ParentID&" and OrderID<"&OrderID&" And Layout='"&Layout&"' order by OrderID desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=1
	do while not rs.eof
		tOrderID=rs(1)
		conn.execute("update Class set OrderID="&tOrderID+oldorders+i&" where ClassID="&rs(0) & " And Layout='"&Layout&"'")
		if rs(2)>0 then
			ii=i+1
			set trs=conn.execute("select ClassID,OrderID From Class where ParentPath like '%"&rs(3)&","&rs(0)&"%' And Layout='"&Layout&"' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					conn.execute("update Class set OrderID="&tOrderID+oldorders+ii&" where ClassID="&trs(0) & " And Layout='"&Layout&"'")
					ii=ii+1
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		i=i+1
		if i>MoveNum then
			rs(4)=ClassID
			rs.update
			conn.execute("update Class set NextID=" & rs(0) & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")		
			exit do
		end if
		rs.movenext
		loop
	rs.movenext
	if rs.eof then
		conn.execute("update Class set PrevID=0 where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	else
		rs(5)=ClassID
		rs.update
		conn.execute("update Class set PrevID=" & rs(0) & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的分类的序号
	conn.execute("update Class set OrderID="&tOrderID&" where ClassID="&ClassID & " And Layout='"&Layout&"'")
	'如果有下属分类，则更新其下属分类排序
	if child>0 then
		i=1
		set rs=conn.execute("select ClassID From Class where ParentPath like '%"&ParentPath&"%' And Layout='"&Layout&"' order by OrderID")
		do while not rs.eof
			conn.execute("update Class set OrderID="&tOrderID+i&" where ClassID="&rs(0) & " And Layout='"&Layout&"'")
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
'	Msg "分类上移成功!","Class.asp?Layout="&Layout,0
	Response.Redirect "Class.asp?Layout="&Layout
End Sub

Sub ClassDownN()
	dim sqlOrder,rsOrder,MoveNum,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	ClassID=CheckChar(Trim(ClassID),"分类ID",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	MoveNum=1

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From Class where ClassID="&ClassID & " And Layout='"&Layout&"'")
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & ClassID
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing
	If NextID=0 Then Msg "当前分类在同级分类中已经在最下面!","Class.asp?Layout="&Layout,1
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update Class set NextID=" & NextID & " where ClassID=" & PrevID & " And Layout='"&Layout&"'"
	end if
	if NextID>0 then
		conn.execute "update Class set PrevID=" & PrevID & " where ClassID=" & NextID & " And Layout='"&Layout&"'"
	end if
	
	'和该分类同级且排序在其之下的分类------更新其排序，范围为要下降的数字
	sql="select ClassID,OrderID,child,ParentPath,PrevID,NextID From Class where ParentID="&ParentID&" and OrderID>"&OrderID&" And Layout='"&Layout&"' order by OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=0      '同级分类
	ii=0     '同级分类和子分类
	do while not rs.eof
		conn.execute("update Class set OrderID="&OrderID+ii&" where ClassID="&rs(0) & " And Layout='"&Layout&"'")
		if rs(2)>0 then
			set trs=conn.execute("select ClassID,OrderID From Class where ParentPath like '%"&rs(3)&","&rs(0)&"%' And Layout='"&Layout&"' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					ii=ii+1
					conn.execute("update Class set OrderID="&OrderID+ii&" where ClassID="&trs(0) & " And Layout='"&Layout&"'")
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		ii=ii+1
		i=i+1
		if i>=MoveNum then
			rs(5)=ClassID
			rs.update
			conn.execute("update Class set PrevID=" & rs(0) & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")		
			exit do
		end if
		rs.movenext
	loop
	rs.movenext
	if rs.eof then
		conn.execute("update Class set NextID=0 where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	else
		rs(4)=ClassID
		rs.update
		conn.execute("update Class set NextID=" & rs(0) & " where ClassID=" & ClassID & " And Layout='"&Layout&"'")
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的分类的序号
	conn.execute("update Class set OrderID="&OrderID+ii&" where ClassID="&ClassID & " And Layout='"&Layout&"'")
	'如果有下属分类，则更新其下属分类排序
	if child>0 then
		i=1
		set rs=conn.execute("select ClassID From Class where ParentPath like '%"&ParentPath&"%' And Layout='"&Layout&"' order by OrderID")
		do while not rs.eof
			conn.execute("update Class set OrderID="&OrderID+ii+i&" where ClassID="&rs(0) & " And Layout='"&Layout&"'")
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
'	Msg "分类下移成功!","Class.asp?Layout="&Layout,0
	Response.Redirect "Class.asp?Layout="&Layout
End Sub

Sub ClassDelAll()
	Sql="Select ClassID From Class Where Layout='"&Layout&"'"
	Rs.Open Sql,Conn,1,3
	i=0
	If Rs.Eof Then
		Msg "没有任何分类可以删除的!","Class.asp?Action='"&Layout&"'",1
	Else
		Do While Not Rs.Eof
		Conn.Execute("Delete From "&Layout&" Where ClassID="& Rs("ClassID") &"")
		Rs.Delete
		i=i+1
		Rs.MoveNext
		Loop
	End If
	Rs.Close
	Msg "共删除个"&i&"分类!","Class.asp?Layout="&Layout,0
'	Response.Redirect "Class.asp?Layout="&Layout
End Sub

Sub SaveUnite()
	dim TargetClassID,ParentPath,iParentPath,Depth,iParentID,Child,PrevID,NextID
	dim rs,trs,i,SuccessMsg

	ClassID=CheckChar(Trim(Request.Form("ClassID")),"分类ID",0,"03")
	TargetClassID=CheckChar(Trim(Request.Form("TargetClassID")),"合并到的分类ID",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout,1
	
	If ClassID=TargetClassID Then Msg "请不要在相同分类内进行操作!","Class.asp?Layout="&Layout&"&Action=Unite",1
	'判断目标分类是否有子分类，如果有，则报错。
	set rs=conn.execute("select Child from Class where ClassID=" & TargetClassID & " And Layout='"&Layout&"'")
	if rs.bof and rs.eof then
		Msg "目标分类不存在，可能已经被删除!","Class.asp?Layout="&Layout&"&Action=Unite",1
	else
		if rs(0)>0 then
			Msg "目标分类中含有子分类，不能合并!","Class.asp?Layout="&Layout&"&Action=Unite",1
		end if
	end if
	'得到当前分类信息
	set rs=conn.execute("select ClassID,ParentID,ParentPath,PrevID,NextID,Depth from Class where ClassID="&ClassID & " And Layout='"&Layout&"'")
	iParentID=rs(1)
	Depth=rs(5)
	if iParentID=0 then
		ParentPath=rs(0)
	else
		ParentPath=rs(2) & "," & rs(0)
	end if
	iParentPath=rs(0)
	PrevID=rs(3)
	NextID=rs(4)
	
	'判断是否是合并到其下属分类中
	set rs=conn.execute("select ClassID from Class where ClassID="&TargetClassID&" and ParentPath like '%"&ParentPath&"%' And Layout='"&Layout&"'")
	if not (rs.eof and rs.bof) then
		Msg "不能将一个分类合并到其下属子分类中!","Class.asp?Layout="&Layout&"&Action=Unite",1
	end if

	'得到当前分类的下属分类ID
	set rs=conn.execute("select ClassID from Class where ParentPath like '%"&ParentPath&"%' And Layout='"&Layout&"'")
	i=0
	if not (rs.eof and rs.bof) then
		do while not rs.eof
			iParentPath=iParentPath & "," & rs(0)
			i=i+1
			rs.movenext
		loop
	end if
	if i>0 then
		ParentPath=iParentPath
	else
		ParentPath=ClassID
	end if
	
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update Class set NextID=" & NextID & " where ClassID=" & PrevID & " And Layout='"&Layout&"'"
	end if
	if NextID>0 then
		conn.execute "update Class set PrevID=" & PrevID & " where ClassID=" & NextID & " And Layout='"&Layout&"'"
	end if
	
	'更新文章及评论所属分类
	conn.execute("update "&Layout&" set ClassID="&TargetClassID&" where ClassID in ("&ParentPath&")")
	
	'删除被合并分类及其下属分类
	conn.execute("delete from Class where ClassID in ("&ParentPath&") And Layout='"&Layout&"'")
	
	'更新其原来所属分类的子分类数，排序相当于剪枝而不需考虑
	if Depth>0 then
		conn.execute("update Class set Child=Child-1 where ClassID="&iParentID & " And Layout='"&Layout&"'")
	end if
	
	Msg "分类合并成功,同时删除了被合并的分类及其子分类!","Class.asp?Layout="&Layout,0
End Sub

Sub SaveBatchMove()
	Dim OldClassID,NewClassID
	OldClassID=CheckChar(Trim(Request.Form("OldClassID")),"原分类",0,"03")
	NewClassID=CheckChar(Trim(Request.Form("NewClassID")),"目标分类",0,"03")
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout&"&Action=BatchMove",1
	Sql="Select Child From Class Where ClassID="&OldClassID
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "原分类不存在!","Class.asp?Layout="&Layout,1
	Else
		If Rs("Child")> 0 Then Msg "原分类含有子目录!","Class.asp?Layout="&Layout&"&Action=BatchMove",1
	End If
	Rs.Close
	Sql="Select Child From Class Where ClassID="&NewClassID
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "目标分类不存在!","Class.asp?Layout="&Layout,1
	Else
		If Rs("Child")> 0 Then Msg "目标分类含有子目录!","Class.asp?Layout="&Layout&"&Action=BatchMove",1
	End If
	Rs.Close
	Conn.Execute("Update "&Layout&" Set ClassID="&NewClassID&" Where ClassID="&OldClassID)
	Msg "批量移动成功!",Layout&".asp",0
End Sub

Sub SaveBatchDateMove()
	Dim OldClassID,NewClassID,DateWord
	DateWord=CheckChar(Trim(Request.Form("DateWord")),"转移日期",0,"03")
	OldClassID=CheckChar(Trim(Request.Form("OldClassID")),"原分类",0,"03")
	NewClassID=CheckChar(Trim(Request.Form("NewClassID")),"目标分类",0,"03")
	If DateWord < 1 Then ErrContent = ErrContent & "零天就不要转移了!"
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout&"&Action=BatchMove",1
	Sql="Select Child From Class Where ClassID="&OldClassID

	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "原分类不存在!","Class.asp?Layout="&Layout,1
	Else
		If Rs("Child")> 0 Then Msg "原分类含有子目录!","Class.asp?Layout="&Layout&"&Action=BatchMove",1
	End If
	Rs.Close
	Sql="Select Child From Class Where ClassID="&NewClassID
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "目标分类不存在!","Class.asp?Layout="&Layout,1
	Else
		If Rs("Child")> 0 Then Msg "目标分类含有子目录!","Class.asp?Layout="&Layout&"&Action=BatchMove",1
	End If
	Rs.Close
	If IsSqlDataBase = 1 Then
		Conn.Execute("Update "&Layout&" Set ClassID="&NewClassID&" Where ClassID="&OldClassID&" And datediff(d,DateAndTime,"&SqlNow&")>"&Cint(DateWord))
	Else
		Conn.Execute("Update "&Layout&" Set ClassID="&NewClassID&" Where ClassID="&OldClassID&" And datediff('d',DateAndTime,"&SqlNow&")>"&Cint(DateWord))
	End If
	Msg "批量移动成功!",Layout&".asp",0
End Sub

Sub SaveBatchDateDel()
	Dim ClassID,DateWord
	DateWord=CheckChar(Trim(Request.Form("DateWord")),"删除日期",0,"03")
	ClassID=CheckChar(Trim(Request.Form("ClassID")),"分类",0,"03")
	If DateWord < 1 Then ErrContent = ErrContent & "零天就不要删除了!"
	If ErrContent<>"" Then Msg ErrContent,"Class.asp?Layout="&Layout&"&Action=BatchDel",1
	Sql="Select Child From Class Where ClassID="&ClassID
	Rs.Open Sql,Conn,1,1
	If Rs.Eof Then
		Msg "分类不存在!","Class.asp?Layout="&Layout,1
	Else
		If Rs("Child")> 0 Then Msg "分类含有子目录!","Class.asp?Layout="&Layout&"&Action=BatchDel",1
	End If
	Rs.Close
	If IsSqlDataBase = 1 Then
		Conn.Execute("Delete From "&Layout&" Where ClassID="&ClassID&" And datediff(d,DateAndTime,"&SqlNow&")>"&Cint(DateWord))
	Else
		Conn.Execute("Delete From "&Layout&" Where ClassID="&ClassID&" And datediff('d',DateAndTime,"&SqlNow&")>"&Cint(DateWord))
	End If
	Msg "批量删除成功!",Layout&".asp",0
End Sub

Bottom()
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
