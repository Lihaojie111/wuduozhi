<%
'================��ҳ��ʼ��================
Dim CurPages,PageRs
Sub Page_Init(PageRs)
PageRs.PageSize=Pagenum
CurPages=Trim(Request.QueryString("Page"))	
If Not(Isnumeric(CurPages)) Then CurPages=1
If Int(CurPages)>Int(PageRs.PageCount) or Int(CurPages)<1 Then
CurPages=1
Else
CurPages=Int(CurPages)
End If
PageRs.AbsolutePage=CurPages
End Sub	
%>
<%
'==============��ҳ�б�=================
public sub Pagelist(PageRs) '��ʾ��ҳ
If PageUrl<>"" Then
PageUrl=PageUrl&""
Else
PageUrl=""
End If
select case pagestyle
case 1 '��ҳ��ʽ1
Response.Write"��<Font color='Red'>"&PageRs.RecordCount&"</Font>��&nbsp;"
Response.Write"<Font color='Red'>"&CurPages&"</Font>/"&PageRs.PageCount&"&nbsp;"
If CurPages<=1 Then
Response.Write "��ҳ&nbsp;��һҳ&nbsp;"
Else
Response.Write"<a href='"&PageUrl&"Page=1'>��ҳ</a>&nbsp;"
Response.Write"<a href='"&PageUrl&"Page="&(CurPages-1)&"'>��һҳ</a>&nbsp;"
End If
If CurPages>=PageRs.PageCount Then
Response.Write"��һҳ&nbsp;βҳ&nbsp;"
Else
Response.Write"<a href='"&PageUrl&"Page="&(CurPages+1)&"'>��һҳ</a>&nbsp;"
Response.Write"<a href='"&PageUrl&"Page="&PageRs.PageCount&"'>βҳ</a>&nbsp;"
End If
Response.Write"ת����<input type='text' name='page' size='3' maxlength='5' value='" &CurPages & "' onKeyPress=""if (event.keyCode==13) window.location='" & PageUrl & "page=" & "'+this.value;""'>ҳ"

case 2 ' ��ҳ��ʽ2
response.write "��"&PageRs.RecordCount&"��&nbsp;"&CurPages&"/"&PageRs.PageCount&"&nbsp;"
if PageRs.pagecount<PageListNum then
PageListNum=PageRs.PageCount
end if
if CurPages="" or CurPages=1 or CurPages=2 or CurPages=3 then
PageStart=1
else
pageStart=CurPages-1
end if
pageEnd=pageStart + PageListNum-1
if pageEnd > PageRs.pagecount then
pageEnd = PageRs.pagecount
end if
response.write "<a href='"&PageUrl&"Page=1' class='hui1' >��һҳ</a> "
if CurPages=1 or CurPages=2 then
response.write "<a href='"&PageUrl&"Page=1' class='hui1'>��һҳ</a> "
else
response.write "<a href='"&PageUrl&"Page="&(CurPages-1)&"' class='hui1'>��һҳ</a> "
end if
for tempFor=PageStart to PageEnd
response.write "[<a href='"&PageUrl&"Page="&tempFor&"'>"
if tempFor=CurPages then
response.write "<font color=red><strong>"&tempFor&"</strong></font>"
else
response.write tempFor
end if
response.write "</a>] "
next
if Page=PageRs.pagecount then
response.write "<a href='"&PageUrl&"Page="&(CurPages)&"' class='hui1'>��һҳ</a> "
else
response.write "<a href='"&PageUrl&"Page="&(CurPages+1)&"'>>></a> "
end if
if CurPages=PageRs.pageCount then
response.write "<a href='"&PageUrl&"Page="&CurPages&"' class='hui1'>βҳ</a> "
else
response.write "<a href='"&PageUrl&"Page="&PageRs.PageCount&"'>>>|</a> "
end if
Response.Write"&nbsp;ת����<input type='text' name='page' size='3' maxlength='5' value='" &CurPages & "' onKeyPress=""if (event.keyCode==13) window.location='" & PageUrl & "page=" & "'+this.value;""'>ҳ"

case 3 ' ��ҳ��ʽ3
Response.Write"[��<Font color='Red'><b>"&PageRs.RecordCount&"</b></Font>����¼]"
Response.Write"[ҳ��<Font color='Red'>"&CurPages&"</Font>/"&PageRs.PageCount&"ҳ]"
If CurPages<=1 Then
Response.Write "[��ҳ][��һҳ] "
Else
Response.Write"[<a href='"&PageUrl&"Page=1'>��ҳ</a>&nbsp;]"
Response.Write"[<a href='"&PageUrl&"Page="&(CurPages-1)&"'>��һҳ</a>]"
End If
If CurPages>=PageRs.PageCount Then
Response.Write"[��һҳ][βҳ]"
Else
Response.Write"[<a href='"&PageUrl&"Page="&(CurPages+1)&"'>��һҳ</a>]"
Response.Write"[<a href='"&PageUrl&"Page="&PageRs.PageCount&"'>βҳ</a>]&nbsp;"
End If
Response.Write"ת����<select name='Page' size='1' onchange=""javascript:window.location='" & PageUrl & "Page=" & "'+this.options[this.selectedIndex].value;"">"
For j = 1 To PageRs.Pagecount
If j=CurPages Then
Response.Write"<option selected value="&j&">"&"��"&j&"ҳ"&"</option>"
Else
Response.Write"<option value="&j&">"&"��"&j&"ҳ"&"</option>"
End If
Next 
Response.Write"</select>" 
end select
End Sub
%>
