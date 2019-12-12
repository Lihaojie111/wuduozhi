<%
'================分页初始化================
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
'==============分页列表=================
public sub Pagelist(PageRs) '显示分页
If PageUrl<>"" Then
PageUrl=PageUrl&""
Else
PageUrl=""
End If
select case pagestyle
case 1 '分页样式1
Response.Write"共<Font color='Red'>"&PageRs.RecordCount&"</Font>条&nbsp;"
Response.Write"<Font color='Red'>"&CurPages&"</Font>/"&PageRs.PageCount&"&nbsp;"
If CurPages<=1 Then
Response.Write "首页&nbsp;上一页&nbsp;"
Else
Response.Write"<a href='"&PageUrl&"Page=1'>首页</a>&nbsp;"
Response.Write"<a href='"&PageUrl&"Page="&(CurPages-1)&"'>上一页</a>&nbsp;"
End If
If CurPages>=PageRs.PageCount Then
Response.Write"下一页&nbsp;尾页&nbsp;"
Else
Response.Write"<a href='"&PageUrl&"Page="&(CurPages+1)&"'>下一页</a>&nbsp;"
Response.Write"<a href='"&PageUrl&"Page="&PageRs.PageCount&"'>尾页</a>&nbsp;"
End If
Response.Write"转到第<input type='text' name='page' size='3' maxlength='5' value='" &CurPages & "' onKeyPress=""if (event.keyCode==13) window.location='" & PageUrl & "page=" & "'+this.value;""'>页"

case 2 ' 分页样式2
response.write "共"&PageRs.RecordCount&"条&nbsp;"&CurPages&"/"&PageRs.PageCount&"&nbsp;"
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
response.write "<a href='"&PageUrl&"Page=1' class='hui1' >第一页</a> "
if CurPages=1 or CurPages=2 then
response.write "<a href='"&PageUrl&"Page=1' class='hui1'>上一页</a> "
else
response.write "<a href='"&PageUrl&"Page="&(CurPages-1)&"' class='hui1'>上一页</a> "
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
response.write "<a href='"&PageUrl&"Page="&(CurPages)&"' class='hui1'>下一页</a> "
else
response.write "<a href='"&PageUrl&"Page="&(CurPages+1)&"'>>></a> "
end if
if CurPages=PageRs.pageCount then
response.write "<a href='"&PageUrl&"Page="&CurPages&"' class='hui1'>尾页</a> "
else
response.write "<a href='"&PageUrl&"Page="&PageRs.PageCount&"'>>>|</a> "
end if
Response.Write"&nbsp;转到第<input type='text' name='page' size='3' maxlength='5' value='" &CurPages & "' onKeyPress=""if (event.keyCode==13) window.location='" & PageUrl & "page=" & "'+this.value;""'>页"

case 3 ' 分页样式3
Response.Write"[共<Font color='Red'><b>"&PageRs.RecordCount&"</b></Font>条记录]"
Response.Write"[页次<Font color='Red'>"&CurPages&"</Font>/"&PageRs.PageCount&"页]"
If CurPages<=1 Then
Response.Write "[首页][上一页] "
Else
Response.Write"[<a href='"&PageUrl&"Page=1'>首页</a>&nbsp;]"
Response.Write"[<a href='"&PageUrl&"Page="&(CurPages-1)&"'>上一页</a>]"
End If
If CurPages>=PageRs.PageCount Then
Response.Write"[下一页][尾页]"
Else
Response.Write"[<a href='"&PageUrl&"Page="&(CurPages+1)&"'>下一页</a>]"
Response.Write"[<a href='"&PageUrl&"Page="&PageRs.PageCount&"'>尾页</a>]&nbsp;"
End If
Response.Write"转到：<select name='Page' size='1' onchange=""javascript:window.location='" & PageUrl & "Page=" & "'+this.options[this.selectedIndex].value;"">"
For j = 1 To PageRs.Pagecount
If j=CurPages Then
Response.Write"<option selected value="&j&">"&"第"&j&"页"&"</option>"
Else
Response.Write"<option value="&j&">"&"第"&j&"页"&"</option>"
End If
Next 
Response.Write"</select>" 
end select
End Sub
%>
