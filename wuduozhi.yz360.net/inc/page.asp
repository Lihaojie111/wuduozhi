<%
'---------------------数据库相关函数-----------------
'****************************************************
'函数名：DBDel
'作  用：删除指定条件的数据
'****************************************************
Function DBDel(Table,WhereStr)
	Dim Sql
	Sql="Delete From "&Table&" Where "&WhereStr&""
	Conn.Execute(Sql)
End Function
'****************************************************
'函数名：DBUpdate
'作  用：修改指定条件和字段的内容
'****************************************************
Function DBUpdate(Table,UpVale,WhereStr)
	Dim Sql
	Sql="Update "&Table&" Set "&UpVale&" Where "&WhereStr&""
	Conn.Execute(Sql)
End Function
'****************************************************
'函数名：DBGetNum
'作  用：查询某字段数量 最大值或最小值
'****************************************************
Function DBGetNum(Table,FieldName,sType,WhereStr)
	Dim Rs,Sql
	Select Case sType
		Case 0
			Sql="Select Count("&FieldName&") From "&Table&" Where "&WhereStr&""
		Case 1
			Sql="Select Max("&FieldName&") From "&Table&" Where "&WhereStr&""
		Case 2
			Sql="Select Min("&FieldName&") From "&Table&" Where "&WhereStr&""
	End Select
	Set Rs=Conn.Execute(Sql)
	If Rs.Bof And Rs.Eof Then
		DBGetNum=0
	Else
		DBGetNum=Rs(0)
	End If
	Rs.Close
	Set Rs=Nothing
End Function
'****************************************************
'函数名：DBExist
'作  用：查询某字段的内容是否有重复
'****************************************************
Function DBExist(Table,FieldName,WhereStr)
	Dim Rs,Sql
	Sql="Select "&FieldName&" From "&Table&" Where "&WhereStr&""
	Set Rs=Conn.Execute(Sql)
	If Rs.Bof And Rs.Eof Then
		DBExist=False
	Else
		DBExist=True
	End If
	Rs.Close
	Set Rs=Nothing
End Function
'****************************************************
'函数名：DBGetValue
'作  用：查询指定条件某字段的内容
'****************************************************
Function DBGetValue(Table,FieldName,WhereStr)
	Dim Rs,Sql
	Sql="Select "&FieldName&" From "&Table&" Where "&WhereStr&""
	Set Rs=Conn.Execute(Sql)
	If Rs.Bof And Rs.Eof Then
		DBGetValue=""
	Else
		DBGetValue=Rs(0)
	End If
	Rs.Close
	Set Rs=Nothing
End Function

'---------------------分页相关函数-------------------
'*************************************************
'过程名：PageCute(strUrl,totalput,MaxPerPage,ShowTotal,strUnit,pagecode,CurrentPage)
'作  用：计算总页数以及相关参数
'参  数：无
'返回值：无
'*****************************************************
dim intPageCount,CurrentPage
sub PageCute(intRecordCount,intPageSize)
  if intRecordCount mod intPageSize= 0 then
    intPageCount=intRecordCount\intPageSize
  else
    intPageCount=intRecordCount\intPageSize+1
  end if 
CurrentPage=clng(request("page"))
if not(isnumeric(CurrentPage)) then CurrentPage=1
If CurrentPage< 1 Then CurrentPage=1
If CurrentPage > intPageCount Then CurrentPage =intPageCount
end sub
'*************************************************
'函  数：ShowPage(strUrl,totalput,MaxPerPage,ShowTotal,strUnit,pagecode,CurrentPage)
'作  用：输出一个分页列表[首页][上页][下页][末页]
'参  数：CurrentPage ---- 当前显示的页
'	 	 intintPageCount ---- 总共页数
'        intRecordCount----总记录数
'        strPageURL ---- 页向导所指向的连接
'返回值：无
'*************************************************
function ShowPage(strPageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize)
	If strPageUrl<>"?" Then
		If InStr(strPageUrl,"?")=0 Then
			strPageUrl=strPageUrl&"?"
		Else
			strPageUrl=strPageUrl&"&"
		End If
	End If
RW "<table><tr><td>"
	if CurrentPage<2 then      
    RW "首页 上一页&nbsp;"
  else
    RW "<a href="&strPageurl&"page=1  class='hui1'><font color='#FF0000'>首页</font></a>&nbsp;"
    RW "<a href="&strPageurl&"page=" & CurrentPage-1 & "  class='hui1'><font color='#FF0000'>上一页</font></a>&nbsp;"
  end if
  if intPageCount-CurrentPage<1 then
    RW "下一页 尾页"
  else
    RW "<a href="&strPageurl&"page="& (CurrentPage+1) &"><font color='#FF0000'>下一页</font></a>&nbsp;"
    RW "<a href="&strPageurl&"page="&intPageCount&"><font color='#FF0000'>尾页</font></a>"
  end if
    RW "&nbsp;&nbsp;页次：<strong><font color=red>"&CurrentPage&"</font>/"&intPageCount&"</strong>页 "
    RW "&nbsp;共<b><font color='#FF0000'>"&intRecordCount&"</font></b>条记录<b>"
	RW "</td></tr></table>"
end function


'*************************************************
'函  数：Page_List(strPageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize,pp,font_color)
'作  用：输出一个分页列表1,2,3,4,5
'参  数：CurrentPage ---- 当前显示的页
'	 	 intintPageCount ---- 总共页数
'        intRecordCount----总记录数
'        strPageURL ---- 页向导所指向的连接
'        pp---每页显示的页数
'        font_color-----颜色
'返回值：无
'*************************************************
Const btn_first="首页"  '定义第一页按钮显示样式
Const btn_prev="上一页"  '定义前一页按钮显示样式
Const btn_next="下一页"  '定义下一页按钮显示样式
Const btn_last="尾页"  '定义最后一页按钮显示样式

function Page_List(strPageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize,pp,font_color)
	If strPageUrl<>"?" Then
		If InStr(strPageUrl,"?")=0 Then
			strPageUrl=strPageUrl&"?"
		Else
			strPageUrl=strPageUrl&"&"
		End If
	End If
	page_list="页次：<font color=FF0000>"&CurrentPage&"</font>/"&intPageCount&""
 if clng(intPageCount)=0 then
    page_list=page_list&"<font color=" & font_color & "> "&btn_first&" <b>1</b> "&btn_last&" </font>"
    exit function
  end if
  dim pi,ppp,pl,pr
  pi=1
  ppp=pp\2
  if pp mod 2 = 0 then ppp=ppp-1
  pl=CurrentPage-ppp
  pr=pl+pp-1
  if pl<1 then
    pr=pr-pl+1
    pl=1
    if pr>intPageCount then pr=intPageCount
  end if
  if pr>clng(intPageCount) then
    pl=pl+intPageCount-pr
    pr=intPageCount
    if pl<1 then pl=1
  end if
  
  if clng(CurrentPage)=1 then
     page_list=page_list&" <font color=" & font_color & "> "&btn_first&" </font> "
  else
     page_list=page_list&" <a href='"& strPageUrl &"' title='第一页'  class='hui1'>"&btn_first&"</a> "
  end if
  if clng(CurrentPage)>1 then
	 page_list=page_list&" <a href='"& strPageUrl &"page="&CurrentPage-1&"' class='hui1' title='上一页'>"&btn_prev&"</a> "
  end if
  for pi=pl to pr
    if clng(CurrentPage)=clng(pi) then
      page_list=page_list&" <font color=" & font_color & " face=黑体><b>" & pi & "</b></font> "
    else
      page_list=page_list&" <a  class='hui1' href='"& strPageUrl &"page="& pi &"' title='第 " & pi & " 页' clases=0 style=""font-family:黑体""><b>" & pi & "</b></a> "
    end if
  next
    if clng(CurrentPage)<clng(intPageCount) then 	  
	  page_list=page_list&" <a class='hui1' href='"& strPageUrl &"page="&CurrentPage+1&"' clases=0 title='后一页'>"&btn_next&"</a> "
    end if
	if clng(CurrentPage)=clng(intPageCount) then
      page_list=page_list&" <font color=" & font_color & "> "&btn_last&" </font> "
	else
	  page_list=page_list&" <a class='hui1' href='"& strPageUrl &"page="& intPageCount &"' title='最后一页'>"&btn_last&"</a> "
	end if
end function
%>