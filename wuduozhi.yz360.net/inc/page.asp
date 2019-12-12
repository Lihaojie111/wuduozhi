<%
'---------------------���ݿ���غ���-----------------
'****************************************************
'��������DBDel
'��  �ã�ɾ��ָ������������
'****************************************************
Function DBDel(Table,WhereStr)
	Dim Sql
	Sql="Delete From "&Table&" Where "&WhereStr&""
	Conn.Execute(Sql)
End Function
'****************************************************
'��������DBUpdate
'��  �ã��޸�ָ���������ֶε�����
'****************************************************
Function DBUpdate(Table,UpVale,WhereStr)
	Dim Sql
	Sql="Update "&Table&" Set "&UpVale&" Where "&WhereStr&""
	Conn.Execute(Sql)
End Function
'****************************************************
'��������DBGetNum
'��  �ã���ѯĳ�ֶ����� ���ֵ����Сֵ
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
'��������DBExist
'��  �ã���ѯĳ�ֶε������Ƿ����ظ�
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
'��������DBGetValue
'��  �ã���ѯָ������ĳ�ֶε�����
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

'---------------------��ҳ��غ���-------------------
'*************************************************
'��������PageCute(strUrl,totalput,MaxPerPage,ShowTotal,strUnit,pagecode,CurrentPage)
'��  �ã�������ҳ���Լ���ز���
'��  ������
'����ֵ����
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
'��  ����ShowPage(strUrl,totalput,MaxPerPage,ShowTotal,strUnit,pagecode,CurrentPage)
'��  �ã����һ����ҳ�б�[��ҳ][��ҳ][��ҳ][ĩҳ]
'��  ����CurrentPage ---- ��ǰ��ʾ��ҳ
'	 	 intintPageCount ---- �ܹ�ҳ��
'        intRecordCount----�ܼ�¼��
'        strPageURL ---- ҳ����ָ�������
'����ֵ����
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
    RW "��ҳ ��һҳ&nbsp;"
  else
    RW "<a href="&strPageurl&"page=1  class='hui1'><font color='#FF0000'>��ҳ</font></a>&nbsp;"
    RW "<a href="&strPageurl&"page=" & CurrentPage-1 & "  class='hui1'><font color='#FF0000'>��һҳ</font></a>&nbsp;"
  end if
  if intPageCount-CurrentPage<1 then
    RW "��һҳ βҳ"
  else
    RW "<a href="&strPageurl&"page="& (CurrentPage+1) &"><font color='#FF0000'>��һҳ</font></a>&nbsp;"
    RW "<a href="&strPageurl&"page="&intPageCount&"><font color='#FF0000'>βҳ</font></a>"
  end if
    RW "&nbsp;&nbsp;ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&intPageCount&"</strong>ҳ "
    RW "&nbsp;��<b><font color='#FF0000'>"&intRecordCount&"</font></b>����¼<b>"
	RW "</td></tr></table>"
end function


'*************************************************
'��  ����Page_List(strPageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize,pp,font_color)
'��  �ã����һ����ҳ�б�1,2,3,4,5
'��  ����CurrentPage ---- ��ǰ��ʾ��ҳ
'	 	 intintPageCount ---- �ܹ�ҳ��
'        intRecordCount----�ܼ�¼��
'        strPageURL ---- ҳ����ָ�������
'        pp---ÿҳ��ʾ��ҳ��
'        font_color-----��ɫ
'����ֵ����
'*************************************************
Const btn_first="��ҳ"  '�����һҳ��ť��ʾ��ʽ
Const btn_prev="��һҳ"  '����ǰһҳ��ť��ʾ��ʽ
Const btn_next="��һҳ"  '������һҳ��ť��ʾ��ʽ
Const btn_last="βҳ"  '�������һҳ��ť��ʾ��ʽ

function Page_List(strPageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize,pp,font_color)
	If strPageUrl<>"?" Then
		If InStr(strPageUrl,"?")=0 Then
			strPageUrl=strPageUrl&"?"
		Else
			strPageUrl=strPageUrl&"&"
		End If
	End If
	page_list="ҳ�Σ�<font color=FF0000>"&CurrentPage&"</font>/"&intPageCount&""
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
     page_list=page_list&" <a href='"& strPageUrl &"' title='��һҳ'  class='hui1'>"&btn_first&"</a> "
  end if
  if clng(CurrentPage)>1 then
	 page_list=page_list&" <a href='"& strPageUrl &"page="&CurrentPage-1&"' class='hui1' title='��һҳ'>"&btn_prev&"</a> "
  end if
  for pi=pl to pr
    if clng(CurrentPage)=clng(pi) then
      page_list=page_list&" <font color=" & font_color & " face=����><b>" & pi & "</b></font> "
    else
      page_list=page_list&" <a  class='hui1' href='"& strPageUrl &"page="& pi &"' title='�� " & pi & " ҳ' clases=0 style=""font-family:����""><b>" & pi & "</b></a> "
    end if
  next
    if clng(CurrentPage)<clng(intPageCount) then 	  
	  page_list=page_list&" <a class='hui1' href='"& strPageUrl &"page="&CurrentPage+1&"' clases=0 title='��һҳ'>"&btn_next&"</a> "
    end if
	if clng(CurrentPage)=clng(intPageCount) then
      page_list=page_list&" <font color=" & font_color & "> "&btn_last&" </font> "
	else
	  page_list=page_list&" <a class='hui1' href='"& strPageUrl &"page="& intPageCount &"' title='���һҳ'>"&btn_last&"</a> "
	end if
end function
%>