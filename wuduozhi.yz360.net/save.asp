<!--#include file="System.asp"-->
<% 
username=Trim(Request.Form("username"))
tel=Trim(Request.Form("tel"))
email=Trim(Request.Form("email"))
txtContent=Trim(Request.Form("txtContent"))


If username="" Or tel="" or txtContent=""  Then
	Response.Write("<script>alert(""��������������Ϣ."");history.back()</script>")
Else
	set rs=Server.CreateObject("ADODB.RecordSet")
	Sql="Select * From [Guest]"
	Rs.Open Sql,Conn,1,2
		Rs.AddNew
		Rs("Sender")=username
		Rs("Email")=email
		Rs("Tel")=tel
		Rs("Add")=address
		Rs("company")=company
	    Rs("iphone")=iphone
		Rs("Content")=Server.HTMLEncode(txtContent)
		Rs("DateAndTime")=Now()
		Rs.Update
	Rs.Close
	Response.Write("<script>alert(""����Ϣ���ύ����ȴ��ظ���"");location.href=""zxliuyan.asp""</script>")
End If

Response.End
%>