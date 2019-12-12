<!--#include file="System.asp" -->
<!--#include file="inc/page.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>产品展示</title>
<link type="text/css" href="css.css" rel="stylesheet" />
</head>

<body>
<div class="juzhong">
<!--#include file="head.asp"-->
<div class="aboutus">
  	<div class="aleft">
		<div class="abt">产品展示</div>
<div class="about_b3">
            

  <%
intPageSize=9
PageUrl="product.asp"
SQLSelect="select * from [product]"
SQLWhere=" where 1=1"
''/*分类搜索*/
If Request("ClassID")<>"" Then
	SQLWhere=SQLWhere&" and "&ClassSQL(ClassID)&""
	PageUrl=PageUrl&"?ClassID="&ClassID&""   
end if
SQLOrder=" order by [topis] desc,[dateandtime] desc"
SQLTable=trim(mid(SQLSelect,(instr(SQLSelect,"from")+4)))
SQL1="select count(*) from [Product]"&SQLWhere
intRecordCount=conn.execute(SQL1)(0)
call PageCute(intRecordCount,intPageSize)
SQL=SQLSelect&SQLWhere&SQLOrder
set rs=conn.execute(SQL)
if not rs.eof then 
rs.move (CurrentPage-1)*intPageSize
for i=1 to intPageSize
if rs.eof then exit for
%>	 			
		 <div class="back">
		      <div class="backImg"><a href="product_view.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("pic")%>" height="127" width="172"></a>
			  <br>
			<a href="Product_view.asp?id=<%=Rs("ID")%>" target="_blank" class="hei14cu"><%=LeftStr(Rs("Title"),14)%></a></div>
			  
            </div>
<% 
rs.movenext
next
%>				 
		
			
			
		   
		   <div class="clear"></div>
		   <div align="center" style="padding-top:10px">
<%
set rs=nothing
response.write "共<font color='#FF0000'>"&intRecordCount&"</font>条新闻&nbsp;&nbsp;"
response.write Page_List(PageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize,10,"FF0000")
else
Response.Write " 当前没有新闻！" & vbNewLine
end if
%>
</div>
</div>
</div>
<!--#include file="liu-right.asp"-->
</div>
<!--#include file="foot.asp"-->
</div>
</body>
</html>
