<!--#include file="System.asp"-->
<!--#include file="inc/page.asp"-->
<%
Set Rs2=Server.CreateObject("ADODB.Recordset")
Sql="Select * From seo "
Rs2.Open Sql,Conn,1,2
If Rs2.Eof Then
		Msg "暂未添加!","index.asp",1
Else
		
		Rs2.Update
		Title=Rs2("Title")
		Keywords1=Rs2("Keywords1")
		Keywords2=Rs2("Keywords2")
		Keywords3=Rs2("Keywords3")
		Keywords4=Rs2("Keywords4")
		DescriptionWord=Rs2("DescriptionWord")
	
End If
Rs2.Close 
%>
<!DOCTYPE html>
<html>

<head lang="en">
    <meta charset="gb2312">
<title><%=title%></title><!--标题-->
<meta name="keywords" content="<%=Keywords1%>,<%=Keywords2%>,<%=Keywords3%>,<%=Keywords4%>"  /><!--优化关键词-->
<meta name="description" content="<%=DescriptionWord%>"  /><!--优化关键词-->
    <link href="css/index.css" rel="stylesheet"type="text/css"/>
   <script type="text/javascript" src="js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="js/jquery.SuperSlide.2.1.1.js"></script>
   

<script type="text/javascript" src="js/biglv.js"></script>
<script type="text/javascript" src="js/jquery.scrollbox.js"></script>


</head>
<body  style="background:#faffe7; font-size:13px;">
<div class="con">
   <% dateid=1 %>
 <!--#include file="head.asp"-->


<!---End-smoth-scrolling---->

			    <script type="text/javascript">
					jQuery(function($) {
						$(".swipebox").swipebox();
					});
				</script>
				<!--Animation-->
<script src="js/wow.min.js"></script>
<link href="css/animate.css" rel='stylesheet' type='text/css' />
<style type="text/css">
body,td,th {
	font-family: "Open Sans", sans-serif;
}
</style>
<script>
	new WOW().init();
</script>
<!---/End-Animation---->	

  
  <div class="neirong">
      <div class="shipinzhanshi">
         <div class="biaoti"><img src="images/sp.png"></div><!--biaoti-->
         
         <div class="jichunr">
             <ul>
			 <%
intPageSize=6
PageUrl="vedio.asp"     
SQLSelect="select * from [spzs]"
SQLWhere=" where 1=1"
''/*分类搜索*/
If Request("ClassID")<>"" Then
	SQLWhere=SQLWhere&" and "&ClassSQL(ClassID)&""
	PageUrl=PageUrl&"?ClassID="&ClassID&""   
end if
SQLOrder=" order by [topis] desc,[dateandtime] desc"
SQLTable=trim(mid(SQLSelect,(instr(SQLSelect,"from")+4)))
SQL1="select count(*) from [spzs]"&SQLWhere
intRecordCount=conn.execute(SQL1)(0)
call PageCute(intRecordCount,intPageSize)
SQL=SQLSelect&SQLWhere&SQLOrder
set rs=conn.execute(SQL)
if not rs.eof then 
rs.move (CurrentPage-1)*intPageSize
for i=1 to intPageSize
if rs.eof then exit for
%>
                <li><a href="vedio_view.asp?classid=<%=rs("classid")%>&id=<%=rs("id")%>"><img src="images/spt.jpg"></a> <div class="yincang2"><a href="vedio_view.asp?classid=<%=rs("classid")%>&id=<%=rs("id")%>"><h4><%=left(rs("title"),10)%></h4></a> </div></li>
				
				<%       
rs.movenext
next
else
Response.Write " 暂无！" & vbNewLine
end if
%>
				
                
           
             </ul><div class="clear"></div>
			 <div class="page">
 <%
set rs=nothing
response.write "共<font font-size:13px; color='#FF0000'>"&intRecordCount&"</font>件产品&nbsp;&nbsp;"
response.write Page_List(PageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize,10,"FF0000")

%></div>
			 
         </div><!--jichunr-->
     
                
      </div><!--视频展示-->
     
<!--#include file="liuyan.asp"-->

</div><!--neirong-->

<!--#include file="foot.asp"-->
</div><!--con-->
</body>
</html>
