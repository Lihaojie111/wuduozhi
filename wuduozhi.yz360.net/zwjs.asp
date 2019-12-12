<!--#include file="System.asp"-->
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
<title><%=title%></title>
<meta name="keywords" content="<%=Keywords1%>,<%=Keywords2%>,<%=Keywords3%>,<%=Keywords4%>"  />
<meta name="description" content="<%=DescriptionWord%>"  />
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
      <div class="jiaeuwomen">
         <div class="biaoti"><img src="images/jr.png"></div><!--biaoti-->
         <div class="jrnrjs"> 
                     <%
ID=Request("id")
Sql = "Select  * From [join] where ID="&id&" Order By ID desc,DateAndTime  desc"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof

%> 
<h1>
	<%=rs("zhiwei")%>
</h1>

<span><a href="javascript:window.history.go(-1);">返回上一页</a></span> 
<div class="clear"></div>
<p><%= FormatTime(""&rs("DateAndTime")&"",4) %></p>
<%=rs("Content")%>

     <%

Rs.MoveNext
Loop
Rs.Close
%>
		   
         </div><!--jiaruwomen-->
                
      </div><!--关于我们-->
      
        






 




        
        
        
        
        
        
   
 
<!--#include file="liuyan.asp"-->
  
  
  
  
</div><!--neirong-->
  
     

  
  
<!--#include file="foot.asp"-->
</div><!--con-->
</body>
</html>
