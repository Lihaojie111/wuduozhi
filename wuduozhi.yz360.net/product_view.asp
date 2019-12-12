<!--#include file="System.asp" -->
<!--#include file="inc/page.asp"-->
<%
dim Title,Pic,Keywords,DescriptionWord,Price
ID=CheckStr(trim(Request("ID")))
Sql="Select * From [product] Where ID="&ID
Rs.Open Sql,Conn,1,2
If Rs.Eof Then
		Msg "项目不存在!","index.asp",1
Else
		Rs("Hits")=Rs("Hits")+1
		Rs.Update
		ClassID=Rs("ClassID")
		Pic=Rs("Pic")
		Hits=Rs("Hits")
		Title=Rs("Title")
		Keywords=Rs("Keywords")
		DescriptionWord=Rs("DescriptionWord")
		Content=Rs("Content")
		DateAndTime=Rs("DateAndTime")
End If
Rs.Close 
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
      
     
     <div class="honor">
         <div class="biaoti"><img src="images/cptou.png"></div><!--biaoti-->
         <div class="btfl1">
           <ul>
		            <%
Sql = "Select  * From [class] where Layout='Product' Order By classID ASC"            
Rs.Open Sql,Conn,1,1
Do While Not Rs.Eof

%> 
              <li><a href="product.asp?classid=<%=rs("classid")%>"><%=rs("classname")%></a></li>
	        <%

Rs.MoveNext
Loop
Rs.Close
%>   		  
			  
           </ul></div><!--biaotifenlei--><div class="clear"></div>
         <div class="cpnr1">

    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="40" align="center" valign="bottom"  style="padding-top:15px;"><h4><%= Title %></h4></td>
          </tr>
          <tr>
            <td 
                      style="BORDER-RIGHT: #c5c5c5 1px dashed; BORDER-TOP: #c5c5c5 1px dashed; BORDER-LEFT: #c5c5c5 1px dashed; BORDER-BOTTOM: #c5c5c5 1px dashed" 
                      align="center" height="28">【发布时间：<%= FormatTime(""&DateAndTime&"",4) %>】&nbsp;【作者/来源&nbsp;维护员】&nbsp;【阅读：<%= Hits %>次】&nbsp;【<a href="javascript:window.history.go(-1);" >返 回</a>】</td>
          </tr>
          <tr>
            <td align="center" height="10"></td>
          </tr>
          <tr>
            <td align="center"><a href="<%=pic%>" target="_blank"><img style="padding:10px; border:1px #666666 solid"  src="<%=pic%>" onload = "DrawImage(this)" /></a></td>
          </tr>
          <tr>
            <td height="10"></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  
                  <td align="left"><%= Content %></td>
                </tr>
              </table>
                <br /></td>
          </tr>
        </table>

 </div><!--cpnr1-->
     
     </div><!--zizhirongyu-->
     
         
         
                
 
        






 




        
        
        
        
        
        
   
 
<!--#include file="liuyan.asp"-->
  
  
  
  
</div><!--neirong-->
  
     

  
  
<!--#include file="foot.asp"-->
</div><!--con-->
</body>
</html>