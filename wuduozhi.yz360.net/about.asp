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
       <script type="text/javascript" src="js/pirobox.js"></script>
<script type="text/javascript" src="js/biglv.js"></script>
<script type="text/javascript" src="js/jquery.scrollbox.js"></script>
<script type="text/javascript">

		$(document).ready(function() {

			$('').piroBox({

					my_speed: 400, //animation speed

					bg_alpha: 0.8, //background opacity

					slideShow : true, // true == slideshow on, false == slideshow off

					slideSpeed : 4, //slideshow duration in seconds(3 to 6 Recommended)

					close_all : '.piro_close,.piro_overlay'// add class .piro_overlay(with comma)if you want overlay click close piroBox

			});

		});

	</script>

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
      <div class="guanyuwomen">
         <div class="biaoti"><img src="images/qyjj.png"></div><!--biaoti-->
         <%
Sql = "Select  * From [about] where ID=24 Order By ID desc,DateAndTime  desc"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof

%> 
       <div class="jianjieneirong">  
         <div class="jjtu  wow bounceInLeft animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;"><img src="<%=rs("pic")%>"></div><!--jieshao-->
     <div class="jianjienr  wow bounceInRight animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;">
	 
	 <%=rs("Content")%>
	 </div>

<div class="clear"></div>

        </div> <!----------------------jianjieneirong----------------------------->
        <%

Rs.MoveNext
Loop
Rs.Close
%>      
         
         
        <div class="wenhua">
         <div class="biaoti"><img src="images/wh.png"></div><!--biaoti-->
           <div class="wenhuanr">
          <%
Sql = "Select  * From [about] where ID=25 Order By ID desc,DateAndTime  desc"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof

%> 
           <div class="jjtu1  wow bounceInLeft animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;"><img src="<%=rs("pic")%>"></div><!--jieshao-->
     <div class="jianjienr1  wow bounceInRight animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;">
	 
	 <%=rs("Content")%>
	 
	 </div>

<div class="clear"></div>

     <%

Rs.MoveNext
Loop
Rs.Close
%>

        </div><!--neirong-->
    </div><!--wenhua--> 
         

     
     
     <div class="honor">
         <div class="biaoti"><img src="images/zz.png"></div><!--biaoti-->
         <div class="honornr">
            <ul>
			 <%
Sql = "Select  * From [zzry] where ClassID=1 Order By ID desc,DateAndTime  desc"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof

%> 
              <li><a href="<%=rs("pic")%>" class='pirobox_gall'  title='<%=rs("title")%>' ><img src="<%=rs("pic")%>"></a></li>
			  
<%

Rs.MoveNext
Loop
Rs.Close
%>

             </ul> <div class="clear"></div>
         </div><!--honornr-->
     
     </div><!--zizhirongyu-->
     
         
         
                
      </div><!--关于我们-->
      
        






 




        
        
        
        
        
        
   
 
<!--#include file="liuyan.asp"-->
  
  
  
  
</div><!--neirong-->
  
     

  
  
<!--#include file="foot.asp"-->
</div><!--con-->
</body>
</html>
