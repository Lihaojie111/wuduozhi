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
<body  style="background:#faffe7; font-size:13px;font-family:"微软雅黑";">
<div class="con">
   <% dateid=1 %>
 <!--#include file="head.asp"-->


<!---End-smoth-scrolling---->

			    <script type="text/javascript">
					jQuery(function($) {
						$(".swipebox").swipebox();
					});
				</script>
				<script type="text/javascript">
				var imgReady = function (url, callback, error) {

        var width, height, intervalId, check, div,

        img = new Image(),

        body = document.body;

        

        img.src = url;

        

        // 从缓存中读取

        if (img.complete) {

        return callback(img.width, img.height);

        };

        

        // 通过占位提前获取图片头部数据

        if (body) {

        div = document.createElement('div');

        div.style.cssText ='visibility:hidden;position:absolute;left:0;top:0;width:1px;height:1px;overflow:hidden';

        div.appendChild(img)

        body.appendChild(div);

        width = img.offsetWidth;

        height = img.offsetHeight;

        

        check = function () {

        if (img.offsetWidth !== width || img.offsetHeight !== height) {

        clearInterval(intervalId);

        callback(img.offsetWidth, img.clientHeight);

        img.onload = null;

        div.innerHTML = '';

        div.parentNode.removeChild(div);

        };

        };

        

        intervalId = setInterval(check, 150);

        };

        

        // 加载完毕后方式获取

        img.onload = function () {

        callback(img.width, img.height);

        img.onload = img.onerror = null;

        clearInterval(intervalId);

        body && img.parentNode.removeChild(img);

        };

        

        // 图片加载错误

        img.onerror = function () {

        error && error();

        clearInterval(intervalId);

        body && img.parentNode.removeChild(img);

        };

        

        };
				
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
      <div class="chanpin wow bounceIn animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;">
         <div class="biaoti"><img src="images/cf.png"></div><!--biaoti-->
         <div class="cpnr">
            <ul>
			 <%
i=1
Sql = "Select Top 4  * From [class] where Layout='Product' Order By classID ASC"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof

%> 
              <li style="border-left:1px #7c970c solid;"><a href="product.asp?classid=<%=rs("classid")%>"><div class="cptu<%=i%>"></div></a><a href="product.asp?classid=<%=rs("classid")%>"><h1><%=rs("classname")%></h1><p><%=rs("enname")%></p></a></li>
			  
			  	        <%
i=i+1
Rs.MoveNext
Loop
Rs.Close
%>   
   
            </ul><div class="clear"></div>
         </div><!--cpnr-->
            
                
      </div><!--chanpin-->
      
        

<%
  Set Rs3=Server.CreateObject("ADODB.Recordset")
  
  sql1="select * from [product] "
  Rs3.Open Sql1,Conn,1,1

	Do While Not Rs3.Eof

  cci=rs3.recordcount
 
  ccii=cci/2
  
 cciii=int(ccii)

%> 
<%
Rs3.MoveNext
Loop
Rs3.Close

%>




  <div class="ht_hunyanjiudian_box">
    <div class="ht_hunyanjiudian_box_cen_fa ew_cls">
      <div class="ht_hunyanjiudian_box_le"></div>
      <div class="ht_hunyanjiudian_box_cen">
        <ul>
   
<%

for a=1 to cci step 2

b=a-2





%>   

 <li>  
          
          <%
		  if a=1 then
		  
		  Sql = "Select Top 2 * From product  Order By ID desc"  
            
Rs2.Open Sql,Conn,1,1

Do While Not Rs2.Eof
		  
		  
		  %>
          
          
                        <div class="ht_hunyanjiudian_box_cen_one"><a href="product_view.asp?id=<%=rs2("id")%>&classid=<%=rs2("classid")%>"><img src="<%=rs2("pic")%>" width="271" height="191" /></a>
              <div class="ht_hovertext"><%=left(rs2("title"),7)%>...</div>
            </div>
            
            
<%

Rs2.MoveNext
Loop
Rs2.Close
a=2

%>
          
     <% else %>     
          
         

          <%

Sql = "Select Top "&cci&" * From product Where  ID not in (Select Top "&b&" ID From product Order By ID desc) Order By ID desc"  
            
Rs2.Open Sql,Conn,1,1

Do While Not Rs2.Eof
%> 
            
                        <div class="ht_hunyanjiudian_box_cen_one"><a href="product_view.asp?id=<%=rs2("id")%>&classid=<%=rs2("classid")%>"><img src="<%=rs2("pic")%>" width="271" height="191" /></a>
              <div class="ht_hovertext"><p><%=left(rs2("title"),7)%></p></div>
            </div>
            
            
<%

Rs2.MoveNext
Loop
Rs2.Close

end if
%>


 


                        

            

                        
              <%=cci %>

      
            
             </li>
                        
            
            
<%
next

%>
                        

	
        </ul>
		 <div class="clear"></div>
      </div>
      <div class="ht_hunyanjiudian_box_ri"></div>
    </div>
  </div>
</div>
  <!---------------------------------------------------hyjd--------------------------------------------->



        
        
        
        
        
        
        
        
       <!--产品图片-->
      
  
  <div class="jianjie">
     <div class="jieshao  wow bounceInLeft animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;">
	 	          <%
Sql = "Select  * From [about] where ID=24 Order By ID desc,DateAndTime  desc"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof

%> 
	 <img src="<%=rs("pic")%>">
     <%=left(rs("Content"),300)%>...
	        <%

Rs.MoveNext
Loop
Rs.Close
%>  

<div class="gd"><a href="about.asp">查看更多>></a></div><!--gengduo--><div class="clear"></div>
</div><!--jieshao-->
  
     <div class="shipin  wow bounceInRight animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;">
	 <%
Sql = "Select Top 1 * From [spzs] where Topis > 0  Order By ID Desc,DateAndTime  desc"
Rs.Open Sql,Conn,1,1
Do While Not Rs.Eof

%>
	 <%=rs("Content")%>
	         <%

Rs.MoveNext
Loop
Rs.Close
%>  
	 
	 </div><!--jieshao-->
 <div class="clear"></div>
  </div><!--jianjie-->
  
 <!--------------------------------------------------------------新闻----------------------------------------------------->
 
        
 
        
        <div id="section1" init="true" class="section section1">
        <header class="hearer">
        
    
         
        
     
          
       
         
        
        
        </header>
		
       
        
        
     
		<script src="js/base.js" type="text/javascript"></script>
		<script src="js/index.js" type="text/javascript"></script>
		<script src="js/jquery.scrollbox.js" type="text/javascript"></script>
		<script src="js/jquery.swipebox.min.js" type="text/javascript"></script>
       
        
 

 
 
 <div class="dongtai">
   
      <div class="biaoti"><img src="images/dt.png"></div><!--biaoti-->
 
 
         <div class="wrap">
            <div class="hn_main fl  wow bounceInLeft animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;"><div class="home_shade"></div>
            
            <!-- <img src="images/554c8faf964d7.jpg" alt="" class='top_news_img' /> -->
            <img src="images/xw.jpg" alt="" class='top_news_img' />
            <%
Sql = "Select Top 1 * From [news]  Order By ID Desc,DateAndTime  desc"
Rs.Open Sql,Conn,1,1
Do While Not Rs.Eof

%>
            <h2><a href="news_view.asp?classid=<%=rs("classid")%>&id=<%=rs("id")%>" target="_blank"><%=left(rs("title"),20)%></a></h2><p class="time"><i></i><%= FormatTime(""&rs("DateAndTime")&"",4) %></p>
            <p class="text"><a href="news_view.asp?classid=<%=rs("classid")%>&id=<%=rs("id")%>" target="_blank">    <%=cutStr(rs("Content"),128)%></a></p>
            <%

Rs.MoveNext
Loop
Rs.Close
%>
 </div>           
            <div class="hn_column fr  wow bounceInRight animated" data-wow-delay="0.5s" style="visibility: visible; -webkit-animation-delay: 0.5s;"><ul>
                        <%
Sql = "Select Top 5 * From [news]  where ID not in (select top 1 ID from news order by [id] desc)  Order By ID Desc,DateAndTime  desc"
Rs.Open Sql,Conn,1,1
Do While Not Rs.Eof

%>
            <li><h2><a href="news_view.asp?classid=<%=rs("classid")%>&id=<%=rs("id")%>" target="_blank"><%=left(rs("title"),20)%></a><p class="time" style="float:right; font-size:12px; font-weight:normal; margin-right:50px;"><i></i><%= FormatTime(""&rs("DateAndTime")&"",4) %></p></h2><p class="text"><a href="news_view.asp?classid=<%=rs("classid")%>&id=<%=rs("id")%>" target="_blank"><%=cutStr(rs("Content"),360)%></a></p></li>
			            <%

Rs.MoveNext
Loop
Rs.Close
%>
            
       
            
            
</ul>
        </div>


</div><!--xinwenneirong--> <div class="clear"></div>
 
 <div class="gd1" style="margin-top:45px;"><a href="news.asp?classid=6">查看更多>></a></div><!--gengduo-->

 
 </div><!--gongsidongtai-->
 
 
 
 
  <!--------------------------------------------------------------新闻----------------------------------------------------->
 
 
 
 <!--#include file="liuyan.asp"-->
  
  
</div><!--neirong-->
  
     

  
  
<!--#include file="foot.asp"-->
</div><!--con-->
</body>
</html>
