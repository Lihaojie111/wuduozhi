<!--#include file="System.asp" -->
<!--#include file="inc/page.asp"-->
<%
dim Title,DateAndTime
ID=CheckStr(trim(Request("ID")))
Sql="Select * From [spzs] Where ID="&ID
Rs.Open Sql,Conn,1,2
If Rs.Eof Then
		Msg "��Ƶ������!","vedio.asp",1
Else
		Rs.Update
		ClassID=Rs("ClassID")
		Title=Rs("Title")
		Content=Rs("Content")
		DateAndTime=Rs("DateAndTime")
End If
Rs.Close 
%>
<!DOCTYPE html>
<html>

<head lang="en">
    <meta charset="gb2312">
<title><%=title%></title><!--����-->
<meta name="keywords" content="<%=Keywords1%>,<%=Keywords2%>,<%=Keywords3%>,<%=Keywords4%>"  /><!--�Ż��ؼ���-->
<meta name="description" content="<%=DescriptionWord%>"  /><!--�Ż��ؼ���-->
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="40" align="center"><h4><%= Title %></h4></td>
        </tr>
        <tr>
          <td 
                      style="BORDER-RIGHT: #c5c5c5 1px dashed; BORDER-TOP: #c5c5c5 1px dashed; BORDER-LEFT: #c5c5c5 1px dashed; BORDER-BOTTOM: #c5c5c5 1px dashed" 
                      align="center" height="28">������ʱ�䣺<%= FormatTime(""&DateAndTime&"",4) %>��&nbsp;������/��Դ&nbsp;ά��Ա��&nbsp;��<a href="javascript:window.history.go(-1);" >�� ��</a>��</td>
        </tr>
        <tr>
          <td class="lh20" style="line-height:20px; padding:10px 10px" align="center"><br />
          <%= Content %>
          
          </td>
        </tr>
      </table>
</div><!--jichunr-->
          
      </div><!--��Ƶչʾ-->

<!--#include file="liuyan.asp"-->

</div><!--neirong-->

<!--#include file="foot.asp"-->
</div><!--con-->
</body>
</html>
