<!--#include file="System.asp"-->
<!--#include file="inc/page.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="河南万正远洋船舶管理有限公司|万正远洋|海员招收|海乘招收" />
<meta name="description" content="河南万正远洋船舶管理有限公司业务始于2002年，是具有海事局海员外派资质[资质编号HYWP03026]的国际海员服务机构。" />
<title>河南万正远洋船舶管理有限公司</title>
<link href="css/index.css" rel="stylesheet" type="text/css" />
<script src="js/jquery.js" type="text/javascript"></script>
<SCRIPT type=text/javascript src="js/head.js"></SCRIPT>
<script type="text/javascript">
function AutoScroll(obj){
        $(obj).find("ul:first").animate({
                marginTop:"-15px"
        },500,function(){
                $(this).css({marginTop:"0px"}).find("li:first").appendTo(this);
        });
}
$(document).ready(function(){
setInterval('AutoScroll("#scrollDiv")',1000)
});
</script>
<script type=text/javascript>
	function setTab03Syn ( i )
	{
		selectTab03Syn(i);
	}
	
	function selectTab03Syn ( i )
	{
		switch(i){
			case 1:
			document.getElementById("TabTab03Con1").style.display="block";
			document.getElementById("TabTab03Con2").style.display="none";
	
			document.getElementById("font1").style.color="#fff";
			document.getElementById("font2").style.color="#000";
	
			break;
			case 2:
			document.getElementById("TabTab03Con1").style.display="none";
			document.getElementById("TabTab03Con2").style.display="block";
		
			document.getElementById("font1").style.color="#000";
			document.getElementById("font2").style.color="#fff";
			break;
			
		}
	}

</script>
</head>
<body>
<!--#include file="head.asp"-->
<div class="all">
<div class="con">
<div class="gongsi">
<div class="down">
<div class="bu"><img src="images/renli.jpg" width="203" height="67" /></div>
<div class="women">
<ul>
<li><a href="messages.asp">　　　　在职船员应聘</a></li>
<li><a href="#">　　　　人才招聘</a></li>
</ul>
</div><!--women end-->
</div><!--down end-->
<!--#include file="left.asp"-->

</div><!--gongsi end-->
<div class="youbian">
<div class="ppoe"><img src="images/US.jpg" width="448" height="22" /></div><!--ppoe end-->
<div class="timu"><span>人力资源</span>　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　您当前所在位置：首页 / 人力资源</div><!--timu end-->
<div class="con_right_con_22">
            <table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0"  style="line-height:35px; margin:auto">
            <tr bgcolor="#E3E3E3">
            	<th  width="20%">职务</th>
                <th width="20%">船型</th>
                <th width="20%">航线</th>
                <th width="20%">薪资</th>
                <th width="10%">招聘人数</th>
                <th width="10%">详情</th>
            	
            </tr>
              
     <%
intPageSize=10
PageUrl="job.asp"
SQLSelect="select * from [join]"
SQLWhere=" where 1=1"

SQLOrder=" order by [topis] desc,[dateandtime] desc"
SQLTable=trim(mid(SQLSelect,(instr(SQLSelect,"from")+4)))
SQL1="select count(*) from [join]"&SQLWhere
intRecordCount=conn.execute(SQL1)(0)
call PageCute(intRecordCount,intPageSize)
SQL=SQLSelect
set rs=conn.execute(SQL)
if not rs.eof then 
rs.move (CurrentPage-1)*intPageSize
for i=1 to intPageSize
if rs.eof then exit for
%>

<tr style="text-align:center">
	<td>　　<%=rs("zhiwei")%></td>
    <td><%=rs("leixing")%></td>
    <td><%=rs("address")%></td>
 
    <td><%=rs("gongzi")%></td>
    <td><%=rs("renshu")%></td>
    <td><A href="job_view.asp?id=<%=rs("id")%>">详情</A></td>
</tr>
                
   <% 

rs.movenext
next
%>                   
               
              </table>
            </div>
            <div id="showpages">
<%
set rs=nothing
response.write "共<font color='#FF0000'>"&intRecordCount&"</font>个职位&nbsp;&nbsp;"
response.write Page_List(PageUrl,CurrentPage,intPageCount,intRecordCount,intPageSize,10,"FF0000")
else
Response.Write " 暂无！" & vbNewLine
end if
%>
 <div class="clear"></div><!--clear end--></div>
<!--neirong end-->
</div><!--youbian end-->
<div class="clear"></div>
</div><!--con end-->

<!--#include file="foot.asp"-->
</div><!--all end-->
</body>
</html>
