<!--#include file="System.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台界面</title>
<style> 
body{}
a{ text-decoration:none;}  
.top_menu a{
	display:block;
	height:38px;
	 width:174px;
	 padding-left:54px;
	line-height:36px;
	font-size:14px; 
	color:#3f3f3f;
    font-family:"微软雅黑";} 
.top_menu a:hover{background:url(left_bg1.jpg);}
.sub_meuu a{ font-size:12px; color:#FFF;
	display:block;
	width:164px;
	padding-left:64px;
	height:36px;
	line-height:36px;
	font-family:"微软雅黑";
} 
.hover:hover{background:url(bg3.jpg);
display:block;
color:#dee318;}
.right_box_t{background:url(images/shangyouBg.gif) repeat-x;height:36px;line-height:36px;padding-left:20px;}
.right_box{background:#fff;}
</style> 
<script language="javascript"> 
tempj=2; 
function showed(tempi) 
{ 
if(document.getElementById("show"+tempj.toString()).style.display==''&&tempi!=tempj) 
{ 
document.getElementById("show"+tempj.toString()).style.display='none'; 
} 
if(document.getElementById("show"+tempi.toString()).style.display=='none') 
{ 
document.getElementById("show"+tempi.toString()).style.display=''; 
} 
else 
{ 
document.getElementById("show"+tempi.toString()).style.display='none'; 
} 
tempj=tempi; 
} 


</script> 
<script language="javascript" type="text/javascript">

function login(){
	var loginname = $('#username').val();
	var loginpsw = $('#userpsw').val();
	if(loginname=='' || loginpsw=='') {
		alert('用户名和密码不允许为空!')
		return false;	
	}
	return true;
}
</script>
<link type="text/css" rel="stylesheet" href="css.css" />
</head>
<body>
<!--头部-->
<div class="wramp">
  	 <!--#include file="head.asp"-->
  <!--左侧-->

  <div class="clear"></div>
  
  
  <div id="content">
  
  	 <!--#include file="left.asp"-->
	
	<div class="c-right" style="padding-bottom:15px;">
	<FORM onSubmit="return login()" method=post name=form1 action=http://e.258.com/login.php>
<DIV class=right_box style="background:#f0efef">
<DIV class=right_box_t style="border-bottom:1px #dbdbdb solid;"><SPAN  style="color:#ff6600; font-size:14px; font-weight:bold;">商友会员,马上登录>></SPAN></DIV>
<!--end right_box_t-->
<DIV class=right_box_c style="height:100%;background:#f0efef">
<TABLE border=0 cellSpacing=2 cellPadding=2 width=310 align=left>
  <TBODY>
  <TR>
    <TD width="302">
      <TABLE border=0 cellSpacing=3 cellPadding=3 width=220 align=center style="margin-top:10px">
        <TBODY>
        <TR>
          <TD width=40><SPAN class="f14px fB">帐号</SPAN></TD>
          <TD><INPUT style="WIDTH: 150px" id=username type=text 
          name=username></TD></TR>
        <TR>
          <TD><SPAN class="f14px fB">密码</SPAN></TD>
          <TD><input style="WIDTH: 150px" id=userpsw type=password 
            name=userpsw /></TD></TR>
        <TR>
          <TD>&nbsp;</TD>
          <TD><INPUT src="images/huiyuan_denglu.gif" 
            type=image name=button>            <span><A  target="_blank"
            href="http://e.258.com/getpassword.php">忘记密码？</A></span></TD>
        </TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD height=6></TD></TR>
  
  <TR>
    <TD><SPAN  style="color:#ff6600; font-size:14px; font-weight:bold;padding-left:40px;"><a href="#" target="_blank">还不是会员，马上注册</a></SPAN></TD></TR>
  <TR style="padding-left:50px;">
    <TD><a target="_blank" href="http://e.258.com/register_free.php"><img border=0 
      src="images/tj_btn.gif" width=220 
    height=41  style="padding-left:40px;"/></a></TD>
  </TR></TBODY></TABLE>
<div style="float:right;padding:50px 200px 45px 15px;background:url(images/tips.gif) no-repeat left top;"><span style="font-size:14px;font-weight:bold;color:#000000;">[开始做大生意吧] </span>
<div>
免费注册后您即可 
<ul>
<li>·发布信息，推广产品、宣传公司</li>
 <li>·创建公司商铺，全球展示商机</li>
 <li>·查看营销效果，与客户及时交流</li></ul></div> </div><div style="clear:both"></div>
</DIV>
<!--end right_box_c-->
 <!--#include file="gongju.asp"-->
</DIV>

	</FORM>
	</div>
  </div>
</div>
</body>
</html>
