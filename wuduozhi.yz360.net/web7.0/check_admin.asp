<!--#include file="System.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台界面</title>
<style> 
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
.xiu_bt{
padding-left:20px;
font:"宋体",22px;
color:#66aa42;}
.xiu2{padding-left:150px;
padding-top:50px;
	}
.wramp{background-color:#f1f1f1;}
.red{color:red;}
.zhti{text-align:center;
margin-top:20px;
magin-left:20px;
border-collapse:separate;
}
.th{font-weight:bold;}
.bot{margin-top:10px;
}
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
<link type="text/css" rel="stylesheet" href="css.css" />
</head>
<body>
<!--头部-->
<div class="wramp">
  <div class="head">
  <div class="headce">
      <div class="logo"><img src="images/logo.jpg" /></div>
	  <div class="logo_right">
	      <div class="bt">
		      <div class="bt_left"><img src="images/bt.jpg" /></div>
			  <div class="bt_right">
			     您好，admin,欢迎您登录管理后台！    【退出】   2012年2月6日
			  </div>
			  <div class="clear"></div>
		  </div>
		  <div class="dh">
		     <ul>
			    <li><a href="index.asp"><span class="spec">后台首页</span></a></li>
				<li><a href="#">SEO系统</a> </li>
				<li><a href="#">新闻系统</a></li>
				<li><a href="#">图文系统</a></li>
				<li><a href="#">单页管理</a></li>
				<li><a href="#">在线系统</a></li>
				<li><a href="#">会员系统</a></li>
				<li><a href="#">帮助中心</a></li>
				
			 </ul>
		  </div>
	  </div><div class="clear"></div>
	  </div>
  </div>
  <!--左侧-->
  <div id="content">
  
 <!--#include file="left.asp"-->
 <div class="c-right">
 <p class="xiu_bt">查看所有管理员</p>
 <table width="600" border="1" cellpadding="0" align="center" class="zhti" cellspacing="0" >
  
  <tr class="th">
    <td>选择</td>
    <td>用户名</td>
    <td>密码</td>
    <td>会员类型</td>
    <td>操作</td>
  </tr>
  <tr>
    <td><input name="" type="checkbox" value="" /></td>
    <td>yz360</td>
    <td>e830393998438189</td>
    <td>超级管理员</td>
    <td>删除</td>
  </tr>
  <tr>
    <td><input name="" type="checkbox" value="" /></td>
    <td>yz360</td>
    <td>e830393998438189</td>
    <td>超级管理员</td>
    <td>删除</td>
  </tr>
  </table>
  <p class="bot"><input name="" type="checkbox" value="">全选</input>&nbsp;<input name="" type="button" value="删除选择" />
       共有：2条 每页显示：20条   [首页] [上页] [次页] [尾页]   页次：1/1页 转到：<select name=""></select>
 </p>
 
 
 
 </div>
 <div class="clear"></div>
  </div>
</div>
</body>
</html>
