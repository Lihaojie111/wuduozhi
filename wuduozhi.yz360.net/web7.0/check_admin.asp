<!--#include file="System.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����</title>
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
    font-family:"΢���ź�";} 
.top_menu a:hover{background:url(left_bg1.jpg);}
.sub_meuu a{ font-size:12px; color:#FFF;
	display:block;
	width:164px;
	padding-left:64px;
	height:36px;
	line-height:36px;
	font-family:"΢���ź�";
} 
.hover:hover{background:url(bg3.jpg);
display:block;
color:#dee318;}
.xiu_bt{
padding-left:20px;
font:"����",22px;
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
<!--ͷ��-->
<div class="wramp">
  <div class="head">
  <div class="headce">
      <div class="logo"><img src="images/logo.jpg" /></div>
	  <div class="logo_right">
	      <div class="bt">
		      <div class="bt_left"><img src="images/bt.jpg" /></div>
			  <div class="bt_right">
			     ���ã�admin,��ӭ����¼�����̨��    ���˳���   2012��2��6��
			  </div>
			  <div class="clear"></div>
		  </div>
		  <div class="dh">
		     <ul>
			    <li><a href="index.asp"><span class="spec">��̨��ҳ</span></a></li>
				<li><a href="#">SEOϵͳ</a> </li>
				<li><a href="#">����ϵͳ</a></li>
				<li><a href="#">ͼ��ϵͳ</a></li>
				<li><a href="#">��ҳ����</a></li>
				<li><a href="#">����ϵͳ</a></li>
				<li><a href="#">��Աϵͳ</a></li>
				<li><a href="#">��������</a></li>
				
			 </ul>
		  </div>
	  </div><div class="clear"></div>
	  </div>
  </div>
  <!--���-->
  <div id="content">
  
 <!--#include file="left.asp"-->
 <div class="c-right">
 <p class="xiu_bt">�鿴���й���Ա</p>
 <table width="600" border="1" cellpadding="0" align="center" class="zhti" cellspacing="0" >
  
  <tr class="th">
    <td>ѡ��</td>
    <td>�û���</td>
    <td>����</td>
    <td>��Ա����</td>
    <td>����</td>
  </tr>
  <tr>
    <td><input name="" type="checkbox" value="" /></td>
    <td>yz360</td>
    <td>e830393998438189</td>
    <td>��������Ա</td>
    <td>ɾ��</td>
  </tr>
  <tr>
    <td><input name="" type="checkbox" value="" /></td>
    <td>yz360</td>
    <td>e830393998438189</td>
    <td>��������Ա</td>
    <td>ɾ��</td>
  </tr>
  </table>
  <p class="bot"><input name="" type="checkbox" value="">ȫѡ</input>&nbsp;<input name="" type="button" value="ɾ��ѡ��" />
       ���У�2�� ÿҳ��ʾ��20��   [��ҳ] [��ҳ] [��ҳ] [βҳ]   ҳ�Σ�1/1ҳ ת����<select name=""></select>
 </p>
 
 
 
 </div>
 <div class="clear"></div>
  </div>
</div>
</body>
</html>
