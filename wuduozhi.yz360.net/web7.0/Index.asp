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
    <!--#include file="head.asp"-->
  <!--���-->

  <div class="clear"></div>
  
  
  <div id="content">
  
 <!--#include file="left.asp"-->
	
	<div class="c-right">
		  <!--#include file="baidu.asp"-->
		<div class="c-rgt02">
			<div class="rgt02-left">
			<img src="images/right-hi.jpg" />
			</div>
			
			<div class="rgt02-right">
			<p>��ϲ��ɹ������̨��ף�����춼�к����飡</span></p>
            <p>������Ա�վ���й��༭�������뼰ʱ�����������վ��Ϣ��</p>
			</div>
		  <div class="tianqi">
		    <iframe name="weather_inc" style="background-color:#f1f1f1" allowtransparency="true" src="http://cache.xixik.com.cn/8/nanyang/"  width="225" height="80" frameborder="0" marginwidth="0" marginheight="0" scrolling="no"></iframe>
		  </div>
		</div>
		<div class="c-rgt03">
			<p class="rgt03-bt">��������Ϣͳ��</p>
			<div class="ul-bg">
				<div class="rgt03-left">
				<ul>
				<li>���������ͣ�<%=Request.ServerVariables("OS")&"(IP:"&Request.ServerVariables("LOCAL_ADDR")&")"%></li>
				<li>վ������·����<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></li>
				<li>FSO�ı���д��<%If Not IsObjInstalled("Scripting.FileSystemObject") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></li>
				<li>Jmail�������֧�֣�<%If Not IsObjInstalled("JMail.SMTPMail") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%> </li>
				</ul>
			  </div>
				<div class="rgt03-right">
				<ul>
				<li>�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></li>
				<li>���ݿ��ַ��<%If IsSqlDataBase <> 1 Then%><%=Server.MapPath(MyDbPath&DB)%><%Else%>MsSQL���ݿ�<%End If%></li>
				<li>���ݿ�ʹ�ã�<%If Not IsObjInstalled("Adodb.Connection") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></li>
				<li>CDONTS�������֧�֣�<%If Not IsObjInstalled("CDONTS.NewMail") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></li>
				</ul>
			  </div>
			</div>
		</div>
		   <!--#include file="gongju.asp" -->
		<div class="bottom">
		<p>��Ȩ����:����������Ϣ�������޹�˾ ��ַ:�����б�����·88�ű�����Է4-4-5F ��������:473000</p>
		</div>
  </div>
	
	
	
  </div>
</div>
</body>
</html>
