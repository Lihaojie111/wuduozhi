<!--#include file="System.asp" -->
<%
response.Write(request("ID2"))
response.Write(request("Pic"))
if request("ID2")<>"" then 
set Rs5=server.CreateObject("adodb.recordset")
Sql="Select * From [Users] Where ID="&Request("ID2")
	Rs5.Open Sql,Conn,1,2
	if request("Pic")<>"" then
		  Rs5("Pic")=request("Pic")
		end if
		Rs5.Update  
		
		response.Redirect("members.asp?id="&Rs5("id")&"")
		Rs5.Close
		end if
		
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>南阳一世情缘婚介</title>
<style type="text/css">
<!--
@import url("images/css.css");
@import url("images/css-div.css");
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-image: url(images/bg.jpg);
}
.syswidth02 {	WIDTH: 70px
}
.STYLE2 {font-size: 12px}
.STYLE3 {color: #CC0066}
-->
</style>


<script charset="utf-8" src="./kindeditor/kindeditor.js"></script>
<script charset="utf-8" src="./kindeditor/lang/zh_CN.js"></script>
<script charset="utf-8" src="./kindeditor/plugins/code/prettify.js"></script>
<script>     KindEditor.ready(function (K) {
		 var editor1 = K.create('#content1', {             
		 cssPath: '../kindeditor/plugins/code/prettify.css',             
		 uploadJson: '../kindeditor/asp/upload_json.asp',            
		  fileManagerJson: '../kindeditor/aspt/file_manager_json.asp',            
		   allowFileManager: true,             
		   afterCreate: function () {                 
		   var self = this;                 
		   K.ctrl(document, 13, function () {                    
			self.sync();                     
			K('form[name=example]')[0].submit();                 
			});                 
			K.ctrl(self.edit.doc, 13, function () {                    
			 self.sync();                    
			  K('form[name=example]')[0].submit();                
			   });             
			   }         
			   });        
				prettyPrint();     
				}); 
				
				
				KindEditor.ready(function(K) {
				var editor = K.editor({
					allowFileManager : true
				});
				
						K('#image1').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							imageUrl : K('#url1').val(),
							clickFn : function(url, title, width, height, border, align) {
								K('#url1').val(url);
								editor.hideDialog();
							}
						});
					});
				});	
				});
          </script>

  <script language="javascript">
   function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//以下style=coolblue,值可以依据实际需要修改为您的样式名,通过此样式的后台设置来达到控制允许上传文件类型及文件大小
	var arr = showModalDialog("Ok_Editor/dialog/i_upload.htm?style=Cases&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}    
       
  

	</script>
</head>

<body>
<table width="998" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="32" colspan="3"><!--#include file="head01.asp" --></td>
  </tr>
  <tr>
    <td width="12">&nbsp;</td>
    <td width="974" align="center"><table width="980" align="center" cellspacing="0">
      <tr>
        <td width="284"><!--#include file="left-hy.asp" --></td>
        <td width="690" valign="top"><table width="647" height="591" cellspacing="0">
          <tr>
            <td width="643" height="32" align="left" background="images/hyx.jpg"><table width="643" cellspacing="0">
              <tr>
                <td width="26"><img src="images/mlp1.jpg" width="21" height="19" /></td>
                <td width="318" align="left"><span class="css14 STYLE3">我的头像</span></td>
                <td width="291" align="right" class="css"><a href="memebers.asp?id=<%=session("id")%>" class="hui1">返回会员中心&gt;&gt;</a></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="51" valign="middle"><table width="245" height="37" cellspacing="0">
              <tr>
                <td height="35" align="center" background="images/member_right_title.jpg" class="mei12">在线编辑头像 </td>
              </tr>
            </table></td>
          </tr>

          <tr>
            <td height="32" align="left" background="images/hy2.jpg" class="bai12"><table width="205" cellspacing="0">
              <tr>
                <td width="18">&nbsp;</td>
                <td width="181" align="left"><span class="hui14xx">我的头像</span></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="155" align="left" valign="baseline"><table width="117" cellspacing="0">
              <tr>
                <td width="100">
		<% 
								SQLSelect="select  * from [Users] where ID="&request("id")
Rs.Open SQLSelect,Conn,1,1
idy=Rs("id")
if Rs("Pic")<>"" then%><img  src="<%=Rs("Pic")%>" width="100" height="125" /><%elseif Rs("Gender")="0" then%><img src="images/default_woman.jpg"  width="100" height="125" /><%else%><img src="images/default_man.jpg"  width="100" height="125" /><%end if
Rs.close%></td>
                <td width="11">&nbsp;</td>
              </tr>
              <tr>
                <td align="center" class="css">当前头像</td>
                <td>&nbsp;</td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="275" align="center" background="images/member_right_pic_nav.jpg">&nbsp;</td>
          </tr>
           <tr>
                      <td colspan="2" align="left" >请上传头像：
                        <form name="myform" method="post" action="UploadPic.asp">
                        <input type="text" id="url1" value="" name="pic"  /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）<input name="ID2" type="hidden" value="<%=idy%>">
<label>
<input type="submit" name="button" id="button" value="提交" />
</label>


                        </form></td>
              </tr>
                  
        </table></td>
      </tr>
    </table></td>
    <td width="10">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td valign="top"><table id="__01" width="972" height="99" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <!--#include file="foot.asp" -->
      </tr>
</table>  
      </td>
    <td>&nbsp;</td>
  </tr>
</table>

</body>
</html>
