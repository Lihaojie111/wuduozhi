<!--#include file="System.asp"-->
<!--#include file="inc/Md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����һ����Ե���</title>
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
.STYLE1 {color: #333333}
.syswidth02 {	WIDTH: 70px
}
.STYLE2 {font-size: 12px}
-->
</style>

<script src="js/calendar.js" type="text/javascript"></script>
<script type="text/javascript" src="js/Common.js"></script>

    <script language="javascript" src="js/area.js"></script>

    <script type="text/javascript" src="js/PCASClass.js" charset="gb2312" defer="defer"></script>

    <script language="javascript" src="js/menu.js"></script>
    <link type="text/css" rel="stylesheet" href="css.css" />

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
		  function SetSignbtnState()
   {
      var cb_Agree = document.getElementById('cb_Agree');
      if( cb_Agree.checked)
      {
           document.getElementById('button').disabled = false;
      }
      else
      {
           document.getElementById('button').disabled = true;
      }
   }
   
	  function SetArea(ProvinceID, CityID, AreaID)
    {
        var objProvince = document.getElementById(ProvinceID);
        var objCity = document.getElementById(CityID);        
        var ProvinceText = objProvince.options[objProvince.selectedIndex].text;
        var CityText = objCity.options[objCity.selectedIndex].text;
        var objArea = document.getElementById(AreaID);
        var AreaText = objArea.options[objArea.selectedIndex].text;
        
        if (arguments.length > 3)
        {
            if(ProvinceText != "ʡ" && ProvinceText != "����")
            {
                document.getElementById("hidden_YqProvince").value = ProvinceText;
                if(CityText != "==����==" && CityText != "����")
                {
                    document.getElementById("hidden_YqCity").value = CityText;
                    if (AreaText != "==����==" && AreaText != "==��ѡ��==" && AreaText != "��" && AreaText != "����")
                        document.getElementById("hidden_YqArea").value = AreaText;
                    if (AreaText == "==����==" || AreaText == "==��ѡ��==" || AreaText == "��" && AreaText != "����")
                        document.getElementById("hidden_YqArea").value = "";
                }
                
                else
                {
                    document.getElementById("hidden_YqCity").value = "";
                    document.getElementById("hidden_YqArea").value = "";
                }
            }       
        }
        else
        {
            var objArea = document.getElementById(AreaID);
            var AreaText = objArea.options[objArea.selectedIndex].text;
            if(ProvinceText != "ʡ" && ProvinceText != "����")
            {
                
                document.getElementById("hidden_Province").value = ProvinceText;
                if(CityText != "==����==" && CityText != "����")
                {
                    document.getElementById("hidden_City").value = CityText;
                    if (AreaText != "==����==" && AreaText != "==��ѡ��==" && AreaText != "��" && AreaText != "����")
                        document.getElementById("hidden_Area").value = AreaText;
                    if (AreaText == "==����==" || AreaText == "==��ѡ��==" || AreaText == "��" && AreaText != "����")
                        document.getElementById("hidden_Area").value = "";
                }
                
                else
                {
                    document.getElementById("hidden_City").value = "";
                    document.getElementById("hidden_Area").value = "";
                }
            }
        }
    }
    
	 function signs(birthday) 
    {        
        if (birthday != "" & birthday.length > 9)
        {
            var datenow = new Date();
            var birthArr = birthday.split("-");
            if (birthArr.length > 0)
            {
                var start = 1901, birthyear = birthArr[0], date=parseInt(birthArr[2]), month=(birthArr[1]);
                var shuxiang, star;
                
                if (month == 1 && date >=20 || month == 2 && date <=18) {star = "ˮƿ��";}
                if (month == 1 && date > 31) {value = "Huh?";}
                if (month == 2 && date >=19 || month == 3 && date <=20) {star = "˫����";}
                if (month == 2 && date > 29) {value = "Say what?";}
                if (month == 3 && date >=21 || month == 4 && date <=19) {star = "������";}
                if (month == 3 && date > 31) {value = "OK.  Whatever.";}
                if (month == 4 && date >=20 || month == 5 && date <=20) {star = "��ţ��";}
                if (month == 4 && date > 30) {value = "I'm soooo sorry!";}
                if (month == 5 && date >=21 || month == 6 && date <=21) {star = "˫����";}
                if (month == 5 && date > 31) {value = "Umm ... no.";}
                if (month == 6 && date >=22 || month == 7 && date <=22) {star = "��з��";}
                if (month == 6 && date > 30) {value = "Sorry.";}
                if (month == 7 && date >=23 || month == 8 && date <=22) {star = "ʨ����";}
                if (month == 7 && date > 31) {value = "Excuse me?";}
                if (month == 8 && date >=23 || month == 9 && date <=22) {star = "��Ů��";}
                if (month == 8 && date > 31) {value = "Yeah. Right.";}
                if (month == 9 && date >=23 || month == 10 && date <=22) {star = "�����";}
                if (month == 9 && date > 30) {value = "Try Again.";}
                if (month == 10 && date >=23 || month == 11 && date <=21) {star = "��Ы��";}
                if (month == 10 && date > 31) {value = "Forget it!";}
                if (month == 11 && date >=22 || month == 12 && date <=21) {star = "������";}
                if (month == 11 && date > 30) {value = "Invalid Date";}
                if (month == 12 && date >=22 || month == 1 && date <=19) {star = "Ħ����";}
                if (month == 12 && date > 31) {value = "No way!";}
                
                x = (start - birthyear) % 12
                
                if (x == 1 || x == -11) {shuxiang = "��";}
                if (x == 0) {shuxiang = "ţ";}
                if (x == 11 || x == -1) {shuxiang = "��";}
                if (x == 10 || x == -2) {shuxiang = "��";}
                if (x == 9 || x == -3)  {shuxiang = "��";}
                if (x == 8 || x == -4)  {shuxiang ="��";}
                if (x == 7 || x == -5)  {shuxiang = "��";}
                if (x == 6 || x == -6)  {shuxiang = "��";}
                if (x == 5 || x == -7)  {shuxiang = "��";}
                if (x == 4 || x == -8)  {shuxiang = "��";}
                if (x == 3 || x == -9)  {shuxiang = "��";}
                if (x == 2 || x == -10)  {shuxiang = "��";}
                
                document.getElementById("txb_Shuxiang").value = shuxiang;
                document.getElementById("txb_star").value = star;
            }
        }
        
        else
        {
            document.getElementById("txb_Shuxiang").value = "";
            document.getElementById("txb_star").value = "";
        }        
    }
	   function limitChars(id, count){   
        var obj = document.getElementById(id);   
        if (obj.value.length > count){   
            obj.value = obj.value.substr(0, count);   
        }   
    }  

    
    

    
    String.prototype.trim= function(){  
        // ��������ʽ��ǰ��ո�  
        // �ÿ��ַ��������  
        return this.replace(/(^\s*)|(\s*$)/g, "");  
    } 
	
   function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//����style=coolblue,ֵ��������ʵ����Ҫ�޸�Ϊ������ʽ��,ͨ������ʽ�ĺ�̨�������ﵽ���������ϴ��ļ����ͼ��ļ���С
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
    <td width="5">&nbsp;</td>
    <td width="986"><table width="983" cellspacing="0">
      <tr>
        <td width="640"><form name="myform" method="post" action="register.asp" >
        <table width="674" border="0" align="left" cellpadding="0" cellspacing="0" id="__01">
          <tr>
            <td width="674" height="34" colspan="4" background="images/g_011.png"><table width="653" height="19" cellspacing="0">
                <tr>
                  <td width="34">&nbsp;</td>
                  <td width="476" class="mei12">�û���¼</td>
                  <td width="135" class="bai12">&nbsp;</td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td colspan="4" valign="top" background="images/g_031.png"><div>
              <div class="member_register_1_title"><span>�û���¼��Ϣ�����ڵ�¼��վʱ����ʹ�ã�</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table cellspacing="5" width="600" border="0">
                                            <tr style="height: 65px">
                                                <td width="150" height="0" align="right">
                                                    ������¼������ ��</td>
                                                <td width="180">
                                                    <input name="txb_email" type="text" id="txb_email" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" /><br />
                                                    <br />
                                                  
                                                </td>
                                                <td>
                                                    <span id="showEmailInfo"><span style="color: #ff0000">*</span> ����ʵ�ʼ����ղ����Է����������ʼ�</span></td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    �����û������� ��</td>
                                              <td>
                                                <input name="txb_PassWord" type="password" maxlength="20" id="txb_PassWord" class="input2" style="height: 17px;
                                                        background-color: #FFF; border: 1px solid #ffcdd2;" /><br /></td>
                                                <td>
                                                    <span style="color: #ff0000">*</span> �������û�������</td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    ȷ������ ��</td>
                                                <td>
                                                    <input name="txb_rePassword" type="password" maxlength="20" id="txb_rePassword" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                                <td>
                                              </td>
                                            </tr>
                                        </table>
                </P>
              </div>
              <div 
class="member_register_1_title"><span>���Ᵽ�ܵĻ�����Ϣ���ϣ���������д����Щ����ֻ�����͹���Ա���ܿ�����</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table width="650" border="0" cellspacing="5">
                                            <tr style="height: 75px">
                                                <td width="88" height="0" align="right">
                                                    ��ʵ���� ��</td>
                                                <td width="211">
                                                    <input name="txb_RealName" type="text" id="txb_RealName" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                              </td>
                                                <td width="101" align="right">
                                                    �ҵ��ǳ� ��</td>
                                                <td width="217">
                                                    <input name="txb_NickName" type="text" id="txb_NickName" class="input2" onblur="checkNickName(this.value)" style="height: 17px; width: 150px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                    <span style="color: #ff0000">*</span><br />
                                                    <span id="nickname"></span>
                                              </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    ��ϵ�绰 ��</td>
                                                <td><input name="txb_TelePhone" type="text" id="txb_TelePhone" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />                                                  <span style="color: #ff0000">*</span><br />
                                                </td>
                                                <td align="right">
                                                    Q Q���� ��
                                                </td>
                                                <td>
                                                    <input name="txb_QQ" type="text" id="txb_QQ" class="input2" style="height: 17px; width: 150px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    ���֤�� ��</td>
                                                <td>
                                                    <input name="txb_Identity" type="text" id="txb_Identity" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" /></td>
                                                <td align="right">
                                                    ��ҵԺУ ��</td>
                                                <td>
                                                    <input name="txb_FinishSchool" type="text" id="txb_FinishSchool" class="input2" style="height: 17px; width: 150px;
                                                        background-color: #FFF; border: 1px solid #ffcdd2;" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    ������λ ��</td>
                                                <td>
                                                    <input name="txb_WorkCompany" type="text" id="txb_WorkCompany" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                                <td align="right">
                                                    ��ͥ���� ��</td>
                                                <td>
                                                    <select name="ddl_Family" id="ddl_Family" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2; color: #e65479; height: 22px;">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">��ͨ��ͥ ������Ů</option>
	<option value="2">��ͨ��ͥ �Ƕ�����Ů</option>
	<option value="3">�в���ͥ ������Ů</option>
	<option value="4">�в���ͥ �Ƕ�����Ů</option>
	<option value="5">�ߵȼ�ͥ ������Ů</option>
	<option value="6">�ߵȼ�ͥ �Ƕ�����Ů</option>

</select>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    ����֤�� ��</td>
                                                <td>
                                                    <input name="txb_LiYiNum" type="text" id="txb_LiYiNum" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                                <td align="right">
                                                </td>
                                                <td>
                                                </td>
                                            </tr>
                                        </table>
                </P>
              </div>
              <div 
class="member_register_1_title"><span>���⹫���Ļ�����Ϣ���ϣ���������д���⽫�������������ɹ���</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table width="600" border="0" cellspacing="5" onload="AddList('age1', 'age2'); new PCAS('province','city');" onclick="signs(document.getElementById('txb_Birthday').value)" >
                                            <tr>
                                                <td width="82" style="text-align: right">
                                                    ����״�� ��</td>
                                                <td width="208">
                                                    <select name="ddl_Marry" id="ddl_Marry">
	<option value="">��ѡ��</option>
	<option value="1">δ��</option>
	<option value="2">������к�</option>
	<option value="3">�����Ů��</option>
	<option value="4">����δ��</option>
	<option value="5">�����к���Է�</option>
	<option value="12">����Ů����Է�</option>
	<option value="6">ɥż���к�</option>
	<option value="7">ɥż��Ů��</option>
	<option value="8">ɥżδ��</option>
	<option value="9">�̻�δ��</option>
	<option value="10">ʵ�»������к�</option>
	<option value="11">ʵ�»�����Ů��</option>
</select>
                                                    <span style="color: #ff0000">*</span><br />                                              </td>
                                                <td width="104" style="text-align: right">
                                                    ��&nbsp; &nbsp;�� ��</td>
                                                <td width="173">
                                                    <select name="rbtn_Gender" id="rbtn_Gender">
	<option selected="selected" value="1">��</option>
	<option value="0">Ů</option>
</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    �������� ��</td>
                                                <td>
                                                    <input name="txb_BirthDay" type="text" id="txb_BirthDay" class="input2" onblur="signs(this.value)" onchange="signs(this.value)" onfocus="calendar()" style="height: 17px; background-color: #FFF; border: 1px solid #ffcdd2;" />
                                                    <span style="color: #ff0000">*</span><br />
                                                </td>
                                                <td style="text-align: right">
                                                    ��&nbsp; &nbsp;�� ��</td>
                                                <td>
                                                    <input name="txb_Shuxiang" type="text" id="txb_Shuxiang" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ��&nbsp; &nbsp;�� ��</td>
                                                <td>
                                                    <input name="txb_star" type="text" id="txb_star" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />                                                </td>
                                              <td style="text-align: right">
                                              </td>
                                                <td>
                                              </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ѧ&nbsp; &nbsp;�� ��</td>
                                                <td><select name="ddl_Education" id="ddl_Education">
                                                  <option selected="selected" value="">��ѡ��</option>
                                                  <option value="1">û��ѧ��</option>
                                                  <option value="2">Сѧ</option>
                                                  <option value="3">����</option>
                                                  <option value="4">����</option>
                                                  <option value="5">��ר</option>
                                                  <option value="6">��ר</option>
                                                  <option value="7">����</option>
                                                  <option value="8">�о���</option>
                                                  <option value="9">��ʿ</option>
                                                </select>
                                                <span style="color: #ff0000">*</span>                                                </td>
                                                <td style="text-align: right">
                                                    ר&nbsp; &nbsp;ҵ ��</td>
                                                <td><select name="ddl_Specialty" id="ddl_Specialty">
                                                  <option selected="selected" value="">��ѡ��</option>
                                                  <option value="1">����</option>
                                                  <option value="2">�Ŀ���</option>
                                                  <option value="3">ҽѧ��</option>
                                                  <option value="4">ũѧ��</option>
                                                  <option value="5">������</option>
                                                  <option value="6">������</option>
                                                  <option value="7">������</option>
                                                  <option value="8">������</option>
                                                  <option value="9">û��רҵ</option>
                                                </select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    Ѫ&nbsp;&nbsp;�� ��</td>
                                                <td>
                                                    <select name="ddl_BloodType" id="ddl_BloodType">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">A</option>
	<option value="2">B</option>
	<option value="3">AB</option>
	<option value="4">O</option>
	<option value="5">����</option>
	<option value="6">��֪��</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    ��&nbsp;&nbsp;�� ��</td>
                                                <td>
                                                    <select name="ddl_Minzu" id="ddl_Minzu">
	<option value="">��ѡ��</option>
	<option value="1">����</option>
	<option value="2">׳��</option>
	<option value="3">����</option>
	<option value="4">����</option>
	<option value="5">�ɹ���</option>
	<option value="6">����</option>
	<option value="7">������</option>
	<option value="8">ά�����</option>
	<option value="9">����</option>
	<option value="10">����</option>
	<option value="11">��������</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ��ߣ�cm����</td>
                                                <td>
                                                    <select name="ddl_Height" id="ddl_Height">
	<option value="0">��ѡ��</option>
	<option value="145">145</option>
	<option value="146">146</option>
	<option value="147">147</option>
	<option value="148">148</option>
	<option value="149">149</option>
	<option value="150">150</option>
	<option value="151">151</option>
	<option value="152">152</option>
	<option value="153">153</option>
	<option value="154">154</option>
	<option value="155">155</option>
	<option value="156">156</option>
	<option value="157">157</option>
	<option value="158">158</option>
	<option value="159">159</option>
	<option value="160">160</option>
	<option value="161">161</option>
	<option value="162">162</option>
	<option value="163">163</option>
	<option value="164">164</option>
	<option value="165">165</option>
	<option value="166">166</option>
	<option value="167">167</option>
	<option value="168">168</option>
	<option value="169">169</option>
	<option value="170">170</option>
	<option value="171">171</option>
	<option value="172">172</option>
	<option value="173">173</option>
	<option value="174">174</option>
	<option value="175">175</option>
	<option value="176">176</option>
	<option value="177">177</option>
	<option value="178">178</option>
	<option value="179">179</option>
	<option value="180">180</option>
	<option value="181">181</option>
	<option value="182">182</option>
	<option value="183">183</option>
	<option value="184">184</option>
	<option value="185">185</option>
	<option value="186">186</option>
	<option value="187">187</option>
	<option value="188">188</option>
	<option value="189">189</option>
	<option value="190">190</option>
	<option value="191">191</option>
	<option value="192">192</option>
	<option value="193">193</option>
	<option value="194">194</option>
	<option value="195">195</option>
	<option value="196">196</option>
	<option value="197">197</option>
	<option value="198">198</option>
	<option value="199">199</option>
	<option value="200">200</option>
	<option value="201">201</option>
	<option value="202">202</option>
	<option value="203">203</option>
	<option value="204">204</option>
	<option value="205">205</option>
	<option value="206">206</option>
	<option value="207">207</option>
	<option value="208">208</option>
	<option value="209">209</option>
	<option value="210">210</option>
	<option value="211">211</option>
	<option value="212">212</option>
	<option value="213">213</option>
	<option value="214">214</option>
	<option value="215">215</option>
	<option value="216">216</option>
	<option value="217">217</option>
	<option value="218">218</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    ���أ�kg����</td>
                                                <td>
                                                    <select name="ddl_weight" id="ddl_weight">
	<option value="0">��ѡ��</option>
	<option value="35">35</option>
	<option value="36">36</option>
	<option value="37">37</option>
	<option value="38">38</option>
	<option value="39">39</option>
	<option value="40">40</option>
	<option value="41">41</option>
	<option value="42">42</option>
	<option value="43">43</option>
	<option value="44">44</option>
	<option value="45">45</option>
	<option value="46">46</option>
	<option value="47">47</option>
	<option value="48">48</option>
	<option value="49">49</option>
	<option value="50">50</option>
	<option value="51">51</option>
	<option value="52">52</option>
	<option value="53">53</option>
	<option value="54">54</option>
	<option value="55">55</option>
	<option value="56">56</option>
	<option value="57">57</option>
	<option value="58">58</option>
	<option value="59">59</option>
	<option value="60">60</option>
	<option value="61">61</option>
	<option value="62">62</option>
	<option value="63">63</option>
	<option value="64">64</option>
	<option value="65">65</option>
	<option value="66">66</option>
	<option value="67">67</option>
	<option value="68">68</option>
	<option value="69">69</option>
	<option value="70">70</option>
	<option value="71">71</option>
	<option value="72">72</option>
	<option value="73">73</option>
	<option value="74">74</option>
	<option value="75">75</option>
	<option value="76">76</option>
	<option value="77">77</option>
	<option value="78">78</option>
	<option value="79">79</option>
	<option value="80">80</option>
	<option value="81">81</option>
	<option value="82">82</option>
	<option value="83">83</option>
	<option value="84">84</option>
	<option value="85">85</option>
	<option value="86">86</option>
	<option value="87">87</option>
	<option value="88">88</option>
	<option value="89">89</option>
	<option value="90">90</option>
	<option value="91">91</option>
	<option value="92">92</option>
	<option value="93">93</option>
	<option value="94">94</option>
	<option value="95">95</option>
	<option value="96">96</option>
	<option value="97">97</option>
	<option value="98">98</option>
	<option value="99">99</option>
	<option value="100">100</option>
	<option value="101">101</option>
	<option value="102">102</option>
	<option value="103">103</option>
	<option value="104">104</option>
	<option value="105">105</option>
	<option value="106">106</option>
	<option value="107">107</option>
	<option value="108">108</option>
	<option value="109">109</option>
	<option value="110">110</option>
	<option value="111">111</option>
	<option value="112">112</option>
	<option value="113">113</option>
	<option value="114">114</option>
	<option value="115">115</option>
	<option value="116">116</option>
	<option value="117">117</option>
	<option value="118">118</option>
	<option value="119">119</option>
	<option value="120">120</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ��&nbsp; &nbsp;�� ��</td>
                                                <td colspan="3">
                                                    <select name="mprovince" id="mprovince" style="width: 100px" onchange="javascript:InitMenu2(document.getElementById('mprovince'),document.getElementById('mcity'),document.getElementById('mcountry'),0);SetArea('mprovince','mcity','mcountry')">
                                                    </select>
                                                    <select name="mcity" id="mcity" onchange="InitMenu3(document.getElementById('mprovince'),document.getElementById('mcity'), document.getElementById('mcountry'), 0);SetArea('mprovince','mcity','mcountry')">
                                                    </select>
                                                    <select name="mcountry" id="mcountry" onchange="SetArea('mprovince','mcity','mcountry')">
                                                    </select>
                                                    <span style="color: #ff0000">*</span>

                                                  <script language="javascript">
											        InitMenu1(document.getElementById("mprovince"),document.getElementById("mcity"),document.getElementById("mcountry"));
											        InitMenu2(document.getElementById("mprovince"),document.getElementById("mcity"),document.getElementById("mcountry"));
											        InitMenu3(document.getElementById("mprovince"),document.getElementById("mcity"),document.getElementById("mcountry"));
											        document.getElementById("mprovince").options[0].text="ʡ";
											        document.getElementById("mcity").options[0].text="��";
											        document.getElementById("mcountry").options[0].text="��";
                                                    </script>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ��λ���� ��</td>
                                                <td>
                                                    <select name="ddl_Danweixingzhi" id="ddl_Danweixingzhi">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">���һ���</option>
	<option value="2">��������</option>
	<option value="3">ҽ�ƻ���</option>
	<option value="4">����ҵ</option>
	<option value="5">����ҵ</option>
	<option value="6">��˾</option>
	<option value="7">����</option>
	<option value="8">���徭Ӫ</option>
	<option value="9">����</option>
	<option value="10">��Ӫ����</option>
	<option value="11">������ҵ</option>
	<option value="12">������ҵ</option>
	<option value="13">û�е�λ</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    ְҵְ�� ��</td>
                                                <td>
                                                    <select name="ddl_job" id="ddl_job">
	<option value="">��ѡ��</option>
	<option value="1">����Ա</option>
	<option value="2">���Ҹɲ�</option>
	<option value="3">��ʦ</option>
	<option value="4">��ѧ��</option>
	<option value="5">�˶�Ա</option>
	<option value="6">����ʦ</option>
	<option value="7">����</option>
	<option value="8">�ܾ���</option>
	<option value="9">IT��ҵ</option>
	<option value="10">����ҵ��</option>
	<option value="11">ְԱ</option>
	<option value="12">����</option>
	<option value="13">������</option>
	<option value="14">����</option>
	<option value="15">��Ա</option>
	<option value="16">����</option>
	<option value="17">ҽ��</option>
	<option value="18">��ʿ</option>
	<option value="19">����</option>
	<option value="20">ʿ��</option>
	<option value="21">��ʦ</option>
	<option value="22">����ְҵ</option>
	<option value="23">����</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ��&nbsp;��&nbsp;�룺</td>
                                                <td>
                                                    <select name="ddl_MonthIn" id="ddl_MonthIn">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">500Ԫ</option>
	<option value="2">1000Ԫ����</option>
	<option value="3">1000-2000Ԫ</option>
	<option value="4">2000-3000Ԫ</option>
	<option value="5">3000-4000Ԫ</option>
	<option value="6">4000-5000Ԫ</option>
	<option value="7">5000-6000Ԫ</option>
	<option value="8">6000-7000Ԫ</option>
	<option value="9">7000-8000Ԫ</option>
	<option value="10">8000-9000Ԫ</option>
	<option value="11">10000Ԫ����</option>
	<option value="12">20000Ԫ����</option>
	<option value="13">30000Ԫ����</option>
	<option value="14">40000Ԫ����</option>
	<option value="15">50000Ԫ����</option>
	<option value="16">60000Ԫ����</option>
	<option value="17">100000Ԫ����</option>
	<option value="18">150000Ԫ����</option>
	<option value="19">200000Ԫ����</option>
	<option value="20">300000Ԫ����</option>
	<option value="21">500000Ԫ����</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    ס����� ��</td>
                                                <td>
                                                    <select name="ddl_House" id="ddl_House">
	<option selected="selected" value="">��ѡ��</option>
	<option value="2">��ͨס��</option>
	<option value="3">�е�ס��</option>
	<option value="4">�ߵ�ס��</option>
	<option value="5">����ס��</option>
	<option value="6">����ס��</option>
	<option value="7">ũ��ס��</option>
	<option value="1">û��ס��</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ������� ��</td>
                                                <td>
                                                    <select name="ddl_Car" id="ddl_Car">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">�޳�</option>
	<option value="2">��ͨ�γ�</option>
	<option value="3">�м��γ�</option>
	<option value="4">�߼��γ�</option>
	<option value="5">�����γ�</option>
	<option value="6">��������</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    ��Ȥ���� ��</td>
                                                <td>
                                                    <select name="ddl_AiHao" id="ddl_AiHao">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">������ѧ</option>
	<option value="2">ҵ����</option>
	<option value="3">���ù㷺</option>
	<option value="4">�����˶�</option>
	<option value="5">��������</option>
	<option value="6">�˼ʽ���</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    �Ը����� ��</td>
                                                <td>
                                                    <select name="ddl_Xingge" id="ddl_Xingge">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">����</option>
	<option value="2">����</option>
	<option value="3">�������</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    �Ļ����� ��</td>
                                                <td>
                                                    <select name="ddl_Wenhua" id="ddl_Wenhua">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">һ��ˮ׼</option>
	<option value="2">�е�ˮ׼</option>
	<option value="3">�ߵ�ˮ׼</option>
	<option value="4">ר��ˮ׼</option>
</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    �ۺ����� ��</td>
                                                <td>
                                                    <select name="ddl_Zonghetezhi" id="ddl_Zonghetezhi">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">������</option>
	<option value="2">������</option>
	<option value="3">������������</option>
	<option value="4">������</option>
	<option value="5">������</option>
	<option value="6">�쵼��</option>
	<option value="7">ְԱ��</option>
	<option value="8">������</option>
</select></td>
                                                <td style="text-align: right">
                                                    �ر�˵�� ��</td>
                                                <td>
                                                    <select name="ddl_TeBieshuoming" id="ddl_TeBieshuoming">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">�ҵ�����������</option>
	<option value="2">���ڵ���Ե��</option>
	<option value="3">����������Ĳ���</option>
	<option value="4">��û��̸������</option>
	<option value="5">��û��ʲô˵��</option>
</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    �������� ��</td>
                                                <td>
                                                    <select name="ddl_Xingxiangsuzhi" id="ddl_Xingxiangsuzhi">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">�������ȥ</option>
	<option value="2">��������Ϻ�</option>
	<option value="3">�������ʽϸ�</option>
	<option value="4">���������</option>
	<option value="5">�������ʸ�</option>
	<option value="6">�������ʽϸ�</option>
	<option value="7">�������ʸ�</option>
</select></td>
                                                <td style="text-align: right">
                                                    �������� ��</td>
                                                <td>
                                                    <select name="ddl_Gerenzongshu" id="ddl_Gerenzongshu">
	<option value="">��ѡ��</option>
	<option value="1">����ò�ž��</option>
	<option value="2">����������</option>
	<option value="3">������Ů��</option>
	<option value="4">�������ڻ���</option>
	<option value="5">��û��ʲô˵��</option>
	<option value="6">���������Ұ�</option>
</select></td>
                                            </tr>
                                        </table>
                </P>
              </div>
              <div 
class="member_register_1_title"><span>��żҪ��(��������д���⽫����������׼ȷ��׽����ż������߻����ɹ��ʣ�</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table width="650" border="0" cellspacing="5">
                                            <tr>
                                                <td width="101" style="text-align: right">
                                                    ����Ҫ�� ��</td>
                                                <td width="209">
                                                    <select name="ddl_YqBeginAge" id="ddl_YqBeginAge">
	<option value="0">��ѡ��</option>
	<option value="18">18</option>
	<option value="19">19</option>
	<option value="20">20</option>
	<option value="21">21</option>
	<option value="22">22</option>
	<option value="23">23</option>
	<option value="24">24</option>
	<option value="25">25</option>
	<option value="26">26</option>
	<option value="27">27</option>
	<option value="28">28</option>
	<option value="29">29</option>
	<option value="30">30</option>
	<option value="31">31</option>
	<option value="32">32</option>
	<option value="33">33</option>
	<option value="34">34</option>
	<option value="35">35</option>
	<option value="36">36</option>
	<option value="37">37</option>
	<option value="38">38</option>
	<option value="39">39</option>
	<option value="40">40</option>
	<option value="41">41</option>
	<option value="42">42</option>
	<option value="43">43</option>
	<option value="44">44</option>
	<option value="45">45</option>
	<option value="46">46</option>
	<option value="47">47</option>
	<option value="48">48</option>
	<option value="49">49</option>
	<option value="50">50</option>
	<option value="51">51</option>
	<option value="52">52</option>
	<option value="53">53</option>
	<option value="54">54</option>
	<option value="55">55</option>
	<option value="56">56</option>
	<option value="57">57</option>
	<option value="58">58</option>
	<option value="59">59</option>
	<option value="60">60</option>
	<option value="61">61</option>
	<option value="62">62</option>
	<option value="63">63</option>
	<option value="64">64</option>
	<option value="65">65</option>
	<option value="66">66</option>
	<option value="67">67</option>
	<option value="68">68</option>
	<option value="69">69</option>
	<option value="70">70</option>
	<option value="71">71</option>
	<option value="72">72</option>
	<option value="73">73</option>
	<option value="74">74</option>
	<option value="75">75</option>
	<option value="76">76</option>
	<option value="77">77</option>
	<option value="78">78</option>
	<option value="79">79</option>
	<option value="80">80</option>

</select>�� ��
                                                    <select name="ddl_YqEndAge" id="ddl_YqEndAge">
	<option value="90">��ѡ��</option>
	<option value="18">18</option>
	<option value="19">19</option>
	<option value="20">20</option>
	<option value="21">21</option>
	<option value="22">22</option>
	<option value="23">23</option>
	<option value="24">24</option>
	<option value="25">25</option>
	<option value="26">26</option>
	<option value="27">27</option>
	<option value="28">28</option>
	<option value="29">29</option>
	<option value="30">30</option>
	<option value="31">31</option>
	<option value="32">32</option>
	<option value="33">33</option>
	<option value="34">34</option>
	<option value="35">35</option>
	<option value="36">36</option>
	<option value="37">37</option>
	<option value="38">38</option>
	<option value="39">39</option>
	<option value="40">40</option>
	<option value="41">41</option>
	<option value="42">42</option>
	<option value="43">43</option>
	<option value="44">44</option>
	<option value="45">45</option>
	<option value="46">46</option>
	<option value="47">47</option>
	<option value="48">48</option>
	<option value="49">49</option>
	<option value="50">50</option>
	<option value="51">51</option>
	<option value="52">52</option>
	<option value="53">53</option>
	<option value="54">54</option>
	<option value="55">55</option>
	<option value="56">56</option>
	<option value="57">57</option>
	<option value="58">58</option>
	<option value="59">59</option>
	<option value="60">60</option>
	<option value="61">61</option>
	<option value="62">62</option>
	<option value="63">63</option>
	<option value="64">64</option>
	<option value="65">65</option>
	<option value="66">66</option>
	<option value="67">67</option>
	<option value="68">68</option>
	<option value="69">69</option>
	<option value="70">70</option>
	<option value="71">71</option>
	<option value="72">72</option>
	<option value="73">73</option>
	<option value="74">74</option>
	<option value="75">75</option>
	<option value="76">76</option>
	<option value="77">77</option>
	<option value="78">78</option>
	<option value="79">79</option>
	<option value="80">80</option>

</select>��</td>
                                                <td width="110" style="text-align: right">
                                                    ���Ҫ�� ��</td>
                                                <td width="197">
                                                    <select name="ddl_YqHeightL" id="ddl_YqHeightL">
	<option value="0">��ѡ��</option>
	<option value="145">145</option>
	<option value="146">146</option>
	<option value="147">147</option>
	<option value="148">148</option>
	<option value="149">149</option>
	<option value="150">150</option>
	<option value="151">151</option>
	<option value="152">152</option>
	<option value="153">153</option>
	<option value="154">154</option>
	<option value="155">155</option>
	<option value="156">156</option>
	<option value="157">157</option>
	<option value="158">158</option>
	<option value="159">159</option>
	<option value="160">160</option>
	<option value="161">161</option>
	<option value="162">162</option>
	<option value="163">163</option>
	<option value="164">164</option>
	<option value="165">165</option>
	<option value="166">166</option>
	<option value="167">167</option>
	<option value="168">168</option>
	<option value="169">169</option>
	<option value="170">170</option>
	<option value="171">171</option>
	<option value="172">172</option>
	<option value="173">173</option>
	<option value="174">174</option>
	<option value="175">175</option>
	<option value="176">176</option>
	<option value="177">177</option>
	<option value="178">178</option>
	<option value="179">179</option>
	<option value="180">180</option>
	<option value="181">181</option>
	<option value="182">182</option>
	<option value="183">183</option>
	<option value="184">184</option>
	<option value="185">185</option>
	<option value="186">186</option>
	<option value="187">187</option>
	<option value="188">188</option>
	<option value="189">189</option>
	<option value="190">190</option>
	<option value="191">191</option>
	<option value="192">192</option>
	<option value="193">193</option>
	<option value="194">194</option>
	<option value="195">195</option>
	<option value="196">196</option>
	<option value="197">197</option>
	<option value="198">198</option>
	<option value="199">199</option>
	<option value="200">200</option>
	<option value="201">201</option>
	<option value="202">202</option>
	<option value="203">203</option>
	<option value="204">204</option>
	<option value="205">205</option>
	<option value="206">206</option>
	<option value="207">207</option>
	<option value="208">208</option>
	<option value="209">209</option>
	<option value="210">210</option>
	<option value="211">211</option>
	<option value="212">212</option>
	<option value="213">213</option>
	<option value="214">214</option>
	<option value="215">215</option>
	<option value="216">216</option>
	<option value="217">217</option>
	<option value="218">218</option>

</select>
                                                    ��
                                                    <select name="ddl_YqHeightH" id="ddl_YqHeightH">
	<option value="300">��ѡ��</option>
	<option value="145">145</option>
	<option value="146">146</option>
	<option value="147">147</option>
	<option value="148">148</option>
	<option value="149">149</option>
	<option value="150">150</option>
	<option value="151">151</option>
	<option value="152">152</option>
	<option value="153">153</option>
	<option value="154">154</option>
	<option value="155">155</option>
	<option value="156">156</option>
	<option value="157">157</option>
	<option value="158">158</option>
	<option value="159">159</option>
	<option value="160">160</option>
	<option value="161">161</option>
	<option value="162">162</option>
	<option value="163">163</option>
	<option value="164">164</option>
	<option value="165">165</option>
	<option value="166">166</option>
	<option value="167">167</option>
	<option value="168">168</option>
	<option value="169">169</option>
	<option value="170">170</option>
	<option value="171">171</option>
	<option value="172">172</option>
	<option value="173">173</option>
	<option value="174">174</option>
	<option value="175">175</option>
	<option value="176">176</option>
	<option value="177">177</option>
	<option value="178">178</option>
	<option value="179">179</option>
	<option value="180">180</option>
	<option value="181">181</option>
	<option value="182">182</option>
	<option value="183">183</option>
	<option value="184">184</option>
	<option value="185">185</option>
	<option value="186">186</option>
	<option value="187">187</option>
	<option value="188">188</option>
	<option value="189">189</option>
	<option value="190">190</option>
	<option value="191">191</option>
	<option value="192">192</option>
	<option value="193">193</option>
	<option value="194">194</option>
	<option value="195">195</option>
	<option value="196">196</option>
	<option value="197">197</option>
	<option value="198">198</option>
	<option value="199">199</option>
	<option value="200">200</option>
	<option value="201">201</option>
	<option value="202">202</option>
	<option value="203">203</option>
	<option value="204">204</option>
	<option value="205">205</option>
	<option value="206">206</option>
	<option value="207">207</option>
	<option value="208">208</option>
	<option value="209">209</option>
	<option value="210">210</option>
	<option value="211">211</option>
	<option value="212">212</option>
	<option value="213">213</option>
	<option value="214">214</option>
	<option value="215">215</option>
	<option value="216">216</option>
	<option value="217">217</option>
	<option value="218">218</option>

</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ����Ҫ�� ��</td>
                                                <td>
                                                    <select name="ddl_YqHouseAndCar" id="ddl_YqHouseAndCar">
	<option selected="selected" value="">��ѡ��</option>
	<option value="00">��Ҫ��</option>
	<option value="01">Ҫ���з�</option>
	<option value="10">Ҫ���г�</option>
	<option value="11">�з��г�</option>

</select></td>
                                                <td style="text-align: right">
                                                    ���Ҫ�� ��</td>
                                                <td>
                                                    <select name="ddl_YqIsMarry" id="ddl_YqIsMarry">
	<option selected="selected" value="">��ѡ��</option>
	<option value="0">����Ҫ��</option>
	<option value="1">Ҫ��δ��</option>
	<option value="11">�ɴ�����</option>
	<option value="2">�ɴ�Ů��</option>
	<option value="12">�����к�</option>
	<option value="4">����δ��</option>

</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    �������� ��</td>
                                                <td>
                                                    <select name="ddl_YqMonthIn" id="ddl_YqMonthIn">
	<option selected="selected" value="">����</option>
	<option value="1">һ������</option>
	<option value="2">�е�����</option>
	<option value="3">�ߵ�����</option>

</select></td>
                                                <td style="text-align: right">
                                                    �������� ��</td>
                                                <td>
                                                    <select name="ddl_YqXingxiangsuzhi" id="ddl_YqXingxiangsuzhi" style="width:145px;">
	<option selected="selected" value="">��ѡ��</option>
	<option value="1">����ȥ����</option>
	<option value="2">��������Ϻ�</option>
	<option value="3">�������ʽϸ�</option>
	<option value="4">���������</option>
	<option value="5">�������ʸ�</option>
	<option value="6">�������ʽϸ�</option>
	<option value="7">�������ʸ�</option>

</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ѧ��Ҫ�� ��</td>
                                                <td>
                                                    <select name="ddl_YQEducation" id="ddl_YQEducation">
	<option selected="selected" value="0">����</option>
	<option value="1">û��ѧ��</option>
	<option value="2">Сѧ����</option>
	<option value="3">��������</option>
	<option value="4">��������</option>
	<option value="5">��ר����</option>
	<option value="6">��ר����</option>
	<option value="7">��������</option>
	<option value="8">�о�������</option>
	<option value="9">��ʿ</option>

</select></td>
                                                <td style="text-align: right">
                                                </td>
                                                <td>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ����Ҫ�� ��</td>
                                                <td colspan="3">
                                                    <select id="YQprovince" name="YQprovince" onchange="javascript:InitMenu2x(document.getElementById('YQprovince'),document.getElementById('YQcity'),document.getElementById('YQarea'), 0);SetArea('YQprovince','YQcity','YQarea','Y')"
                                                        style="width: 65px">
                                                    </select>
                                                    <select id="YQcity" name="YQcity" onchange="InitMenu3(document.getElementById('YQprovince'),document.getElementById('YQcity'), document.getElementById('YQarea'), 0);SetArea('YQprovince','YQcity','YQarea','Y')"
                                                        style="width: 80px">
                                                    </select>
                                                    <select id="YQarea" name="YQarea" onchange="SetArea('YQprovince','YQcity','YQarea','Y')">
                                                    </select>

                                                  <script language="javascript">
											        InitMenu1(document.getElementById("YQprovince"),document.getElementById("YQcity"),document.getElementById("YQarea"));
											        InitMenu2(document.getElementById("YQprovince"),document.getElementById("YQcity"),document.getElementById("YQarea"));
											        InitMenu3(document.getElementById("YQprovince"),document.getElementById("YQcity"),document.getElementById("YQarea"));
											        document.getElementById("YQprovince").options[0].text="����";
											        document.getElementById("YQcity").options[0].text="����";
											        document.getElementById("YQarea").options[0].text="����";
                                                    </script>

                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    ������żҪ��</td>
                                                <td colspan="3">
                                                    <textarea name="txb_OtherCase" rows="2" cols="20" id="txb_OtherCase" class="input2" onkeydown="limitChars(this.id, 50)" onchange="limitChars(this.id, 50)" onpropertychange="limitChars(this.id, 50)" style="height:50px;width:75%;"></textarea>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    �������� ��</td>
                                                <td colspan="3">
                                                    <table id="cbl_TSFW" border="0">
	<tr>
		<td><input name="cbl_TSFW$0" type="checkbox" id="cbl_TSFW_0"  value="1" />
		<label for="cbl_TSFW_0">����μ����׽��ѻ</label></td><td><input id="cbl_TSFW_1" type="checkbox" name="cbl_TSFW$1" value="1" /><label for="cbl_TSFW_1">�Ҷ������Ե����Ȥ</label></td><td><input id="cbl_TSFW_2" type="checkbox" name="cbl_TSFW$2" value="1" /><label for="cbl_TSFW_6">������������</label></td>
	</tr><tr>
		<td><input id="cbl_TSFW_3" type="checkbox" name="cbl_TSFW$3" value="1" /><label for="cbl_TSFW_3">������ȥ����</label></td><td><input id="cbl_TSFW_4" type="checkbox" name="cbl_TSFW$4"  value="1"/><label for="cbl_TSFW_4">������ѯ��������</label></td><td><input id="cbl_TSFW_5" type="checkbox" name="cbl_TSFW$5"  value="1"/>
		<label for="cbl_TSFW_5">�������һ����Ե</label></td>
	</tr><tr>
		<td><input id="cbl_TSFW_6" type="checkbox" name="cbl_TSFW$6" value="1" /><label for="cbl_TSFW_6">��������ͷ����</label></td><td></td><td></td>
	</tr>
</table></td>
                                            </tr>
                                        </table>
                </P>
                <p>&nbsp;</p>
              </div>
              <div 
class="member_register_1_title"><span>������(��������д�������˽����򵥻����߻����ɹ��ʣ�</span></div>
              <div class="member_register_1_content">
                <table cellspacing="5" width="630" border="0">
                  <tbody>
                    <tr>
                      <td valign="center" align="right" width="30" height="0">&nbsp;</td>
                      <td valign="center" colspan="2"><span style="COLOR: #ff0000">*</span><strong> ������:</strong>(200����˵˵���������硢������������������а���,�⽫�����������˽��㡢�����㣡) <br /></td>
                    </tr>
                    <tr>
                      <td align="right" height="0">&nbsp;</td>
                      <td colspan="2" height="400" rowspan="2"><label>&nbsp;
                            <textarea class="input2" onkeypress="return noNumbers(event)" onpaste="" id="txb_SigNature" onkeydown="limitChars(this.id, 200)" style="WIDTH: 85%; HEIGHT: 120px" onpropertychange="limitChars(this.id, 200)" name="txb_SigNature" onchange="limitChars(this.id, 200)"></textarea>
                        </label>
                          <br />
                          <span id="RequiredFieldValidator6" 
      style="VISIBILITY: hidden; COLOR: red">������һ�����İ�����</span></td>
                    </tr>
                    <tr>
                      <td align="right" height="0">&nbsp;</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div 
class="member_register_1_title"><span>�ϴ�ͷ��(����Ƭ�Ļ�Ա�ܵ��Ĺ�ע�ʿ����30��Ŷ����߻����ɹ��ʣ�</span></div>
              <div class="member_register_1_content">
                <table cellspacing="5" width="623" border="0">
                  <tbody>
                    <tr>
                      <td width="93" align="right">��ܰ��ʾ��</td>
                      <td width="511" align="left">Ϊ��ʹ��ͷ��ﵽ�����ʾЧ�����ñ�������Ŀ����㡣<br />
                        ���ϴ��ߴ�Ϊ��120���� * 
                        ��145���ص���Ƭ����ͼƬ����ֱ�������</td>
                    </tr>
                    <tr>
                      <td align="right" >���ϴ�ͷ��</td>
                      <td><input type="text" id="url1" value="" name="pic" /> <input type="button" id="image1" value="ѡ��ͼƬ" />������ͼƬ + �����ϴ���</td>
                    </tr>
                     <tr>
                      <td align="right" >ͷ�����ã�</td>
                      <td><input type="radio" name="quanxian"  value="1" checked="checked"/>�������˿ɼ�<input type="radio" name="quanxian"  value="1"/>���Ի�Ա�ɼ�</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div class="member_register_1_title"><span>��֤��Ϣ�����������Ķ������������˽���ߣ�</span></div>
              <div class="member_register_button">
                <table cellspacing="5" width="640" border="0">
                  <tbody>
                    <tr>
                      <td width="628" align="middle">��֤�룺
                       <input name="txtCode" type="text" id="txtCode" size="20" maxlength="20" />
                        <a href="#" onclick="refreshimg();"><img src="inc/Code.asp" id="getcode" alt="����������ٻ�һ��" border="0"/></a><span id="lbError" 
      style="COLOR: red"></span><span id="RequiredFieldValidator4" 
      style="VISIBILITY: hidden; COLOR: red">����д��֤��</span></td>
                    </tr>
                    <tr>
                      <td align="middle"><textarea class="input2" id="TextBox1" style="WIDTH: 85%; HEIGHT: 150px" name="TextBox1">1�������������л����񹲺͹��ĸ���ɡ����棻�����л����񹲺͹��йػ������ĸ���桢����������һ����Ե����������վ�³̡�
2�����������ػ��������£��������أ��԰��ĶԴ����ˣ��������������й����Ρ�����ɫ��⽨���š���������Υ����Ϣ������ɧ�Ż���ƭ������Ա����������վ�ڹ��������κη̰������衢���ŵ����ݣ������໰��ɫ�����֣�������Ƭ������ʹ��һ����Ե������ҵ����Ƿ���Ϊ������һ����Ե��Ȩɾ�����������ϣ�������ͬʱ�е��ɴ�ֱ�ӻ�����������»����·������Ρ�
3�����Ļ�Ա������سƲ����Ժ������ĵ绰���롢�����ʼ���QQ���롢��ַ�����������絽�������ϡ�
4����������ʵ��д�������Ϻ��ϴ����˵���ʵ���գ���������������Ϊ����վ��֤ʵ����Ȩȡ�����Ա�ʸ�ˡ���˷ѣ���׷���䷨�����Ρ�
5������ע��ɹ���������һ��֮��Я����ʵ��Ч�����֤������֤����ѧ��֤������֤��ԭ���͸�ӡ����һ�ݣ�1��5�Ÿ�����ʵ���յ�һ����Ե�ķ�վ�ͻ��������˴�����������ܲ���ʽ���������

��������
1��һ����Ե�����κΰ����ڡ����ɡ������ᡢ���ػ���κ��뱾��վ�йط���(���¼�Ʒ���)�����֮��Ѷ�����ݻ���(���¼������)����������֤������֮��ȷ�Ի�ɿ��ԣ����ң�������͸��������֮��桢��Ѷ��ҪԼ��չʾ�������ȡ��֮�κβ�Ʒ����Ѷ������(���¼�Ʋ�Ʒ)������վ�಻���κ����Ρ�
2��һ����Ե֮�����������ǻ�����״�ṩ������һ����Ե��ȷ�ر�ʾ�ܾ����ڷ������ϻ��Ʒ�����κ���ʾ��ʾ֮��֤�������������ڣ���Ϊ��ҵʹ�û��ʺ����ض�Ŀ��֮��֤��
3��һ����Ե����������һ���������û���Ҫ��Ҳ���������񲻻����жϣ��Է���ļ�ʱ�ԡ���ȫ�ԡ�������������������
4���û����Լ��Ķ����жϴ����뽻����ص���Ϊ���������е����ܲ��������Σ�һ����Ե���е��κη������Ρ�
5��һ����Ե��������������
1) ʹ����վ����ȡ��֮�����ȷ�ɿ���
2����վ��������Ƭ����ʵ�ԣ�
2) ��������վ�������ȡ��֮�κβ�Ʒ��������Ѷ��������Ϣ����������������
6��һ����Ե���ڱ���վ�߻���������֯���ǳа���κλ�Ա�(��������������ȡ�����Լ���ȫ����Ļ)������һ���������û�����Ҫ��Ҳ��������Ա�μӻ�Ľ���������κ�����ж�����ȷ�ԡ��Ϸ��Ϲ��ԡ��ɴ˲������κζ��ڻ�Ա���˻������˵���������������Լ������𺦣�����վ���е��κ����Ρ�
7�������û���һ����Եһ��ͬ���йط��������Լ�ʹ��һ����Ե�ķ�����������齻���������ٲ�ίԱ�����������κη��������뷨����ִ�������Щ����������ܽӽ��ķ������½����������������򱣳ֶ��û���������Ч����Ӱ�졣
8�����������������ս���Ȩ���޸�Ȩ��һ����Ե���С� </textarea></td>
                    </tr>
                    <tr>
                      <td align="middle"><input id="cb_Agree" onclick="SetSignbtnState();" 
      type="checkbox" checked="checked" name="cb_Agree" />
                          <label for="cb_Agree">���Ķ���ͬ������������˽���� </label></td>
                    </tr>
                    <tr>
                      <td>                          <span id="lb_ShowError" 
      style="FONT-WEIGHT: bold; COLOR: red"></span>&nbsp;&nbsp;&nbsp;
                          <input 
      id="hidden_Province" style="WIDTH: 1px" type="hidden" name="hidden_Province" />
                          <input id="hidden_City" style="WIDTH: 1px" type="hidden" name="hidden_City" />
                          <input id="hidden_Area" style="WIDTH: 1px" type="hidden" name="hidden_Area" />
                          <input id="hidden_YqProvince" style="WIDTH: 1px" type="hidden" 
      name="hidden_YqProvince" />
                          <input id="hidden_YqCity" style="WIDTH: 1px" 
      type="hidden" name="hidden_YqCity" />
                          <input id="hidden_YqArea" style="WIDTH: 1px" 
      type="hidden" name="hidden_YqArea" />
      <input id="flag" style="WIDTH: 1px" type="hidden" name="flag" value="1" />

                        &nbsp;
                        <label>
                          <input type="submit" name="button" id="button" value="�ύ" />
                        </label>                        <br /></td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div></td>
          </tr>
          <tr>
            <td colspan="4" align="left"><img src="images/z_08.png" width="675" height="9" alt="" /></td>
          </tr>
        </table></form></td>
        <td width="319" align="center" valign="top"><table id="__01" width="309" height="202" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="26" valign="top" background="images/j_01.png">&nbsp;</td>
          </tr>
          <tr>
            <td background="images/j_02.png"><table width="288" height="103" align="center" cellspacing="0">
                <tr>
                  <td height="89" align="left" class="css">һ����Ե����һ����˳����ʾ��������������ͨ������Ĳ���ʹ�Լ������������������ǰ�棺<br />
                    1��������д�������û�������ڱȽϿ�ǰ��λ��<br />
                    2������Ƭ���û���������ʾ����<br />
                    3����󣬼ǵþ�������䣬�ȽϸߵĻ�Ծ�����ø����˿����㣬˵�������TAҲ�����С�<br />
                    &gt; �Ҳ���?�ð�����ʰ����� &lt;</td>
                  </tr>
              </table>
                </td>
          </tr>
          <tr>
            <td height="8"><img src="images/j_04.png" width="309" height="8" alt="" /></td>
          </tr>
        </table>
          <table width="307" height="101" border="0" align="center" cellpadding="0" cellspacing="0" id="__01">
            <tr>
              <td height="9" colspan="3"><img src="images/c_01.png" width="307" height="9" alt="" /></td>
            </tr>
            <tr>
              <td height="28" colspan="3" background="images/c_02.png"><table width="296" align="left" cellspacing="0">
                <tr>
                  <td width="68" align="right" class="hui12">��Ե�</td>
                  <td width="231" align="right" class="hui12"><a href="qqhd.asp" class="bai12a">&gt;&gt;����</a></td>
                </tr>
              </table></td>
            </tr>
            <tr>
              <td colspan="3" valign="top" background="images/c_03.png"><table width="292" align="center" cellspacing="0">

                  <tr>
                    <td align="left">  <%				SQLSelect="select  top 5 * from [Product]  where ClassID=20  Order By [id] desc "
Rs.Open SQLSelect,Conn,1,1
do while not Rs.eof %>
					<table width="287" cellspacing="0">
                        <tr>
                          <td height="20" class="xuxian" style="padding-left:5px"><a href="qqhd_view.asp?id=<%=Rs("id")%>" class="hui1"><%=leftstr(trim(Rs("title")),30)%></a></td>
                          </tr>
                      </table><%rs.movenext
					  loop
					  rs.close%>
                 <!--       <table width="287" cellspacing="0">
                          <tr>
                            <td height="21" align="left" class="xuxian STYLE1">[����] 12��16��  ��ʯ��Ů ����ɶ� </td>
                          </tr>
                      </table>
                        <table width="287" cellspacing="0">
                          <tr>
                            <td height="21" align="left" class="xuxian STYLE1">[����] 12��16��  ��ʯ��Ů ����ɶ� </td>
                          </tr>
                        </table>
                        <table width="287" cellspacing="0">
                          <tr>
                            <td height="21" align="left" class="xuxian STYLE1">[����] 12��16��  ��ʯ��Ů ����ɶ� </td>
                          </tr>
                        </table>
                        <table width="287" cellspacing="0">
                          <tr>
                            <td height="21" align="left" class="xuxian STYLE1">[����] 12��16��  ��ʯ��Ů ����ɶ� </td>
                          </tr>
                        </table>--></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td height="12" colspan="3"><img src="images/c_09.png" width="307" height="12" alt="" /></td>
            </tr>
          </table>
            <table width="307" height="101" border="0" align="center" cellpadding="0" cellspacing="0" id="__01">
          <tr>
            <td height="9" colspan="3"><img src="images/c_01.png" width="307" height="9" alt="" /></td>
          </tr>
          <tr>
            <td height="28" colspan="3" background="images/c_02.png"><table width="295" align="left" cellspacing="0">
              <tr>
                <td width="68" align="right" class="hui12">�ɹ�����</td>
                <td width="221" align="right" class="hui12"><a href="SuccessStory.asp" class="bai12a">&gt;&gt;����</a><a href="news.asp" class="bai12"></a></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td colspan="3" background="images/c_03.png"><table width="292" align="center" cellspacing="0">
              <tr>
                <td align="center"><a href="SuccessStory.asp"><img src="images/c01.jpg" width="134" height="113" border="0" /></a></td>
                <td align="center"><a href="SuccessStory.asp"><img src="images/c02.jpg" width="140" height="110" border="0" /></a></td>
              </tr>
              <tr>
                <td align="center" class="css">�Ҹ��Ŀ�ʼ</td>
                <td align="center"><span class="css">���۵���֪</span></td>
              </tr>
              <tr>
                <td colspan="2" align="left"><table width="287" cellspacing="0">
                  <tr>
                    <td width="19" height="20" align="center"><img src="images/c_06.png" width="12" height="10" /></td>
                    <td class="xuxian"><strong>������</strong> �� <strong>��С��</strong><a href="SuccessStory_con.asp?id=5" class="hui1">�鿴���ǵ� �ɹ�����&gt;&gt;</a> </td>
                  </tr>
                </table>
                  <table width="287" cellspacing="0">
                    <tr>
                      <td width="19" height="21" align="center"><img src="images/c_06.png" width="12" height="10" /></td>
                      <td class="xuxian"><strong>������</strong> �� <strong>��С��</strong><a href="SuccessStory_con.asp?id=4" class="hui1">�鿴���ǵ� �ɹ�����&gt;&gt;</a> </td>
                    </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="12" colspan="3"><img src="images/c_09.png" width="307" height="12" alt="" /></td>
          </tr>
        </table>
          <table id="__01" width="309" height="202" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="26" valign="top" background="images/j_01.png"><table width="290" align="center" cellspacing="0">
                  <tr>
                    <td width="72" height="40" align="center" class="hui14xx">��������</td>
                    <td width="212" align="right" ><a href="hqly.asp" class="bai12a" >&gt;&gt;����</a></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td background="images/j_02.png"> <%SQLSelect="select  top 2 * from [about]  Order By [id] desc "
Rs.Open SQLSelect,Conn,1,1
do while not Rs.eof %>
              <table width="299" height="103" align="center" cellspacing="0">
                <tr>
                  <td width="153" height="89" class="css"><a href="hqly.asp?id=<%=Rs("id")%>"><img src="images/hq.jpg" width="153" height="81" border="0" /></a></td>
                  <td width="140" class="css"><a href="hqly.asp?id=<%=Rs("id")%>" class="hui1"><%=Rs("title")%></a></td>
                </tr>
              </table><%rs.movenext
			  loop
			  rs.close%>
             <!--   <table width="299" height="103" align="center" cellspacing="0">
                  <tr>
                    <td width="153" height="89" class="css"><img src="images/hq.jpg" width="153" height="81" /></td>
                    <td width="140" class="css">����Ļ����ǴӸ�ˮ׼���⣬����߻���ʼ��..</td>
                  </tr>
                </table>--></td>
            </tr>
            <tr>
              <td height="8"><img src="images/j_04.png" width="309" height="8" alt="" /></td>
            </tr>
          </table>
          <table id="__01" width="309" height="265" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="158" valign="top"><img src="images/ad_hlgw.jpg" width="307" height="151" /></td>
              </tr>
              <tr>
                <td height="78"><img src="images/ad_lkss.jpg" width="307" height="107" /></td>
              </tr>
          </table></td>
      </tr>
    </table></td>
    <td width="5">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table id="__01" width="972" height="99" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <!--#include file="foot.asp" -->
      </tr>
    </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
<%	if request("flag")<>"" then
flag=1
	email=request("txb_email")
	if email="" then
flag=0%>
<script language="vbscript">
msgbox"��½���䲻��Ϊ�գ�����������!"
</script>
<%
	
		end if 
		for i=1 to len(trim(email))
if mid(trim(email),i,1)="@" then
flag1=1
end if
next
if flag1<>1 then
flag=0%>
<script language="vbscript">
msgbox"��������ȷ������!"
</script>
<%

end if 

sql="select * from [Users] where email='"&trim(request("txb_email"))&"'"
rs.open sql,conn,3,2
if  not rs.eof then
flag=0%>
<script language="vbscript">
msgbox"�������ѱ���ע��!"
</script><%
end if

rs.close

Password=request("txb_PassWord")
	rePassword=request("txb_rePassword")
	if Password="" then
flag=0%>
<script language="vbscript">
msgbox"���벻��Ϊ��!"
</script>
<%end if

	if Password<>rePassword then
flag=0%>
<script language="vbscript">
msgbox"�������벻һ�£�����������!"
</script>
<%end if 
response.Write("1")
	
	NickName=request("txb_NickName")
	
if NickName="" then
flag=0%>
<script language="vbscript">
msgbox"�ǳƲ���Ϊ��!"
</script>
<%
response.Write("1")

else
	sql="select * from [Users] where NickName='"&trim(request("txb_NickName"))&"'"
rs.open sql,conn,3,2
if  not rs.eof then
flag=0%>
<script language="vbscript">
msgbox"�ǳ��ѱ���ʹ��!"
</script>
<%end if 
end if
response.Write("1")
rs.close
TelePhone=request("txb_TelePhone")
if TelePhone="" then
flag=0%>
<script language="vbscript">
msgbox"����д��ϵ�绰!"
</script>
<%
elseif len(trim(request("txb_TelePhone")))<>11then
flag=0%>
<script language="vbscript">
msgbox"��������ȷ�ĵ绰!"
</script><%

end if
response.Write("1")
Marry=request("ddl_Marry")
if  Marry="" then
flag=0%>
<script language="vbscript">
msgbox"����д����״��!"
</script><%

end if
response.Write("1")
BirthDay=request("txb_BirthDay")
if  BirthDay="" then
flag=0%>
<script language="vbscript">
msgbox"����д��������!"
</script><%

end if
response.Write("1")
Education=request("ddl_Education")
if  Education="" then
flag=0%>
<script language="vbscript">
msgbox"����дѧ��!"
</script><%
end if
response.Write("1")
province=request("mprovince")
city=request("mcity")
country=request("mcountry")
if  province="" or  city=""  or  country=""  then
flag=0%>
<script language="vbscript">
msgbox"����д����!"
</script><%
end if
response.Write("1")
SigNature=request("txb_SigNature")
if  SigNature="" then
flag=0
Response.Write("<script>alert(""����д������"");history.back()</script>")
end if
response.Write("1")
txtCode=Trim(Request.Form("txtCode"))
If IsNumeric(txtCode)=False Or txtCode="" Or txtCode <> Session("GetCode") Then 
flag=0
Response.Write("<script>alert(""��֤���������"");history.back()</script>")
end if
Pic=Request.Form("Pic")
if flag=1 then
	Rs.Open "Select * From [Users]",conn,1,2
		Rs.Addnew
		
		ID=Rs("ID")
		Rs("ClassID")="4"
		Rs("Minzu")=request("ddl_Minzu")
		Rs("Height")=request("ddl_Height")
		Rs("weight")=request("ddl_weight")
		Rs("RealName")=request("txb_RealName")
		Rs("email")=email
		Rs("BloodType")=request("ddl_BloodType")
		Rs("Password")=MD5(Password)
		Rs("NickName")=NickName
		Rs("BirthDay")=BirthDay
		Rs("csn")=year(BirthDay)
	
		Rs("Pic")=Pic
	    Rs("Danweixingzhi")=request("ddl_Danweixingzhi")
        Rs("job")=request("ddl_job")
        Rs("province")=province
		Rs("city")=city
		Rs("country")=country
		Rs("MonthIn")=request("ddl_MonthIn")
        Rs("House")=request("ddl_House")
		Rs("Car")=request("ddl_Car")
		Rs("AiHao")=request("ddl_AiHao")
		Rs("Xingge")=request("ddl_Xingge")
		Rs("Wenhua")=request("ddl_Wenhua")
		Rs("Zonghetezhi")=request("ddl_Zonghetezhi")
		Rs("TeBieshuoming")=request("ddl_TeBieshuoming")
		Rs("Xingxiangsuzhi")=request("ddl_Xingxiangsuzhi")
		Rs("Gerenzongshu")=request("ddl_Gerenzongshu")
		Rs("YqBeginAge")=request("ddl_YqBeginAge")
		Rs("YqEndAge")=request("ddl_YqEndAge")
		Rs("YqHeightL")=request("ddl_YqHeightL")
		Rs("YqHeightH")=request("ddl_YqHeightH")
		Rs("YqHouseAndCar")=request("ddl_YqHouseAndCar")
		Rs("YqIsMarry")=request("ddl_YqIsMarry")
		Rs("YqXingxiangsuzhi")=request("ddl_YqXingxiangsuzhi")
		Rs("YqMonthIn")=request("ddl_YqMonthIn")
		Rs("YQprovince")=request("YQprovince")
		Rs("YQcity")=request("YQcity")
		Rs("YQarea")=request("YQarea")
		Rs("OtherCase")=request("txb_OtherCase")
		Rs("YQEducation")=request("ddl_YQEducation")
		Rs("LastLogin")=now()
		Rs("Hits")=0
	 
		Rs("QQ")=request("txb_QQ")
		Rs("Family")=request("ddl_Family") 
		Rs("LiYiNum")=request("txb_LiYiNum") 
		Rs("Marry")=Marry
		Rs("Specialty")=request("ddl_Specialty")
		Rs("Education")=Education
	    Rs("star")=request("txb_star")
		Rs("TelePhone")=TelePhone
	    Rs("Gender")=request("rbtn_Gender")
		Rs("Identity")=request("txb_Identity") 
        Rs("Shuxiang")=request("txb_Shuxiang")
		Rs("FinishSchool")=request("txb_FinishSchool") 
	    Rs("WorkCompany")=request("txb_WorkCompany") 
	    Rs("DateAndTime")=Now()
		Rs("SigNature")=SigNature
		Rs("TSFW$3")=request("cbl_TSFW$3")
		Rs("TSFW$1")=request("cbl_TSFW$1")
		Rs("TSFW$0")=request("cbl_TSFW$0")
		Rs("TSFW$2")=request("cbl_TSFW$2")
		Rs("TSFW$4")=request("cbl_TSFW$4")
		Rs("TSFW$5")=request("cbl_TSFW$5")
		Rs("TSFW$6")=request("cbl_TSFW$6")
		Rs.Update
		Rs.Close

	
	%>
		<script language="vbscript">
msgbox"��Ϣ��ӳɹ�!"
</script><%	
response.Redirect "login.asp"
end if


end if%>
</body>
</html>
