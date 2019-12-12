<!--#include file="System.asp"-->
<!--#include file="inc/Md5.asp"-->
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
            if(ProvinceText != "省" && ProvinceText != "不限")
            {
                document.getElementById("hidden_YqProvince").value = ProvinceText;
                if(CityText != "==所有==" && CityText != "不限")
                {
                    document.getElementById("hidden_YqCity").value = CityText;
                    if (AreaText != "==所有==" && AreaText != "==请选择==" && AreaText != "县" && AreaText != "不限")
                        document.getElementById("hidden_YqArea").value = AreaText;
                    if (AreaText == "==所有==" || AreaText == "==请选择==" || AreaText == "县" && AreaText != "不限")
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
            if(ProvinceText != "省" && ProvinceText != "不限")
            {
                
                document.getElementById("hidden_Province").value = ProvinceText;
                if(CityText != "==所有==" && CityText != "不限")
                {
                    document.getElementById("hidden_City").value = CityText;
                    if (AreaText != "==所有==" && AreaText != "==请选择==" && AreaText != "县" && AreaText != "不限")
                        document.getElementById("hidden_Area").value = AreaText;
                    if (AreaText == "==所有==" || AreaText == "==请选择==" || AreaText == "县" && AreaText != "不限")
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
                
                if (month == 1 && date >=20 || month == 2 && date <=18) {star = "水瓶座";}
                if (month == 1 && date > 31) {value = "Huh?";}
                if (month == 2 && date >=19 || month == 3 && date <=20) {star = "双鱼座";}
                if (month == 2 && date > 29) {value = "Say what?";}
                if (month == 3 && date >=21 || month == 4 && date <=19) {star = "白羊座";}
                if (month == 3 && date > 31) {value = "OK.  Whatever.";}
                if (month == 4 && date >=20 || month == 5 && date <=20) {star = "金牛座";}
                if (month == 4 && date > 30) {value = "I'm soooo sorry!";}
                if (month == 5 && date >=21 || month == 6 && date <=21) {star = "双子座";}
                if (month == 5 && date > 31) {value = "Umm ... no.";}
                if (month == 6 && date >=22 || month == 7 && date <=22) {star = "巨蟹座";}
                if (month == 6 && date > 30) {value = "Sorry.";}
                if (month == 7 && date >=23 || month == 8 && date <=22) {star = "狮子座";}
                if (month == 7 && date > 31) {value = "Excuse me?";}
                if (month == 8 && date >=23 || month == 9 && date <=22) {star = "处女座";}
                if (month == 8 && date > 31) {value = "Yeah. Right.";}
                if (month == 9 && date >=23 || month == 10 && date <=22) {star = "天秤座";}
                if (month == 9 && date > 30) {value = "Try Again.";}
                if (month == 10 && date >=23 || month == 11 && date <=21) {star = "天蝎座";}
                if (month == 10 && date > 31) {value = "Forget it!";}
                if (month == 11 && date >=22 || month == 12 && date <=21) {star = "射手座";}
                if (month == 11 && date > 30) {value = "Invalid Date";}
                if (month == 12 && date >=22 || month == 1 && date <=19) {star = "摩羯座";}
                if (month == 12 && date > 31) {value = "No way!";}
                
                x = (start - birthyear) % 12
                
                if (x == 1 || x == -11) {shuxiang = "鼠";}
                if (x == 0) {shuxiang = "牛";}
                if (x == 11 || x == -1) {shuxiang = "虎";}
                if (x == 10 || x == -2) {shuxiang = "兔";}
                if (x == 9 || x == -3)  {shuxiang = "龙";}
                if (x == 8 || x == -4)  {shuxiang ="蛇";}
                if (x == 7 || x == -5)  {shuxiang = "马";}
                if (x == 6 || x == -6)  {shuxiang = "羊";}
                if (x == 5 || x == -7)  {shuxiang = "猴";}
                if (x == 4 || x == -8)  {shuxiang = "鸡";}
                if (x == 3 || x == -9)  {shuxiang = "狗";}
                if (x == 2 || x == -10)  {shuxiang = "猪";}
                
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
        // 用正则表达式将前后空格  
        // 用空字符串替代。  
        return this.replace(/(^\s*)|(\s*$)/g, "");  
    } 
	
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
    <td width="5">&nbsp;</td>
    <td width="986"><table width="983" cellspacing="0">
      <tr>
        <td width="640"><form name="myform" method="post" action="register.asp" >
        <table width="674" border="0" align="left" cellpadding="0" cellspacing="0" id="__01">
          <tr>
            <td width="674" height="34" colspan="4" background="images/g_011.png"><table width="653" height="19" cellspacing="0">
                <tr>
                  <td width="34">&nbsp;</td>
                  <td width="476" class="mei12">用户登录</td>
                  <td width="135" class="bai12">&nbsp;</td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td colspan="4" valign="top" background="images/g_031.png"><div>
              <div class="member_register_1_title"><span>用户登录信息（您在登录网站时必须使用）</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table cellspacing="5" width="600" border="0">
                                            <tr style="height: 65px">
                                                <td width="150" height="0" align="right">
                                                    用来登录的邮箱 ：</td>
                                                <td width="180">
                                                    <input name="txb_email" type="text" id="txb_email" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" /><br />
                                                    <br />
                                                  
                                                </td>
                                                <td>
                                                    <span id="showEmailInfo"><span style="color: #ff0000">*</span> 非真实邮件将收不到对方给您发的邮件</span></td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    设置用户名密码 ：</td>
                                              <td>
                                                <input name="txb_PassWord" type="password" maxlength="20" id="txb_PassWord" class="input2" style="height: 17px;
                                                        background-color: #FFF; border: 1px solid #ffcdd2;" /><br /></td>
                                                <td>
                                                    <span style="color: #ff0000">*</span> 请输入用户名密码</td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    确认密码 ：</td>
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
class="member_register_1_title"><span>对外保密的基本信息资料（请认真填写，这些资料只有您和管理员才能看到）</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table width="650" border="0" cellspacing="5">
                                            <tr style="height: 75px">
                                                <td width="88" height="0" align="right">
                                                    真实姓名 ：</td>
                                                <td width="211">
                                                    <input name="txb_RealName" type="text" id="txb_RealName" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                              </td>
                                                <td width="101" align="right">
                                                    我的昵称 ：</td>
                                                <td width="217">
                                                    <input name="txb_NickName" type="text" id="txb_NickName" class="input2" onblur="checkNickName(this.value)" style="height: 17px; width: 150px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                    <span style="color: #ff0000">*</span><br />
                                                    <span id="nickname"></span>
                                              </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    联系电话 ：</td>
                                                <td><input name="txb_TelePhone" type="text" id="txb_TelePhone" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />                                                  <span style="color: #ff0000">*</span><br />
                                                </td>
                                                <td align="right">
                                                    Q Q号码 ：
                                                </td>
                                                <td>
                                                    <input name="txb_QQ" type="text" id="txb_QQ" class="input2" style="height: 17px; width: 150px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    身份证号 ：</td>
                                                <td>
                                                    <input name="txb_Identity" type="text" id="txb_Identity" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" /></td>
                                                <td align="right">
                                                    毕业院校 ：</td>
                                                <td>
                                                    <input name="txb_FinishSchool" type="text" id="txb_FinishSchool" class="input2" style="height: 17px; width: 150px;
                                                        background-color: #FFF; border: 1px solid #ffcdd2;" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    工作单位 ：</td>
                                                <td>
                                                    <input name="txb_WorkCompany" type="text" id="txb_WorkCompany" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                                <td align="right">
                                                    家庭背景 ：</td>
                                                <td>
                                                    <select name="ddl_Family" id="ddl_Family" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2; color: #e65479; height: 22px;">
	<option selected="selected" value="">请选择</option>
	<option value="1">普通家庭 独生子女</option>
	<option value="2">普通家庭 非独生子女</option>
	<option value="3">中产家庭 独生子女</option>
	<option value="4">中产家庭 非独生子女</option>
	<option value="5">高等家庭 独生子女</option>
	<option value="6">高等家庭 非独生子女</option>

</select>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="0" align="right">
                                                    离异证号 ：</td>
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
class="member_register_1_title"><span>对外公开的基本信息资料（请认真填写，这将有助于您婚恋成功）</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table width="600" border="0" cellspacing="5" onload="AddList('age1', 'age2'); new PCAS('province','city');" onclick="signs(document.getElementById('txb_Birthday').value)" >
                                            <tr>
                                                <td width="82" style="text-align: right">
                                                    婚姻状况 ：</td>
                                                <td width="208">
                                                    <select name="ddl_Marry" id="ddl_Marry">
	<option value="">请选择</option>
	<option value="1">未婚</option>
	<option value="2">离异带男孩</option>
	<option value="3">离异带女孩</option>
	<option value="4">离异未育</option>
	<option value="5">离异男孩归对方</option>
	<option value="12">离异女孩归对方</option>
	<option value="6">丧偶带男孩</option>
	<option value="7">丧偶带女孩</option>
	<option value="8">丧偶未育</option>
	<option value="9">短婚未育</option>
	<option value="10">实事婚姻带男孩</option>
	<option value="11">实事婚姻带女孩</option>
</select>
                                                    <span style="color: #ff0000">*</span><br />                                              </td>
                                                <td width="104" style="text-align: right">
                                                    性&nbsp; &nbsp;别 ：</td>
                                                <td width="173">
                                                    <select name="rbtn_Gender" id="rbtn_Gender">
	<option selected="selected" value="1">男</option>
	<option value="0">女</option>
</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    出生年月 ：</td>
                                                <td>
                                                    <input name="txb_BirthDay" type="text" id="txb_BirthDay" class="input2" onblur="signs(this.value)" onchange="signs(this.value)" onfocus="calendar()" style="height: 17px; background-color: #FFF; border: 1px solid #ffcdd2;" />
                                                    <span style="color: #ff0000">*</span><br />
                                                </td>
                                                <td style="text-align: right">
                                                    属&nbsp; &nbsp;相 ：</td>
                                                <td>
                                                    <input name="txb_Shuxiang" type="text" id="txb_Shuxiang" class="input2" style="height: 17px; background-color: #FFF;
                                                        border: 1px solid #ffcdd2;" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    星&nbsp; &nbsp;座 ：</td>
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
                                                    学&nbsp; &nbsp;历 ：</td>
                                                <td><select name="ddl_Education" id="ddl_Education">
                                                  <option selected="selected" value="">请选择</option>
                                                  <option value="1">没有学历</option>
                                                  <option value="2">小学</option>
                                                  <option value="3">初中</option>
                                                  <option value="4">高中</option>
                                                  <option value="5">中专</option>
                                                  <option value="6">大专</option>
                                                  <option value="7">本科</option>
                                                  <option value="8">研究生</option>
                                                  <option value="9">博士</option>
                                                </select>
                                                <span style="color: #ff0000">*</span>                                                </td>
                                                <td style="text-align: right">
                                                    专&nbsp; &nbsp;业 ：</td>
                                                <td><select name="ddl_Specialty" id="ddl_Specialty">
                                                  <option selected="selected" value="">请选择</option>
                                                  <option value="1">理工类</option>
                                                  <option value="2">文科类</option>
                                                  <option value="3">医学类</option>
                                                  <option value="4">农学类</option>
                                                  <option value="5">军事类</option>
                                                  <option value="6">体育类</option>
                                                  <option value="7">天文类</option>
                                                  <option value="8">地理类</option>
                                                  <option value="9">没有专业</option>
                                                </select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    血&nbsp;&nbsp;型 ：</td>
                                                <td>
                                                    <select name="ddl_BloodType" id="ddl_BloodType">
	<option selected="selected" value="">请选择</option>
	<option value="1">A</option>
	<option value="2">B</option>
	<option value="3">AB</option>
	<option value="4">O</option>
	<option value="5">其他</option>
	<option value="6">不知道</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    民&nbsp;&nbsp;族 ：</td>
                                                <td>
                                                    <select name="ddl_Minzu" id="ddl_Minzu">
	<option value="">请选择</option>
	<option value="1">汉族</option>
	<option value="2">壮族</option>
	<option value="3">回族</option>
	<option value="4">满族</option>
	<option value="5">蒙古族</option>
	<option value="6">藏族</option>
	<option value="7">朝鲜族</option>
	<option value="8">维吾尔族</option>
	<option value="9">彝族</option>
	<option value="10">苗族</option>
	<option value="11">其它民族</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    身高（cm）：</td>
                                                <td>
                                                    <select name="ddl_Height" id="ddl_Height">
	<option value="0">请选择</option>
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
                                                    体重（kg）：</td>
                                                <td>
                                                    <select name="ddl_weight" id="ddl_weight">
	<option value="0">请选择</option>
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
                                                    地&nbsp; &nbsp;区 ：</td>
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
											        document.getElementById("mprovince").options[0].text="省";
											        document.getElementById("mcity").options[0].text="市";
											        document.getElementById("mcountry").options[0].text="县";
                                                    </script>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    单位性质 ：</td>
                                                <td>
                                                    <select name="ddl_Danweixingzhi" id="ddl_Danweixingzhi">
	<option selected="selected" value="">请选择</option>
	<option value="1">国家机构</option>
	<option value="2">教育机构</option>
	<option value="3">医疗机构</option>
	<option value="4">金融业</option>
	<option value="5">保险业</option>
	<option value="6">公司</option>
	<option value="7">工厂</option>
	<option value="8">个体经营</option>
	<option value="9">部队</option>
	<option value="10">国营大企</option>
	<option value="11">外资企业</option>
	<option value="12">合资企业</option>
	<option value="13">没有单位</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    职业职务 ：</td>
                                                <td>
                                                    <select name="ddl_job" id="ddl_job">
	<option value="">请选择</option>
	<option value="1">公务员</option>
	<option value="2">国家干部</option>
	<option value="3">教师</option>
	<option value="4">大学生</option>
	<option value="5">运动员</option>
	<option value="6">工程师</option>
	<option value="7">经理</option>
	<option value="8">总经理</option>
	<option value="9">IT行业</option>
	<option value="10">个体业主</option>
	<option value="11">职员</option>
	<option value="12">主管</option>
	<option value="13">经纪人</option>
	<option value="14">警官</option>
	<option value="15">科员</option>
	<option value="16">教练</option>
	<option value="17">医生</option>
	<option value="18">护士</option>
	<option value="19">演艺</option>
	<option value="20">士兵</option>
	<option value="21">律师</option>
	<option value="22">自由职业</option>
	<option value="23">其它</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    月&nbsp;收&nbsp;入：</td>
                                                <td>
                                                    <select name="ddl_MonthIn" id="ddl_MonthIn">
	<option selected="selected" value="">请选择</option>
	<option value="1">500元</option>
	<option value="2">1000元左右</option>
	<option value="3">1000-2000元</option>
	<option value="4">2000-3000元</option>
	<option value="5">3000-4000元</option>
	<option value="6">4000-5000元</option>
	<option value="7">5000-6000元</option>
	<option value="8">6000-7000元</option>
	<option value="9">7000-8000元</option>
	<option value="10">8000-9000元</option>
	<option value="11">10000元左右</option>
	<option value="12">20000元左右</option>
	<option value="13">30000元左右</option>
	<option value="14">40000元左右</option>
	<option value="15">50000元左右</option>
	<option value="16">60000元左右</option>
	<option value="17">100000元左右</option>
	<option value="18">150000元以上</option>
	<option value="19">200000元以上</option>
	<option value="20">300000元以上</option>
	<option value="21">500000元以上</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    住房情况 ：</td>
                                                <td>
                                                    <select name="ddl_House" id="ddl_House">
	<option selected="selected" value="">请选择</option>
	<option value="2">普通住房</option>
	<option value="3">中档住房</option>
	<option value="4">高档住房</option>
	<option value="5">豪华住房</option>
	<option value="6">郊区住房</option>
	<option value="7">农村住房</option>
	<option value="1">没有住房</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    车辆情况 ：</td>
                                                <td>
                                                    <select name="ddl_Car" id="ddl_Car">
	<option selected="selected" value="">请选择</option>
	<option value="1">无车</option>
	<option value="2">普通轿车</option>
	<option value="3">中级轿车</option>
	<option value="4">高级轿车</option>
	<option value="5">豪华轿车</option>
	<option value="6">其他车辆</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    兴趣爱好 ：</td>
                                                <td>
                                                    <select name="ddl_AiHao" id="ddl_AiHao">
	<option selected="selected" value="">请选择</option>
	<option value="1">语言文学</option>
	<option value="2">业务技术</option>
	<option value="3">爱好广泛</option>
	<option value="4">体育运动</option>
	<option value="5">休闲娱乐</option>
	<option value="6">人际交往</option>
</select>                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    性格特征 ：</td>
                                                <td>
                                                    <select name="ddl_Xingge" id="ddl_Xingge">
	<option selected="selected" value="">请选择</option>
	<option value="1">外向</option>
	<option value="2">内向</option>
	<option value="3">内外兼有</option>
</select>                                                </td>
                                                <td style="text-align: right">
                                                    文化修养 ：</td>
                                                <td>
                                                    <select name="ddl_Wenhua" id="ddl_Wenhua">
	<option selected="selected" value="">请选择</option>
	<option value="1">一般水准</option>
	<option value="2">中等水准</option>
	<option value="3">高等水准</option>
	<option value="4">专家水准</option>
</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    综合特质 ：</td>
                                                <td>
                                                    <select name="ddl_Zonghetezhi" id="ddl_Zonghetezhi">
	<option selected="selected" value="">请选择</option>
	<option value="1">体力型</option>
	<option value="2">智力型</option>
	<option value="3">体力加智力型</option>
	<option value="4">完美型</option>
	<option value="5">魅力型</option>
	<option value="6">领导型</option>
	<option value="7">职员型</option>
	<option value="8">艺术型</option>
</select></td>
                                                <td style="text-align: right">
                                                    特别说明 ：</td>
                                                <td>
                                                    <select name="ddl_TeBieshuoming" id="ddl_TeBieshuoming">
	<option selected="selected" value="">请选择</option>
	<option value="1">我的恋爱经历简单</option>
	<option value="2">我在等有缘人</option>
	<option value="3">历经风雨痴心不改</option>
	<option value="4">我没有谈过恋爱</option>
	<option value="5">我没有什么说的</option>
</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    形象素质 ：</td>
                                                <td>
                                                    <select name="ddl_Xingxiangsuzhi" id="ddl_Xingxiangsuzhi">
	<option selected="selected" value="">请选择</option>
	<option value="1">还算过得去</option>
	<option value="2">个人形象较好</option>
	<option value="3">个人素质较高</option>
	<option value="4">个人形象好</option>
	<option value="5">个人素质高</option>
	<option value="6">形象素质较高</option>
	<option value="7">形象素质高</option>
</select></td>
                                                <td style="text-align: right">
                                                    个人综述 ：</td>
                                                <td>
                                                    <select name="ddl_Gerenzongshu" id="ddl_Gerenzongshu">
	<option value="">请选择</option>
	<option value="1">我体貌才俱佳</option>
	<option value="2">我是美男子</option>
	<option value="3">我是美女子</option>
	<option value="4">我属于内慧型</option>
	<option value="5">我没有什么说的</option>
	<option value="6">你来发现我吧</option>
</select></td>
                                            </tr>
                                        </table>
                </P>
              </div>
              <div 
class="member_register_1_title"><span>择偶要求(请认真填写，这将有助于异性准确捕捉您择偶意向，提高婚恋成功率）</span></div>
              <div class="member_register_1_content">
                <p> </p>
                <table width="650" border="0" cellspacing="5">
                                            <tr>
                                                <td width="101" style="text-align: right">
                                                    年龄要求 ：</td>
                                                <td width="209">
                                                    <select name="ddl_YqBeginAge" id="ddl_YqBeginAge">
	<option value="0">请选择</option>
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

</select>岁 到
                                                    <select name="ddl_YqEndAge" id="ddl_YqEndAge">
	<option value="90">请选择</option>
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

</select>岁</td>
                                                <td width="110" style="text-align: right">
                                                    身高要求 ：</td>
                                                <td width="197">
                                                    <select name="ddl_YqHeightL" id="ddl_YqHeightL">
	<option value="0">请选择</option>
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
                                                    到
                                                    <select name="ddl_YqHeightH" id="ddl_YqHeightH">
	<option value="300">请选择</option>
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
                                                    车房要求 ：</td>
                                                <td>
                                                    <select name="ddl_YqHouseAndCar" id="ddl_YqHouseAndCar">
	<option selected="selected" value="">请选择</option>
	<option value="00">无要求</option>
	<option value="01">要求有房</option>
	<option value="10">要求有车</option>
	<option value="11">有房有车</option>

</select></td>
                                                <td style="text-align: right">
                                                    婚况要求 ：</td>
                                                <td>
                                                    <select name="ddl_YqIsMarry" id="ddl_YqIsMarry">
	<option selected="selected" value="">请选择</option>
	<option value="0">不做要求</option>
	<option value="1">要求未婚</option>
	<option value="11">可带孩子</option>
	<option value="2">可带女孩</option>
	<option value="12">不带男孩</option>
	<option value="4">离异未育</option>

</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    工作收入 ：</td>
                                                <td>
                                                    <select name="ddl_YqMonthIn" id="ddl_YqMonthIn">
	<option selected="selected" value="">不限</option>
	<option value="1">一般收入</option>
	<option value="2">中等收入</option>
	<option value="3">高等收入</option>

</select></td>
                                                <td style="text-align: right">
                                                    形象素质 ：</td>
                                                <td>
                                                    <select name="ddl_YqXingxiangsuzhi" id="ddl_YqXingxiangsuzhi" style="width:145px;">
	<option selected="selected" value="">请选择</option>
	<option value="1">过得去就行</option>
	<option value="2">个人形象较好</option>
	<option value="3">个人素质较高</option>
	<option value="4">个人形象好</option>
	<option value="5">个人素质高</option>
	<option value="6">形象素质较高</option>
	<option value="7">形象素质高</option>

</select></td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    学历要求 ：</td>
                                                <td>
                                                    <select name="ddl_YQEducation" id="ddl_YQEducation">
	<option selected="selected" value="0">不限</option>
	<option value="1">没有学历</option>
	<option value="2">小学以上</option>
	<option value="3">初中以上</option>
	<option value="4">高中以上</option>
	<option value="5">中专以上</option>
	<option value="6">大专以上</option>
	<option value="7">本科以上</option>
	<option value="8">研究生以上</option>
	<option value="9">博士</option>

</select></td>
                                                <td style="text-align: right">
                                                </td>
                                                <td>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    地区要求 ：</td>
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
											        document.getElementById("YQprovince").options[0].text="不限";
											        document.getElementById("YQcity").options[0].text="不限";
											        document.getElementById("YQarea").options[0].text="不限";
                                                    </script>

                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    其他择偶要求：</td>
                                                <td colspan="3">
                                                    <textarea name="txb_OtherCase" rows="2" cols="20" id="txb_OtherCase" class="input2" onkeydown="limitChars(this.id, 50)" onchange="limitChars(this.id, 50)" onpropertychange="limitChars(this.id, 50)" style="height:50px;width:75%;"></textarea>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="text-align: right">
                                                    其他需求 ：</td>
                                                <td colspan="3">
                                                    <table id="cbl_TSFW" border="0">
	<tr>
		<td><input name="cbl_TSFW$0" type="checkbox" id="cbl_TSFW_0"  value="1" />
		<label for="cbl_TSFW_0">我想参加相亲交友活动</label></td><td><input id="cbl_TSFW_1" type="checkbox" name="cbl_TSFW$1" value="1" /><label for="cbl_TSFW_1">我对异地姻缘感兴趣</label></td><td><input id="cbl_TSFW_2" type="checkbox" name="cbl_TSFW$2" value="1" /><label for="cbl_TSFW_6">我想请婚庆服务</label></td>
	</tr><tr>
		<td><input id="cbl_TSFW_3" type="checkbox" name="cbl_TSFW$3" value="1" /><label for="cbl_TSFW_3">我想结伴去旅游</label></td><td><input id="cbl_TSFW_4" type="checkbox" name="cbl_TSFW$4"  value="1"/><label for="cbl_TSFW_4">我想咨询婚恋顾问</label></td><td><input id="cbl_TSFW_5" type="checkbox" name="cbl_TSFW$5"  value="1"/>
		<label for="cbl_TSFW_5">我想加盟一世情缘</label></td>
	</tr><tr>
		<td><input id="cbl_TSFW_6" type="checkbox" name="cbl_TSFW$6" value="1" /><label for="cbl_TSFW_6">我想请猎头服务</label></td><td></td><td></td>
	</tr>
</table></td>
                                            </tr>
                                        </table>
                </P>
                <p>&nbsp;</p>
              </div>
              <div 
class="member_register_1_title"><span>爱情表白(请认真填写，让人了解的最简单话语，提高婚恋成功率）</span></div>
              <div class="member_register_1_content">
                <table cellspacing="5" width="630" border="0">
                  <tbody>
                    <tr>
                      <td valign="center" align="right" width="30" height="0">&nbsp;</td>
                      <td valign="center" colspan="2"><span style="COLOR: #ff0000">*</span><strong> 爱情表白:</strong>(200字内说说您内心世界、工作、理想和您的心中爱人,这将有助于异性了解你、靠近你！) <br /></td>
                    </tr>
                    <tr>
                      <td align="right" height="0">&nbsp;</td>
                      <td colspan="2" height="400" rowspan="2"><label>&nbsp;
                            <textarea class="input2" onkeypress="return noNumbers(event)" onpaste="" id="txb_SigNature" onkeydown="limitChars(this.id, 200)" style="WIDTH: 85%; HEIGHT: 120px" onpropertychange="limitChars(this.id, 200)" name="txb_SigNature" onchange="limitChars(this.id, 200)"></textarea>
                        </label>
                          <br />
                          <span id="RequiredFieldValidator6" 
      style="VISIBILITY: hidden; COLOR: red">请描述一下您的爱情表白</span></td>
                    </tr>
                    <tr>
                      <td align="right" height="0">&nbsp;</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div 
class="member_register_1_title"><span>上传头像(有照片的会员受到的关注率可提高30倍哦，提高婚恋成功率）</span></div>
              <div class="member_register_1_content">
                <table cellspacing="5" width="623" border="0">
                  <tbody>
                    <tr>
                      <td width="93" align="right">温馨提示：</td>
                      <td width="511" align="left">为了使您头像达到最佳显示效果，让别人清楚的看清你。<br />
                        请上传尺寸为宽120像素 * 
                        高145像素的照片否则图片会出现变形现象。</td>
                    </tr>
                    <tr>
                      <td align="right" >请上传头像：</td>
                      <td><input type="text" id="url1" value="" name="pic" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）</td>
                    </tr>
                     <tr>
                      <td align="right" >头像设置：</td>
                      <td><input type="radio" name="quanxian"  value="1" checked="checked"/>对所有人可见<input type="radio" name="quanxian"  value="1"/>仅对会员可见</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div class="member_register_1_title"><span>验证信息（请您认真阅读服务条款和隐私政策）</span></div>
              <div class="member_register_button">
                <table cellspacing="5" width="640" border="0">
                  <tbody>
                    <tr>
                      <td width="628" align="middle">验证码：
                       <input name="txtCode" type="text" id="txtCode" size="20" maxlength="20" />
                        <a href="#" onclick="refreshimg();"><img src="inc/Code.asp" id="getcode" alt="看不清楚？再换一张" border="0"/></a><span id="lbError" 
      style="COLOR: red"></span><span id="RequiredFieldValidator4" 
      style="VISIBILITY: hidden; COLOR: red">请填写验证码</span></td>
                    </tr>
                    <tr>
                      <td align="middle"><textarea class="input2" id="TextBox1" style="WIDTH: 85%; HEIGHT: 150px" name="TextBox1">1、您必须遵守中华人民共和国的各项法律、法规；遵守中华人民共和国有关互联网的各项法规、条例；遵守一世情缘交友中心网站章程。
2、您必须遵守互联网道德，互相尊重，以爱心对待他人；不发布或链接有关政治、淫秽色情封建迷信、人身攻击等违法信息；不得骚扰或欺骗其他会员；不可在网站内公布或传送任何诽谤、侮辱、不雅的内容，包括脏话和色情文字（包括照片）；不使用一世情缘从事商业广告或非法行为，否则一世情缘有权删除当事人资料，当事人同时承担由此直接或间接引起的刑事或民事法律责任。
3、您的会员档案里，呢称不可以含有您的电话号码、电子邮件、QQ号码、地址或其它可联络到您的资料。
4、您必须如实填写个人资料和上传本人的真实近照，如果发现有虚假行为，本站经证实后有权取消其会员资格，恕不退费，并追究其法律责任。
5、当您注册成功后，请您在一周之内携带真实有效的身份证、单身证明、学历证、工作证，原件和复印件各一份，1至5张个人真实近照到一世情缘的分站和婚姻经纪人处办理或传真至总部正式入会手续。

免责条款
1、一世情缘对于任何包含于、经由、或联结、下载或从任何与本网站有关服务(以下简称服务)所获得之资讯、内容或广告(以下简称资料)，不声明或保证其内容之正确性或可靠性；并且，对于您透过服务上之广告、资讯或要约而展示、购买或取得之任何产品、资讯或资料(以下简称产品)，本网站亦不负任何责任。
2、一世情缘之服务与资料是基于现状提供，而且一世情缘明确地表示拒绝对于服务、资料或产品给予任何明示或暗示之保证，包括但不限于，得为商业使用或适合于特定目的之保证。
3、一世情缘不担保服务一定能满足用户的要求，也不担保服务不会受中断，对服务的及时性、安全性、出错发生都不作担保。
4、用户以自己的独立判断从事与交友相关的行为，并独立承担可能产生的责任，一世情缘不承担任何法律责任。
5、一世情缘不对以下事项负责∶
1) 使用网站服务取得之结果正确可靠；
2）网站上所有照片的真实性；
2) 您经由网站服务购买或取得之任何产品、服务、资讯或其它信息将符合您的期望。
6、一世情缘对于本网站策划、发起、组织或是承办的任何会员活动(包括但不限于收取费用以及完全公益的活动)不担保一定能满足用户的需要，也不担保会员参加活动的结果或者是任何相关行动的正确性、合法合规性。由此产生的任何对于会员个人或者他人的人身或者是名誉以及其他损害，本网站不承担任何责任。
7、法律用户和一世情缘一致同意有关服务条款以及使用一世情缘的服务产生的争议交由南阳市仲裁委员会解决。若有任何服务条款与法律相抵触，那这些条款将按尽可能接近的方法重新解析，而其它条款则保持对用户产生法律效力和影响。
8、其他服务条款最终解释权和修改权归一世情缘所有。 </textarea></td>
                    </tr>
                    <tr>
                      <td align="middle"><input id="cb_Agree" onclick="SetSignbtnState();" 
      type="checkbox" checked="checked" name="cb_Agree" />
                          <label for="cb_Agree">已阅读和同意服务条款和隐私政策 </label></td>
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
                          <input type="submit" name="button" id="button" value="提交" />
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
                  <td height="89" align="left" class="css">一世情缘按照一定的顺序显示搜索结果，你可以通过下面的操作使自己的在搜索结果中排在前面：<br />
                    1、资料填写完整的用户会出现在比较靠前的位置<br />
                    2、有照片的用户将优先显示出来<br />
                    3、最后，记得经常来逛逛，比较高的活跃度能让更多人看到你，说不定你的TA也在其中。<br />
                    &gt; 找不到?让爱情顾问帮助您 &lt;</td>
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
                  <td width="68" align="right" class="hui12">情缘活动</td>
                  <td width="231" align="right" class="hui12"><a href="qqhd.asp" class="bai12a">&gt;&gt;更多</a></td>
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
                            <td height="21" align="left" class="xuxian STYLE1">[金凯悦] 12月16日  钻石男女 红酒派对 </td>
                          </tr>
                      </table>
                        <table width="287" cellspacing="0">
                          <tr>
                            <td height="21" align="left" class="xuxian STYLE1">[金凯悦] 12月16日  钻石男女 红酒派对 </td>
                          </tr>
                        </table>
                        <table width="287" cellspacing="0">
                          <tr>
                            <td height="21" align="left" class="xuxian STYLE1">[金凯悦] 12月16日  钻石男女 红酒派对 </td>
                          </tr>
                        </table>
                        <table width="287" cellspacing="0">
                          <tr>
                            <td height="21" align="left" class="xuxian STYLE1">[金凯悦] 12月16日  钻石男女 红酒派对 </td>
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
                <td width="68" align="right" class="hui12">成功故事</td>
                <td width="221" align="right" class="hui12"><a href="SuccessStory.asp" class="bai12a">&gt;&gt;更多</a><a href="news.asp" class="bai12"></a></td>
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
                <td align="center" class="css">幸福的开始</td>
                <td align="center"><span class="css">甜蜜的相知</span></td>
              </tr>
              <tr>
                <td colspan="2" align="left"><table width="287" cellspacing="0">
                  <tr>
                    <td width="19" height="20" align="center"><img src="images/c_06.png" width="12" height="10" /></td>
                    <td class="xuxian"><strong>李先生</strong> 和 <strong>申小姐</strong><a href="SuccessStory_con.asp?id=5" class="hui1">查看他们的 成功故事&gt;&gt;</a> </td>
                  </tr>
                </table>
                  <table width="287" cellspacing="0">
                    <tr>
                      <td width="19" height="21" align="center"><img src="images/c_06.png" width="12" height="10" /></td>
                      <td class="xuxian"><strong>陈先生</strong> 和 <strong>凌小姐</strong><a href="SuccessStory_con.asp?id=4" class="hui1">查看他们的 成功故事&gt;&gt;</a> </td>
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
                    <td width="72" height="40" align="center" class="hui14xx">婚庆礼仪</td>
                    <td width="212" align="right" ><a href="hqly.asp" class="bai12a" >&gt;&gt;更多</a></td>
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
                    <td width="140" class="css">经典的婚礼，是从高水准创意，婚礼策划开始的..</td>
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
msgbox"登陆邮箱不能为空，请重新输入!"
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
msgbox"请输入正确的邮箱!"
</script>
<%

end if 

sql="select * from [Users] where email='"&trim(request("txb_email"))&"'"
rs.open sql,conn,3,2
if  not rs.eof then
flag=0%>
<script language="vbscript">
msgbox"该邮箱已被人注册!"
</script><%
end if

rs.close

Password=request("txb_PassWord")
	rePassword=request("txb_rePassword")
	if Password="" then
flag=0%>
<script language="vbscript">
msgbox"密码不能为空!"
</script>
<%end if

	if Password<>rePassword then
flag=0%>
<script language="vbscript">
msgbox"两次密码不一致，请重新输入!"
</script>
<%end if 
response.Write("1")
	
	NickName=request("txb_NickName")
	
if NickName="" then
flag=0%>
<script language="vbscript">
msgbox"昵称不能为空!"
</script>
<%
response.Write("1")

else
	sql="select * from [Users] where NickName='"&trim(request("txb_NickName"))&"'"
rs.open sql,conn,3,2
if  not rs.eof then
flag=0%>
<script language="vbscript">
msgbox"昵称已被人使用!"
</script>
<%end if 
end if
response.Write("1")
rs.close
TelePhone=request("txb_TelePhone")
if TelePhone="" then
flag=0%>
<script language="vbscript">
msgbox"请填写联系电话!"
</script>
<%
elseif len(trim(request("txb_TelePhone")))<>11then
flag=0%>
<script language="vbscript">
msgbox"请输入正确的电话!"
</script><%

end if
response.Write("1")
Marry=request("ddl_Marry")
if  Marry="" then
flag=0%>
<script language="vbscript">
msgbox"请填写婚姻状况!"
</script><%

end if
response.Write("1")
BirthDay=request("txb_BirthDay")
if  BirthDay="" then
flag=0%>
<script language="vbscript">
msgbox"请填写出生年月!"
</script><%

end if
response.Write("1")
Education=request("ddl_Education")
if  Education="" then
flag=0%>
<script language="vbscript">
msgbox"请填写学历!"
</script><%
end if
response.Write("1")
province=request("mprovince")
city=request("mcity")
country=request("mcountry")
if  province="" or  city=""  or  country=""  then
flag=0%>
<script language="vbscript">
msgbox"请填写地区!"
</script><%
end if
response.Write("1")
SigNature=request("txb_SigNature")
if  SigNature="" then
flag=0
Response.Write("<script>alert(""请填写爱情告白"");history.back()</script>")
end if
response.Write("1")
txtCode=Trim(Request.Form("txtCode"))
If IsNumeric(txtCode)=False Or txtCode="" Or txtCode <> Session("GetCode") Then 
flag=0
Response.Write("<script>alert(""验证码输入错误"");history.back()</script>")
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
msgbox"信息添加成功!"
</script><%	
response.Redirect "login.asp"
end if


end if%>
</body>
</html>
