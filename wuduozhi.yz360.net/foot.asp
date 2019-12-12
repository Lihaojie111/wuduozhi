
  
  
  
  <div class="youqing"><p>友情链接： <%
  Set Rs=Server.CreateObject("ADODB.Recordset")
Sql = "Select Top 14  * From [yqlj]  Order By ID desc,DateAndTime  desc"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof
%> <a target="_blank" href="<%=rs("wzdz")%>" ><%=rs("wzmc")%></a>　　<%

Rs.MoveNext
Loop
Rs.Close
%></p></div><!--友情链接-->


 <div class="foot">
           
    <div class="footn"> <p style=" float:left;">Copyright  2015-2016 河南省五垛源中药材有限责任公司  版权所有 ICP备案证书号：<a href="http://www.beian.miit.gov.cn" target="_blank">豫ICP备16000736号-1</a></p><p style="float:right;">
    <a href="http://yz360.com.cn">技术支持：南阳永正</a></p>
    </div><!--footnr-->
</div><!--foot-->