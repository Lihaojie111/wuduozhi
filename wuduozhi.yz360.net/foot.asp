
  
  
  
  <div class="youqing"><p>�������ӣ� <%
  Set Rs=Server.CreateObject("ADODB.Recordset")
Sql = "Select Top 14  * From [yqlj]  Order By ID desc,DateAndTime  desc"            
Rs.Open Sql,Conn,1,1

Do While Not Rs.Eof
%> <a target="_blank" href="<%=rs("wzdz")%>" ><%=rs("wzmc")%></a>����<%

Rs.MoveNext
Loop
Rs.Close
%></p></div><!--��������-->


 <div class="foot">
           
    <div class="footn"> <p style=" float:left;">Copyright  2015-2016 ����ʡ���Դ��ҩ���������ι�˾  ��Ȩ���� ICP����֤��ţ�<a href="http://www.beian.miit.gov.cn" target="_blank">ԥICP��16000736��-1</a></p><p style="float:right;">
    <a href="http://yz360.com.cn">����֧�֣���������</a></p>
    </div><!--footnr-->
</div><!--foot-->