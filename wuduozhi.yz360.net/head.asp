<% 
Function cutStr(str,strlen) 
Dim re 
Set re=new RegExp 
re.IgnoreCase =True 
re.Global=True 
re.Pattern="<(.[^>]*)>" 
str=re.Replace(str,"") 
set re=Nothing 
Dim l,t,c,i 
l=Len(str) 
t=0 
For i=1 to l 
c=Abs(Asc(Mid(str,i,1))) 
If c>255 Then 
t=t+2 
Else 
t=t+1 
End If 
If t>=strlen Then 
cutStr=left(str,i)&"..." 
Exit For 
Else 
cutStr=str 
End If 
Next 
cutStr=Replace(cutStr,chr(10),"") 
cutStr=Replace(cutStr,chr(13)," ") 
cutStr=Replace(cutStr," ","") 
End Function 


  

%>
 <div class="head">
  <div class="top">
      <div class="topn"> 
       <div class="logo"><img src="images/logo_.png">     </div><!--logo-->
       
       <div class="topy">
        <img src="images/dianhua.png"> </div><!--topy-->
      </div>   <div class="clear"></div>   
   </div><!--top-->
       
    <div class="nav">                                                                                                
          <ul>
             <li <%if dateid=1 Then response.write("class='on'") end if %>><a href="index.asp"> <p>网站首页 </p></a></li>
             <li <%if dateid=2 Then response.write("class='on'") end if %>><a href="about.asp"> <p>关于我们 </p></a></li>
             <li <%if dateid=3 Then response.write("class='on'") end if %>><a href="product.asp?classid=2"> <p>产品中心 </p></a></li>
             <li <%if dateid=4 Then response.write("class='on'") end if %>><a href="news.asp?classid=6"> <p>新闻中心 </p></a></li>
             <li <%if dateid=5 Then response.write("class='on'") end if %>><a href="jiameng.asp"> <p>招商加盟 </p></a></li>
             <li <%if dateid=6 Then response.write("class='on'") end if %>><a href="vedio.asp"> <p>视频展示 </p></a></li>
             <li <%if dateid=7 Then response.write("class='on'") end if %>><a href="jiaru.asp"> <p>加入我们 </p></a></li>
             <li <%if dateid=8 Then response.write("class='on'") end if %>><a href="zxliuyan.asp"> <p>在线留言 </p></a></li>
             <li <%if dateid=9 Then response.write("class='on'") end if %>><a href="lianxi.asp"> <p>联系我们 </p></a></li>
           </ul>
   <div class="clear"></div>
       </div>
     

 
 <div class="banner-box">
	<div class="bd">
        <ul>          	    
             <li  style="background:url(images/banner.5.jpg) center  no-repeat "><a href="#"></a></li>    
             <li  style="background:url(images/banner.6.jpg) center  no-repeat"><a href="#"></a></li>   
             <li  style="background:url(images/banner.7.jpg) center  no-repeat"><a href="#"></a></li>  
              <li  style="background:url(images/banner.8.jpg) center  no-repeat"><a href="#"></a></li>   
        </ul>
    </div>
   
    <div class="banner-btn">
        <a class="prev" href="javascript:void(0);"></a>
        <a class="next" href="javascript:void(0);"></a>
        <div class="clear"></div>
        <div class="hd"><ul></ul></div>
    </div>
    </div>
    
    <script type="text/javascript">
$(document).ready(function(){
/*鼠标移过，左右按钮显示*/
	$(".prev,.next").hover(function(){
		$(this).stop(true,false).fadeTo("show",0.9);
	},function(){
		$(this).stop(true,false).fadeTo("show",0.4);
	});
	
	$(".banner-box").slide({
		titCell:".hd ul",
		mainCell:".bd ul",
		effect:"fold",
		interTime:2800,
		delayTime:500,
		autoPlay:true,
		autoPage:true, 
		trigger:"click" 
	});
});
</script>
<div class="clear"></div></div>

  
  </div><!--head-->