// JavaScript Document





$(function(){
	
	
	



//婚礼展示  婚礼视频  点击展开更多js
$('.ht_show_click').click(function(){
	if($(this).parents('.ht_show_title').siblings('.ht_showvideo_box').height()==237 && $(this).parents('.ht_show_title').siblings('.ht_showvideo_box').find('.ht_showvideo_box_cen_one').size() > 4)
	{
	$(this).addClass('ht_hover_click');
	$(this).parents('.ht_show_title').siblings('.ht_showvideo_box').height('auto');
	$('body,html').animate({scrollTop:$(this).parents('.ht_show_title').siblings('.ht_showvideo_box').offset().top+$(this).parents('.ht_show_title').siblings('.ht_showvideo_box').find('.ht_showvideo_box_cen_one').height()},1000);
	}
	else{
		$(this).removeClass('ht_hover_click');	
		$(this).parents('.ht_show_title').siblings('.ht_showvideo_box').height(237);	
		}
})





//婚礼展示  视频展示 鼠标悬浮效果

$('.ht_showvideo_box_cen_one').hover(
		function(){
			$(this).find('.ht_video_show').stop(true,true).show("slow");
		},
		function(){
			$(this).find('.ht_video_show').stop(true,true).hide("slow");	
		}
)


	
	

//婚礼展示   才艺展示  点击加号  展开下来特效js

$('.ht_show_click').click(function(){
	if($(this).parents('.ht_show_title').siblings('.ht_show_box').height()==592)
	{
	$(this).addClass('ht_hover');
	$(this).parents('.ht_show_title').siblings('.ht_show_box').height('auto');
	$('body,html').animate({scrollTop:$(this).parents('.ht_show_title').siblings('.ht_show_box').offset().top+$(this).parents('.ht_show_title').siblings('.ht_show_box').find('.ht_show_box_one').height()},1000);
	}
	else{
		$(this).removeClass('ht_hover');	
		$(this).parents('.ht_show_title').siblings('.ht_show_box').height(592);	
		}
})



//内页menu  滑动效果 


$('.ht_nei_menu ul li').mouseover(
	function(){
	var ht_nei_hoverNUB = $('.ht_nei_menu ul li').index($(this));
	var ht_nei_hoverwidth = ($('.ht_nei_menu ul li').width()+2)*ht_nei_hoverNUB;
	$('.ht_nei_menu_hover').stop().animate({left:ht_nei_hoverwidth},400);
	$('.ht_nei_menu ul li.ht_hover').addClass('ht_hover_one');
	} 
)

$('.ht_nei_menu').mouseleave(function(){
	var ht_nei_hoverNUB_index = $('.ht_nei_menu ul li').index($('.ht_nei_menu ul li.ht_hover'));
	$('.ht_nei_menu_hover').stop().animate({left:($('.ht_nei_menu ul li').width()+2)*ht_nei_hoverNUB_index},400);
	$('.ht_nei_menu ul li.ht_hover').removeClass('ht_hover_one');
})	


	


//内页下载列表，点击展开效果

$('.ht_nei_gongju_one .ht_nei_gongju_one_top span').click(function(){

	$(this).parents('.ht_nei_gongju_one_top').siblings('.ht_nei_gongju_one_bot').slideToggle(200);
	$(this).parents('.ht_nei_gongju_one').toggleClass('ht_nei_gongju_hover');
})


	
	
	
//导航js
$('.ht_menu_fa').hover(function(){
	$(this).find('.ht_menu_chi').stop(false,true).slideToggle('slow');	
})	
$('.ht_menu_chi').hover(function(){
	$(this).stop(false,true).show();	
})		


//scrollTop



$(window).scroll(function(){
	
	var bodydocument = document.documentElement.scrollTop || window.pageYOffset || document.body.scrollTop;
	if(bodydocument>=100)
	{
	 $('.ht_scroll_top').slideDown('slow');
	}
	else
	{
	$('.ht_scroll_top').slideUp('slow');	
	}	
	
})





//婚礼展示 js

$('.ht_hunlizhanshi_box_one').hover(

		function(){
		$(this).find('.ht_fc_box').stop(true,true).show('slow');
		$(this).find('.ht_fc').stop(true,true).show('slow');
		},
		function(){
		$(this).find('.ht_fc_box').stop(true,true).hide('slow');
		$(this).find('.ht_fc').stop(true,true).hide('slow');
		}
		
)	
	



//婚宴酒店js

$('.ht_hunyanjiudian_box_cen_one').hover(function(){
	
	$(this).find('.ht_hovertext').stop(true,true).slideToggle('slow');
	
})	
	
	
	
	
	
	
	
	
//footer 底部微信  微博  联系  鼠标悬浮效果

$('.ht_footer_q').hover(function(){
	$(this).find('.ht_footer_q_box').stop(true,true).slideToggle('slow');	
})

$('.ht_footer_q_box').mouseover(function(){
	$(this).stop(true,true).show();	
})		
	
	
	








//婚宴酒店   轮播js

$('.ht_hunyanjiudian_box_cen').scrollbox({
    direction: 'h',
    distance: 283
  });
  $('.ht_hunyanjiudian_box_le').click(function () {
    $('.ht_hunyanjiudian_box_cen').trigger('backward');
  });
  $('.ht_hunyanjiudian_box_ri').click(function () {
    $('.ht_hunyanjiudian_box_cen').trigger('forward');
  });	

	
	
	
	
	
	})