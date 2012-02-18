<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_custom.inc"-->
</div>
<script type="text/javascript">

    var currentNews = 'buddhaLightHuiXiang';

function displayNews(name)
{
    Effect.Fade(currentNews);
    currentNews = name;
    Effect.Appear(currentNews);
}

</script>


<div class="grid_4">
	<img src="img/photo-heading.png">
	<div id="slide-show">
        <ul id="slide-images">
        <li><a href="img/photo/143D.jpg" rel="lightbox-cats" style="width:215px;height:143px;"><img src="img/photo/143D.jpg" alt="143D" width="215" height="143"/></a>     </li>
        <li><a href="img/photo/143E.jpg" rel="lightbox-cats" style="width:215px;height:161px;"><img src="img/photo/143E.jpg" alt="143E" width="215" height="161"/></a>     </li>
        
<%
dim activities, rowHTMLStr

set activities = execute_sql("select * from activities where image1_file_name <> '';")
do while not activities.EOF
  rowHTMLStr = fmt("<li><a href='../uploads/%x' title='<a href=./activity.asp?id=%x>%x</a>' rel='lightbox-cats'><img src='../uploads/%x' alt='%x' width='215' height='160'/></a></li>", Array(activities("image1_file_name"), activities("ID_no"), activities("name"), activities("image1_file_name"), activities("image1_file_name")))
  Response.Write rowHTMLStr
  activities.MoveNext
loop

activities.Close
set activities = Nothing
%>
          <li><a href="img/photo/143F.jpg" rel="lightbox-cats" style="width:215px;height:120px;"><img src="img/photo/143F.jpg" alt="143F" width="215" height="120"/></a>     </li>
          <li><a href="img/photo/1440.jpg" rel="lightbox-cats" style="width:215px;height:161px;"><img src="img/photo/1440.jpg" alt="1430" width="215" height="161"/></a>     </li>
          <li><a href="img/photo/1441.jpg" rel="lightbox-cats" style="width:215px;height:143px;"><img src="img/photo/1441.jpg" alt="143D" width="215" height="143"/></a>     </li>
          <li><a href="img/photo/1442.jpg" rel="lightbox-cats" style="width:215px;height:120px;"><img src="img/photo/1442.jpg" alt="143D" width="215" height="120"/></a>     </li>
          <li><a href="img/photo/1443.jpg" rel="lightbox-cats" style="width:215px;height:120px;"><img src="img/photo/1443.jpg" alt="143D" width="215" height="120"/></a>     </li>
          <li><a href="img/photo/1444.jpg" rel="lightbox-cats" style="width:215px;height:142px;"><img src="img/photo/1444.jpg" alt="143D" width="215" height="142"/></a>     </li>
          <li><a href="img/photo/1445.jpg" rel="lightbox-cats" style="width:215px;height:120px;"><img src="img/photo/1445.jpg" alt="143D" width="215" height="120"/></a>     </li>
          <li><a href="img/photo/1446.jpg" rel="lightbox-cats" style="width:215px;height:143px;"><img src="img/photo/1446.jpg" alt="143D" width="215" height="143"/></a>     </li>
          <li><a href="img/photo/1447.jpg" rel="lightbox-cats" style="width:215px;height:143px;"><img src="img/photo/1447.jpg" alt="143D" width="215" height="143"/></a>     </li>
        </ul>
        <DIV style="VERTICAL-ALIGN: top;font-size:1.3em; PADDING-TOP: -1.0em; padding-bottom:1em; TEXT-ALIGN: center">&nbsp;<span id="slideImg-text">1 of 11</span>&nbsp; 點擊看大圖&nbsp;</DIV>        
    </div>
	<img src="img/video.png" />
	<div class="content">
	<object width="215" height="180"><param name="movie" value="http://www.youtube.com/v/ojKz40yuhew?fs=1&amp;hl=en_US"></param>
	<param name="allowFullScreen" value="true"></param><param name="allowscriptaccess" value="always"></param>
	<embed src="http://www.youtube.com/v/ojKz40yuhew?fs=1&amp;hl=en_US" type="application/x-shockwave-flash" allowscriptaccess="always" allowfullscreen="true" width="215" height="180"></embed></object>
	</div>	
	<img src="img/magazine.png" />
	<div>
	<a href="" onclick="return open_EPAPER();"><img id="Epaper1_Image1" title="電子報紙導覽" src="http://www.merit-times.com.tw/image/epaper/epaper_20101231.jpg" style="width:215px;border-width:0px;padding-top:10px;padding-bottom:20px;" /></a> 
<script language="javascript" type="text/javascript" > 
       function open_EPAPER()
        {
            EPAPER_width=window.screen.availwidth-10
            EPAPER_height=window.screen.availheight-25
            window.open("http://www.merit-times.com.tw/newsepaper/","WindowPopup","width="+EPAPER_width+", height=800px,,top=0,left=0,fullscreen=yes"); 
            // window.open("http://www.merit-times.com.tw/newsepaper/","WindowPopup","width="+EPAPER_width+", height=800px,,top=0,left=0, resizable"); 
            // window.open ("http://www.merit-times.com.tw/newsepaper/","epaper", "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width="+EPAPER_width+",height=800,top=0,left=0 ");
            // window.open ("http://www.merit-times.com.tw/newsepaper/","", "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbar=no,resizable=no,width="+EPAPER_width+",height="+EPAPER_height+",top=0,left=0");
        }         
        
</script> 
	</div>
</div>

<div class="grid_8 latest_news">
	<img src="img/zuixin-heading-long.png">
	<div class="content">
		<ul>
		<%
    set activities = execute_sql("select * from activities;")
    do while not activities.EOF
      rowHTMLStr = fmt("<li><a href='./activity.asp?id=%x' onmouseover=""displayNews('activity_%x');""><span>%x</span>%x</a><br /><br /><div id='activity_%x' style='display:none;'>%x</div></li>", Array(activities("ID_no"), activities("ID_no"), activities("date"), activities("name"), activities("ID_no"), Left(activities("content"), 400) & "..."))
      Response.Write rowHTMLStr
      activities.MoveNext
    loop

    activities.Close
    set activities = Nothing
    %>
	  <!--#include file="includes/old_activities.asp"-->
		</ul>
	</div>
	<img src="img/news-heading.png">
	<div class="content">
		<ul>
		</ul>
	</div>	
</div>	

<div class="grid_4 right_col">
    <div class="content">
      <!--#include file="includes/links.asp"-->
    </div>
</div>
<!--#include file="includes/footer.inc"-->	
