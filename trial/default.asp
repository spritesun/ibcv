<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_custom.inc"-->
</div>
<script type="text/javascript">

var currentActivity = "";

function displayActivity(name)
{
  if (currentActivity === name) {
    return;
  }
  if (currentActivity !== "") {
    Effect.Fade(currentActivity);
  }
  currentActivity = name;
  Effect.Appear(currentActivity);
}

var currentNews = "";

function displayNews(name)
{
  if (currentNews === name) {
    return;
  }
  if (currentNews !== "") {
    Effect.Fade(currentNews);
  }
  currentNews = name;
  Effect.Appear(currentNews);
}

</script>


<div class="grid_4">
	<% round_bar("精彩照片") %>
	<div id="slide-show">
        <ul id="slide-images">
        
<%
'TODO: could improve to be a random way
dim activities, rowHTMLStr

set activities = execute_sql("select top 10 * from activities where image1_file_name <> '' order by date desc;")
do while not activities.EOF
  rowHTMLStr = fmt("<li><a href='../uploads/%x' title='<a href=./activity.asp?id=%x>%x</a>' rel='lightbox-cats'><img src='../uploads/%x' alt='%x' width='215' height='160'/></a></li>", Array(activities("image1_file_name"), activities("ID_no"), activities("name"), activities("image1_file_name"), activities("image1_file_name")))
  Response.Write rowHTMLStr
  activities.MoveNext
loop

activities.Close
set activities = Nothing

dim news

set news = execute_sql("select top 10 * from news where image1_file_name <> '' order by date desc;")
do while not news.EOF
  rowHTMLStr = fmt("<li><a href='../uploads/%x' title='<a href=./news.asp?id=%x>%x</a>' rel='lightbox-cats'><img src='../uploads/%x' alt='%x' width='215' height='160'/></a></li>", Array(news("image1_file_name"), news("ID_no"), news("name"), news("image1_file_name"), news("image1_file_name")))
  Response.Write rowHTMLStr
  news.MoveNext
loop

news.Close
set news = Nothing
%>

      <!--#include file="includes/old_image_carrousel.asp"-->
        </ul>
        <DIV style="VERTICAL-ALIGN: top;font-size:1.3em; PADDING-TOP: -1.0em; padding-bottom:1em; TEXT-ALIGN: center">&nbsp;<span id="slideImg-text">1 of 11</span>&nbsp; 點擊看大圖&nbsp;</DIV>        
    </div>
	<% round_bar("佛光視頻") %>
	<div class="content">
	<object width="215" height="180"><param name="movie" value="http://www.youtube.com/v/ojKz40yuhew?fs=1&amp;hl=en_US"></param>
	<param name="allowFullScreen" value="true"></param><param name="allowscriptaccess" value="always"></param>
	<embed src="http://www.youtube.com/v/ojKz40yuhew?fs=1&amp;hl=en_US" type="application/x-shockwave-flash" allowscriptaccess="always" allowfullscreen="true" width="215" height="180"></embed></object>
	</div>	
	<% round_bar("人間福報") %>
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
  <% round_bar("活動預告") %>  
	<div class="content">
		<ul>
		<%
    set activities = execute_sql("select top 10 * from activities order by date desc;")
    dim isFirst
    isFirst = True
    do while not activities.EOF
      dim style
      if isFirst then
        style = "block"
        Response.Write("<script type='text/javascript'> currentActivity = 'activity_" & activities("ID_no") & "';</script>")
        isFirst = False
      else
        style = "none"
      end if

      rowHTMLStr = fmt("<li><a href='./activity.asp?id=%x' onmouseover=""displayActivity('activity_%x');""><span>%x</span>%x</a><br /><br /><div id='activity_%x' style='display:%x;'>%x</div></li>", Array(activities("ID_no"), activities("ID_no"), activities("date"), activities("name"), activities("ID_no"),style, activities("summary") & "..."))
      Response.Write rowHTMLStr
      activities.MoveNext
    loop

    activities.Close
    set activities = Nothing
    %>
		</ul>
	</div>
	
	<% round_bar("活動新聞") %>  

	<div class="content">
		<ul>
		<%
    set news = execute_sql("select top 10 * from news order by date desc;")
    dim isFirstNews
    isFirstNews = True
    do while not news.EOF
      dim newsStyle
      if isFirstNews then
        newsStyle = "block"
        Response.Write("<script type='text/javascript'> currentNews = 'news_" & news("ID_no") & "';</script>")
        isFirstNews = False
      else
        newsStyle = "none"
      end if

      rowHTMLStr = fmt("<li><a href='./news.asp?id=%x' onmouseover=""displayNews('news_%x');""><span>%x</span>%x</a><br /><br /><div id='news_%x' style='display:%x;'>%x</div></li>", Array(news("ID_no"), news("ID_no"), news("date"), news("name"), news("ID_no"),newsStyle, news("summary") & "..."))
      Response.Write rowHTMLStr
      news.MoveNext
    loop

    news.Close
    set news = Nothing
    %>

		</ul>
	</div>	
</div>	

<div class="grid_4 right_col">
    <div class="content">
      <!--#include file="includes/links.asp"-->
    </div>
</div>
<!--#include file="includes/footer.inc"-->	
