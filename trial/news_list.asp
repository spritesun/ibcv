<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_custom.inc"-->	
<div class="grid_10 prefix_1 suffix_1">
	<div class="content">
    <h1>活動新聞</h1>
    
    <div id="newsLists">
      <ul>
<%
dim news
set news = execute_sql("select * from news order by date desc;")

do while not news.EOF
%>
<li>
<div class="newsItem">

  <div class="newsImgContainer">
  <% if news("image1_file_name") <> "" then %>
  <a href="news.asp?id=<%=news("ID_no")%>">
  <img src="../uploads/<%=news("image1_file_name")%>" class="newsImg" alt="image1"/></a>
  </a>
  <% else %>
  <div class="newsImg" style="height:1px;"/></div>
  <% end if %>
  </div>

  <div class="newsDescr"><a href="news.asp?id=<%=news("ID_no")%>"><%=news("name")%></a><br /><br /><%=news("summary")%>...</div>
</div>
</li>
<%
  news.MoveNext
loop

news.Close
set news = Nothing
%>
      </ul>
    </div>
    <div class="clear"></div>
    <h1>過往活動</h1>
    <ul>
      <li><a href="./whatsnew.asp">2009 活動新聞</a></li>
    </ul>
  </div>
</div>

<!--#include file="includes/right-col.inc"-->	

<!--#include file="includes/footer.inc"-->
