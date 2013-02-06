<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_custom.inc"-->

<div class="grid_14 prefix_1 suffix_1">
	<div class="content">
<%
if Request.QueryString("venue") <> "" then
%>
<h1><%= Request.QueryString("venue")%> 社教課程</h1>

<table>
<thead>
  <tr>
    <td>活動<br />Activities</td>
    <td>日期<br />Date</td>
  </tr>
</thead>
<tbody>
<%
dim activities
set activities = execute_sql("select * from activities where category = 'course' and venue ='" & Request.QueryString("venue") & "' order by date desc;")
do while not activities.EOF
  dim rowHTMLStr
  rowHTMLStr = fmt("<tr><td><a href='./activity.asp?id=%x'>%x</a></td><td>%x</td></tr>", Array(activities("ID_no"), activities("name"), activities("date")))
  Response.Write rowHTMLStr
  activities.MoveNext
loop

activities.Close
set activities = Nothing
%>
</tbody>
</table>	  

<h2><a href="./courses.asp">«返回</a></h2>
<%
else
%>
	  <h2><a href="./courses.asp?venue=墨爾本佛光山 Yarraville">墨爾本佛光山(Yarraville)社教課程</a></h2>
	  <h2><a href="./courses.asp?venue=佛光山尓有寺"/>墨爾本佛光山尔有寺(Box Hill)社教課程</a></h2>
	  <h2><a href="./courses.asp?venue=墨爾本佛光緣美術館">墨爾本佛光緣美術館(Queen St)社教課程</a></h2>
<%
end if
%>

  </div>
</div>
<!--#include file="includes/footer.inc"-->