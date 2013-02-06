<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_custom.inc"-->

<div class="grid_14 prefix_1 suffix_1">
	<div class="content">
	  <h1>社教課程</h1>
	  <table>
    <thead>
      <tr>
        <td>活動<br />Activities</td>
        <td>日期<br />Date</td>
        <td>地點<br />Venue</td>
      </tr>
    </thead>
    <tbody>
    <%
    dim activities
    set activities = execute_sql("select * from activities where category = 'course' order by date desc;")

    do while not activities.EOF
      dim rowHTMLStr
      rowHTMLStr = fmt("<tr><td><a href='./activity.asp?id=%x'>%x</a></td><td>%x</td><td>%x</td></tr>", Array(activities("ID_no"), activities("name"), activities("date"), activities("venue")))
      Response.Write rowHTMLStr
      activities.MoveNext
    loop

    activities.Close
    set activities = Nothing
    %>
    </tbody>
    </table>	  
  </div>
</div>
<!--#include file="includes/footer.inc"-->