<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_custom.inc"-->

<div class="grid_14 prefix_1 suffix_1">
	<div class="content">
	  <h1>最新法會活動</h1>
	  <table>
    <thead>
      <tr>
        <td>活動<br />Activities</td>
        <td>日期<br />Date</td>
        <td>時間<br />Time</td>
        <td>地點<br />Venue</td>
      </tr>
    </thead>
    <tbody>
    <%
    dim activities
    set activities = execute_sql("select * from activities where category = 'ceremony' order by date desc;")

    do while not activities.EOF
      dim rowHTMLStr
      rowHTMLStr = fmt("<tr><td><a href='./activity.asp?id=%x'>%x</a></td><td>%x</td><td>%x</td><td>%x</td></tr>", Array(activities("ID_no"), activities("name"), activities("date"), activities("time"), activities("venue")))
      Response.Write rowHTMLStr
      activities.MoveNext
    loop

    activities.Close
    set activities = Nothing
    %>
    </tbody>
    </table>
	  
  	<h1>過往法會</h1>
    <ul>
      <li><a href="./upcoming.asp">2011 法會佛事</a><li>
      <li><a href="./Yarraville/activities.asp">2010 墨爾本佛光山法會佛事</a></li>
      <li><a href="./ArtGallery/activities.asp">2010 美術館法會佛事</a></li>
      <li><a href="./BoxHill/activities.asp">2010 博士山法會佛事</a></li>
    </ul>
  </div>
</div>
<!--#include file="includes/footer.inc"-->