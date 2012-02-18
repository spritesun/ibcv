<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_custom.inc"-->

<div class="grid_10 prefix_1 suffix_1">
	<div class="content">
	  <h1>最新活動預告</h1>
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
    set activities = execute_sql("select * from activities;")

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
	  
  	<h1>過往活動</h1>
    <ul>
      <li><a href="./upcoming.asp">2011 法會佛事</a><li>
      <li><a href="./Yarraville/activities.asp">2010 墨爾本佛光山法會佛事</a></li>
      <li><a href="./ArtGallery/activities.asp">2010 美術館法會佛事</a></li>
      <li><a href="./BoxHill/activities.asp">2010 博士山法會佛事</a></li>
      <li><a href="./Yarraville/class-2010.asp">2010 墨爾本佛光山才藝班</a></li>
      <li><a href="./ArtGallery/class-2010.asp">2010 美術館才藝班</a></li>
      <li><a href="./BoxHill/class-2010.asp">2010 博士山才藝班</a></li>
    </ul>
  </div>
</div>
<!--#include file="includes/footer.inc"-->