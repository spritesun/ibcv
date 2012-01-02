<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	

<div class="grid_14 prefix_1 suffix_1">

<h1>通告列表</h1>

<table>
<thead>
  <tr>
    <td>活動<br />Activities</td>
    <td>日期<br />Date</td>
    <td>時間<br />Time</td>
    <td>地點<br />Venue</td>
    <td>操作</td>
  </tr>
</thead>
<tbody>
<!--#include file="../includes/function_fmt.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<%
dim activities
set activities = execute_sql("select * from activities;")

do while not activities.EOF
  dim rowHTMLStr
  rowHTMLStr = fmt("<tr><td>%x</td><td>%x</td><td>%x</td><td>%x</td><td><a href='./activities.asp?id=%x'>查看</a> <a href='./admin/edit_activities.asp?id=%x'>修改</a> <a href='./admin/delete_activity.asp?id=%x'>删除</a></td></tr>", Array(activities("name"), activities("date"), activities("time"), activities("venue"), activities("ID_no"), activities("ID_no"), activities("ID_no")))
  Response.Write rowHTMLStr
  activities.MoveNext
loop

activities.Close
set activities = Nothing
%>
</tbody>
</table>

</div>

<!--#include file="../includes/footer.inc"-->	