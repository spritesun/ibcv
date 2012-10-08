<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	

<div class="grid_14 prefix_1 suffix_1">

<h1>通告列表</h1>

<table>
<thead>
  <tr>
    <td>活動</td>
    <td>類别</td>
    <td>日期</td>
    <!--td>時間</td-->
    <td>地點</td>
    <td>操作</td>
  </tr>
</thead>
<tbody>
<!--#include file="../includes/function_fmt.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<%
dim activities, rowHTMLStr, category
set activities = execute_sql("select * from activities order by date desc;")

do while not activities.EOF
  category = activities("category")
  if category = "ceremony" then
    category = "法會活動"
  elseif category = "course" then
    category = "社教課程"
  end if
  
  rowHTMLStr = fmt("<tr><td>%x</td><td>%x</td><td>%x</td><td>%x</td><td><a href='./activity.asp?id=%x'>查看</a> <a href='./admin/edit_activity.asp?id=%x'>修改</a> <a href='./admin/delete_activity.asp?id=%x'>删除</a></td></tr>", Array(activities("name"), category, activities("date"), activities("venue"), activities("ID_no"), activities("ID_no"), activities("ID_no")))
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