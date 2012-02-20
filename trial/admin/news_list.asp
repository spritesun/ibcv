<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	
<!--#include file="../includes/function_custom.inc"-->	

<div class="grid_14 prefix_1 suffix_1">

<h1>新聞列表</h1>

<table>
<thead>
  <tr>
    <td>新聞</td>
    <td>日期</td>
    <td>操作</td>
  </tr>
</thead>
<tbody>

<%
dim news
set news = execute_sql("select * from news order by date desc;")

do while not news.EOF
  dim rowHTMLStr
  rowHTMLStr = fmt("<tr><td>%x</td><td>%x</td><td><a href='./news.asp?id=%x'>查看</a> <a href='./admin/edit_news.asp?id=%x'>修改</a> <a href='./admin/delete_news.asp?id=%x'>删除</a></td></tr>", Array(news("name"), news("date"), news("ID_no"), news("ID_no"), news("ID_no")))
  Response.Write rowHTMLStr
  news.MoveNext
loop

news.Close
set news = Nothing
%>
</tbody>
</table>

</div>

<!--#include file="../includes/footer.inc"-->