<!--#include file="function_custom.inc"-->
<div class="grid_4 right_col">
	<% round_bar("星雲法語") %>
	<div class="content">
		<p>
			<b>啟示:</b>
		</p>
		<p>人若自大，其實正顯示自己的無知
		人能謙虛，其實正表示自己的無求</p>
	</div>
	<% round_bar("最新活動") %>
	<div class="content">
		<ul>
		<%
    set activities = execute_sql("select top 5 * from activities order by date desc;")
    do while not activities.EOF
      rowHTMLStr = fmt("<li><a href='./activity.asp?id=%x'><span>%x</span></a></li>", Array(activities("ID_no"), activities("name")))
      Response.Write rowHTMLStr
      activities.MoveNext
    loop

    activities.Close
    set activities = Nothing
    %>
    <li><span>...</span></li>
		</ul>
	</div>
	
</div>