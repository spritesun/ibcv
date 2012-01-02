<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	

<div class="grid_14 prefix_1 suffix_1">
<%
dim activity
set activity = execute_sql("select * from activities where ID_no=" & Request.QueryString("id"))
%>
  <form action="./admin/save_activity.asp" method="post">
    <input type="hidden" name="id" value="<%=Request.QueryString("id")%>"/>
    <h2>通告名稱</h2>
    <input type="text" name="name" value="<%=activity("name")%>"/>
    <h2>日期</h2>             
    <input type="text" name="date" value="<%=activity("date")%>"/>
    <h2>時間</h2>             
    <input type="text" name="time" value="<%=activity("time")%>"/>
    <h2>地點</h2>             
    <input type="text" name="venue" value="<%=activity("venue")%>"/>
    <h2>内容介紹</h2>
    <textarea rows="20" cols="60" name="content"><%=activity("content")%></textarea>
    <input type="submit" value="提交通告"/>
  </form>

</div>

<!--#include file="../includes/footer.inc"-->