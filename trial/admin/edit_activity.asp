<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	

<div class="grid_14 prefix_1 suffix_1">
<%
dim activity, name, date_, time_, venue, content
if (Request.QueryString("id")<>"") then
  set activity = execute_sql("select * from activities where ID_no=" & Request.QueryString("id"))
  name = activity("name")
  date_ = activity("date")
  time_ = activity("time")
  venue = activity("venue")
  content = activity("content")
end if
%>

  <form action="./admin/save_activity.asp" method="post" enctype="multipart/form-data"  accept-charset="utf-8">
    <input type="hidden" name="id" value="<%=Request.QueryString("id")%>"/>
    <h2>通告名稱</h2>
    <input type="text" name="name" value="<%=name%>"/>
    <h2>日期</h2>
    <input type="text" name="date" value="<%=date_%>"/>
    <h2>時間</h2>
    <input type="text" name="time" value="<%=time_%>"/>
    <h2>地點</h2>
    <input type="text" name="venue" value="<%=venue%>"/>
    <h2>内容介紹</h2>
    <textarea rows="20" cols="60" name="content"><%=content%></textarea>
    <h2>上傳/修改通告圖片</h2>
    <input type="file" name="image1"/>

    <%
    if activity("image1_file_name") <> "" then
    %>
    <div class="edit_image">
      <h2>删除舊的通告圖片</h2>
      <img src="../uploads/<%=activity("image1_file_name")%>" alt="activity_image1"></img>
      <a href="./admin/delete_activity_image.asp?id=<%=Request.QueryString("id")%>">删除此圖片</a>
    </div>
    <%
    end if
    %>

    <input type="submit" value="提交通告"/>
  </form>

</div>

<%
set activity = Nothing
%>
<!--#include file="../includes/footer.inc"-->