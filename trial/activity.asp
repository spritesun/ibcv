<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_execute_sql.inc"-->	
<%
dim activity
set activity = execute_sql("select * from activities where ID_no=" & Request.QueryString("id"))

dim text_grid_size
if activity("image1_file_name") <> "" then
  text_grid_size = "grid_9"
else
  text_grid_size = "grid_14"
end if
%>

<div class="<%=text_grid_size%> prefix_1">
  <div class="content">
    <h1><%=activity("name")%></h1>
    <p>時間: <%=activity("date")%>， <%=activity("time")%></p>
    <p>地點: <%=activity("venue")%></p>
    <br />
    <p><%=activity("content")%></p>
  </div>
</div>

<%
if activity("image1_file_name") <> "" then
%>
<div class="grid_5">
  <div class="content">
    <a href="../uploads/<%=activity("image1_file_name")%>" rel="lightbox-cats" style="width:280px;height:179px;"><img src="../uploads/<%=activity("image1_file_name")%>" alt="activity_image1" style="width:280px;padding-top:120px;"/></a>
  </div>
</div>
<%
end if
%>
</div><!-- end of middle-->


<div class="grid_14 prefix_1">
  <hr style="width:100%;"></hr>
</div>

<div class="grid_10 prefix_1 suffix_1">
  <div class="content">
    <p>
    上一條：<a href="activity.asp?id=previous">建設中</a>
    下一條：<a href="activity.asp?id=next">建設中</a>
    </p>
  </div>
</div>

<!--#include file="includes/footer.inc"-->
