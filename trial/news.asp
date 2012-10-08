<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_execute_sql.inc"-->	
<%
'declare
dim news, text_grid_size

function hasImage
  hasImage = (news("image1_file_name") <> "" or news("image2_file_name") <> "" or news("image3_file_name") <> "")
end function


set news = execute_sql("select * from news where ID_no=" & Request.QueryString("id"))

if hasImage() then
  text_grid_size = "grid_9"
else
  text_grid_size = "grid_14"
end if
%>

<div class="<%=text_grid_size%> prefix_1">
  <div class="content">
    <h1><%=news("name")%></h1>
    <p>時間: <%=news("date")%></p>
    <br />
    <p><%=Replace(news("content"), vbCr, "<br>")%></p>
    
    <%
    if news("attach1_file_name") <> "" then
    %>
    <p>
      <span>此處下載:</span>
      <a href="../uploads/<%=news("attach1_file_name")%>" class="file"><%=news("attach1_origin_name")%></a>
    </p>
    <%
    end if
    %>
    
  </div>
</div>

<%
if hasImage() then
%>
<div class="grid_5">
  <div class="content">
<%
for i = 1 to 3
  if news("image" & i & "_file_name") <> "" then
%>
    <a href="../uploads/<%=news("image" & i & "_file_name")%>" rel="lightbox-cats" class="news_image"><img src="../uploads/<%=news("image" & i & "_file_name")%>" alt="news_image" & i/></a>
<%
  end if
next
%>
  </div>
</div>
<%
end if
%>
</div><!-- end of middle-->

<!--#include file="includes/footer.inc"-->
