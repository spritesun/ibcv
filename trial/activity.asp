<!--#include file="includes/header.inc"-->
<!--#include file="includes/function_execute_sql.inc"-->	
<%
dim activity
set activity = execute_sql("select * from activities where ID_no=" & Request.QueryString("id"))
%>

<div class="grid_10 prefix_1">
  <div class="content">
    <h1><%=activity("name")%></h1>
    <p>時間: <%=activity("date")%>， <%=activity("time")%></p>
    <p>地點: <%=activity("venue")%></p>
    <br />
    <p><%=activity("content")%></p>
  </div>
</div>

</div><!-- end of middle-->

<div class="bottomLine prefix_1">
  <hr style="widtd:880px;"></hr>
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
