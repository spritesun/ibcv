<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	

<div class="grid_14 prefix_1 suffix_1">
<%
dim id_, activity, name, date_, time_, venue, content, image1_file_name, attach1_file_name, attach1_origin_name
id_ = Request.QueryString("id")
if (id_<>"") then
  set activity = execute_sql("select * from activities where ID_no=" & id_)
  name = activity("name")
  date_ = activity("date")
  time_ = activity("time")
  venue = activity("venue")
  summary = activity("summary")
  content = activity("content")
  image1_file_name = activity("image1_file_name")
  attach1_file_name = activity("attach1_file_name")
  attach1_origin_name = activity("attach1_origin_name")
end if
%>

  <form action="./admin/save_activity.asp" method="post" enctype="multipart/form-data"  accept-charset="utf-8">
    <input type="hidden" name="id" value="<%=id_%>"/>
    <h2>通告名稱</h2>
    <input type="text" name="name" value="<%=name%>"/>
    <h2>日期</h2>
    <input type="text" name="date" value="<%=date_%>"/>
    <h2>時間</h2>
    <input type="text" name="time" value="<%=time_%>"/>
    <h2>地點</h2>
    <input type="text" name="venue" value="<%=venue%>"/>
    <h2>摘要（將會顯示在首頁）</h2>
    <textarea rows="10" cols="60" name="summary"><%=summary%></textarea>
    <h2>内容介紹</h2>
    <textarea rows="20" cols="60" name="content"><%=content%></textarea>

    <div class="upload_management_sect">
      <h2>管理通告圖片</h2>

<%
if image1_file_name <> "" then
%>
      <img src="../uploads/<%=image1_file_name%>" alt="activity_image1"></img>
      <a href="./admin/delete_activity_image.asp?id=<%=id_%>" class="delete">删除</a>
<%
end if
%>
      <br/><span>上傳新的圖片</span>
      <input type="file" name="image1" value=""/>
    </div>

    <div class="upload_management_sect">
      <h2>管理通告附件</h2>

<%
if attach1_file_name <> "" then
%>
      <span>下載:</span><a href="../uploads/<%=attach1_file_name%>" class="file"><%=attach1_origin_name%></a>
      <a href="./admin/delete_activity_attach.asp?id=<%=id_%>" class="delete">删除</a>
<%
end if
%>
      <br/><span>上傳新的附件</span>
      <input type="file" name="attach1" value=""/>
    </div>

    <input type="submit" value="提交通告"/>
  </form>

</div>

<%
set activity = Nothing
%>
<!--#include file="../includes/footer.inc"-->