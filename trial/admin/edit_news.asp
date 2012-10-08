<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/admin_header.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	

<div class="grid_14 prefix_1 suffix_1">
<%
'declare
dim id_, news, name, date_, content, image1_file_name, image2_file_name, image3_file_name, attach1_file_name, attach1_origin_name

sub render_image_section(image_file_name, index)
  dim rowHTMLStr, params
  if image_file_name <> "" then
    rowHTMLStr = "<img src='../uploads/%x' alt='news_image%x'></img>"
    
    rowHTMLStr = rowHTMLStr & "<a href='./admin/delete_news_image.asp?id=%x&image_index=%x' class='delete'>删除</a>"
    rowHTMLStr = fmt(rowHTMLStr, Array(image_file_name, index, id_, index))
  end if
  rowHTMLStr = rowHTMLStr & "<br/><span>上傳新的圖片%x</span>"
  rowHTMLStr = rowHTMLStr & "<input type='file' name='image%x' value=''/><br/><hr>"
  response.write(fmt(rowHTMLStr, Array(index, index)))
end sub

'init
id_ = Request.QueryString("id")
if (id_<>"") then
  set news = execute_sql("select * from news where ID_no=" & id_)
  name = news("name")
  date_ = news("date")
  summary = news("summary")
  content = news("content")
  image1_file_name = news("image1_file_name")
  image2_file_name = news("image2_file_name")
  image3_file_name = news("image3_file_name")
  attach1_file_name = news("attach1_file_name")
  attach1_origin_name = news("attach1_origin_name")
end if
%>

  <form action="./admin/save_news.asp" method="post" enctype="multipart/form-data"  accept-charset="utf-8">
    <input type="hidden" name="id" value="<%=id_%>"/>
    <h2>新聞名稱</h2>
    <input type="text" name="name" value="<%=name%>"/>
    <h2>日期</h2>
    <input type="text" name="date" value="<%=date_%>"/>
    <h2>摘要（將會顯示在首頁）</h2>
    <textarea rows="10" cols="60" name="summary"><%=summary%></textarea>
    <h2>内容介紹</h2>
    <textarea rows="20" cols="60" name="content"><%=content%></textarea>

    <div class="upload_management_sect">
      <h2>管理新聞圖片</h2>
      <span class="notice">(如果圖片没有立即更新，請嘗試刷新頁面再查看)</span>
      <hr>
      <% render_image_section image1_file_name, "1" %>
      <% render_image_section image2_file_name, "2" %>
      <% render_image_section image3_file_name, "3" %>
    </div>

    <div class="upload_management_sect">
      <h2>管理新聞附件</h2>

<%
if attach1_file_name <> "" then
%>
      <span>下載:</span><a href="../uploads/<%=attach1_file_name%>" class="file"><%=attach1_origin_name%></a>
      <a href="./admin/delete_news_attach.asp?id=<%=id_%>" class="delete">删除</a>
<%
end if
%>
      <br/><span>上傳新的附件</span>
      <input type="file" name="attach1" value=""/>
    </div>

    <input type="submit" value="提交新聞"/>
  </form>

</div>

<%
set news = Nothing
%>
<!--#include file="../includes/footer.inc"-->