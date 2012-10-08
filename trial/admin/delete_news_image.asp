<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_custom.inc"-->	

<%

dim result, id_, image_index

id_ = Request.QueryString("id")
image_index = Request.QueryString("image_index")

'remove old image
set result = execute_sql("SELECT * FROM news where ID_no = " & id_)
rm_in_uploads(result("image" & image_index & "_file_name"))

set result = execute_sql("update news set [image" & image_index & "_file_name] = '' where ID_no=" & id_)

Response.Redirect "edit_news.asp?id=" & Request.QueryString("id")

%>