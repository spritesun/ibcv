<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_custom.inc"-->	

<%

dim result, id_

id_ = Request.QueryString("id")

'remove old attach
set result = execute_sql("SELECT * FROM news where ID_no = " & id_)
rm_in_uploads(result("attach1_file_name"))

set result = execute_sql("update news set [attach1_file_name] = '' where ID_no=" & id_)

Response.Redirect "edit_news.asp?id=" & Request.QueryString("id")

%>