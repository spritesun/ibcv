<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_custom.inc"-->	

<%

dim result, id_

id_ = Request.QueryString("id")

'remove old image
set result = execute_sql("SELECT * FROM activities where ID_no = " & id_)
rm_in_uploads(result("image1_file_name"))

set result = execute_sql("update activities set [image1_file_name] = '' where ID_no=" & id_)

Response.Redirect "edit_activity.asp?id=" & Request.QueryString("id")

%>