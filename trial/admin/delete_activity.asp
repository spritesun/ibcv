<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_custom.inc"-->	
<%
dim result, id_

id_ = Request.QueryString("id")

'remove old image & attach
set result = execute_sql("SELECT * FROM activities where ID_no = " & id_)
rm_in_uploads(result("image1_file_name"))
rm_in_uploads(result("attach1_file_name"))

set result = execute_sql("delete from activities where ID_no=" & id_)
set result = Nothing

Response.Redirect "activities.asp"
%>