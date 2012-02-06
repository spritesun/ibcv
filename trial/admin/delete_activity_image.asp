<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<%
dim result
set result = execute_sql("update activities set [image1_file_name] = '' where ID_no=" & Request.QueryString("id"))
set result = Nothing

Response.Redirect "edit_activity.asp?id=" & Request.QueryString("id")

%>