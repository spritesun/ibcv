<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<%
dim result
set result = execute_sql("delete from activities where ID_no=" & Request.QueryString("id"))
set result = Nothing

Response.Redirect "activities.asp"
%>