<!--#include file="../includes/function_execute_sql.inc"-->	
<%
dim result, strSQL

strSQL = "INSERT INTO activities ([name], [date], [time], [venue], [content]) VALUES ('" & Request.Form("name") & "', '" &  Request.Form("date") & "', '" & Request.Form("time") & "', '" & Request.Form("venue") & "', '" &  Request.Form("content") & "');"

set result = execute_sql(strSQL)

set rsAddComments = Nothing
set adoCon = Nothing

Response.Redirect "activities.asp"
%>