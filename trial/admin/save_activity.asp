<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_fmt.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<%
dim result, strSQL

if Request.Form("id")="" then
  strSQL = "INSERT INTO activities ([name], [date], [time], [venue], [content]) VALUES ('" & Request.Form("name") & "', '" &  Request.Form("date") & "', '" & Request.Form("time") & "', '" & Request.Form("venue") & "', '" &  Request.Form("content") & "');"
else
  strSQL = fmt("UPDATE activities SET [name] = '%x', [date] = '%x', [time] = '%x', [venue] = '%x', [content] = '%x' where ID_no = %x;", Array(Request.Form("name"), Request.Form("date"), Request.Form("time"), Request.Form("venue"), Request.Form("content"), Request.Form("id")))
end if

set result = execute_sql(strSQL)
set result = Nothing

Response.Redirect "activities.asp"
%>
