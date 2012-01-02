<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_fmt.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<%
dim result, strSQL

strSQL = fmt("update activities set [name] = '%x', [date] = '%x', [time] = '%x', [venue] = '%x', [content] = '%x' where ID_no = %x;", Array(Request.Form("name"), Request.Form("date"), Request.Form("time"), Request.Form("venue"), Request.Form("content"), Request.Form("id")))

set result = execute_sql(strSQL)
set result = Nothing

Response.Redirect "activities.asp"
%>
