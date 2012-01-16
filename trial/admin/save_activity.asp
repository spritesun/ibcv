<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_fmt.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<!--#include file="../includes/freeASPUpload.asp" -->

<%
dim uploadsDirVar
uploadsDirVar = Server.MapPath("../../uploads")

dim result, strSQL, name, date_, time_, venue, content, id_, Upload

Set Upload = New FreeASPUpload
Upload.Save(uploadsDirVar)

name = Upload.Form("name")
date_ = Upload.Form("date")
time_ = Upload.Form("time")
venue = Upload.Form("venue")
content = Upload.Form("content")
id_ = Upload.Form("id")

if id_="" then
  strSQL = "INSERT INTO activities ([name], [date], [time], [venue], [content]) VALUES ('" & name & "', '" &  date_ & "', '" & time_ & "', '" & venue & "', '" &  content & "');"
else
  strSQL = fmt("UPDATE activities SET [name] = '%x', [date] = '%x', [time] = '%x', [venue] = '%x', [content] = '%x' where ID_no = %x;", Array(name, date_, time_, venue, content, id_))
end if

set result = execute_sql(strSQL)
set result = Nothing
Response.Redirect "activities.asp"
%>
