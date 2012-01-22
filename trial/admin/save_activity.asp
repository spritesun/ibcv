<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_fmt.inc"-->	
<!--#include file="../includes/function_execute_sql.inc"-->	
<!--#include file="../includes/freeASPUpload.asp" -->

<%
dim uploadsDirVar
uploadsDirVar = Server.MapPath("../../uploads")

dim result, strSQL, name, date_, time_, venue, content, id_, Upload

Set Upload = New FreeASPUpload

Upload.Upload

name = Upload.Form("name")
date_ = Upload.Form("date")
time_ = Upload.Form("time")
venue = Upload.Form("venue")
content = Upload.Form("content")
id_ = Upload.Form("id")

dim image_file_extension, splitted_image_file_name, image_dest_name

if id_="" then
  strSQL = "INSERT INTO activities ([name], [date], [time], [venue], [content]) VALUES ('" & name & "', '" &  date_ & "', '" & time_ & "', '" & venue & "', '" &  content & "');"
  set result = execute_sql(strSQL)
  
  if Upload.UploadedFiles.Exists("image1") then
    strSQL = "SELECT max(ID_no) as max_id FROM activities;"
    set result = execute_sql(strSQL)
    id_ = result("max_id")
  
    strSQL = "SELECT max(ID_no) as max_id FROM activities;"
    set result = execute_sql(strSQL)
    id_ = result("max_id")
  
    splitted_image_file_name = Split(Upload.UploadedFiles("image1").FileName, ".")
    image_file_extension = splitted_image_file_name(UBound(splitted_image_file_name))
    image_dest_name = "activity_" & id_ & "_image1." & image_file_extension
    
    Upload.UploadedFiles("image1").FileName = image_dest_name
    Upload.Save(uploadsDirVar)
    
    strSQL = fmt("update activities set [image1_file_name] = '%x' where ID_no = %x", Array(image_dest_name, id_))
    set result = execute_sql(strSQL)
  end if
else
  strSQL = fmt("UPDATE activities SET [name] = '%x', [date] = '%x', [time] = '%x', [venue] = '%x', [content] = '%x' where ID_no = %x;", Array(name, date_, time_, venue, content, id_))
  set result = execute_sql(strSQL)
  
  if Upload.UploadedFiles.Exists("image1") then
    splitted_image_file_name = Split(Upload.UploadedFiles("image1").FileName, ".")
    image_file_extension = splitted_image_file_name(UBound(splitted_image_file_name))
    image_dest_name = "activity_" & id_ & "_image1." & image_file_extension
  
    Upload.UploadedFiles("image1").FileName = image_dest_name
    Upload.Save(uploadsDirVar)
    
    strSQL = fmt("update activities set [image1_file_name] = '%x' where ID_no = %x", Array(image_dest_name, id_))
    set result = execute_sql(strSQL)
  end if
end if

set result = Nothing
Response.Redirect "activities.asp"
%>
