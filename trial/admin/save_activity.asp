<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_custom.inc"-->	
<!--#include file="../includes/freeASPUpload.asp" -->

<%
'declare
dim uploadsDirVar, result, strSQL, name, date_, time_, venue, content, id_, Upload, image_file_extension, splitted_image_file_name, image_dest_name, splitted_attach_file_name, attach_dest_name, attach_origin_name

sub update_image
  if Upload.UploadedFiles.Exists("image1") then
    'remove old image
    strSQL = "SELECT * FROM activities where ID_no = " & id_
    set result = execute_sql(strSQL)
    rm_in_uploads(result("image1_file_name"))
    
    splitted_image_file_name = Split(Upload.UploadedFiles("image1").FileName, ".")
    image_file_extension = splitted_image_file_name(UBound(splitted_image_file_name))
    image_dest_name = "activity_" & id_ & "_image1." & image_file_extension
  
    Upload.UploadedFiles("image1").FileName = image_dest_name

    strSQL = fmt("update activities set [image1_file_name] = '%x' where ID_no = %x", Array(image_dest_name, id_))
    set result = execute_sql(strSQL)
  end if
end sub

sub update_attach
  if Upload.UploadedFiles.Exists("attach1") then
    'remove old file
    strSQL = "SELECT * FROM activities where ID_no = " & id_
    set result = execute_sql(strSQL)
    rm_in_uploads(result("attach1_file_name"))

    'save attach
    attach_origin_name = Upload.UploadedFiles("attach1").FileName
    splitted_attach_file_name = Split(Upload.UploadedFiles("attach1").FileName, ".")
    attach_file_extension = splitted_attach_file_name(UBound(splitted_attach_file_name))
    attach_dest_name = "activity_" & id_ & "_attach1." & attach_file_extension

    Upload.UploadedFiles("attach1").FileName = attach_dest_name
  
    strSQL = fmt("update activities set [attach1_file_name] = '%x', [attach1_origin_name] = '%x' where ID_no = %x", Array(attach_dest_name, attach_origin_name, id_))
    set result = execute_sql(strSQL)
  end if
end sub

function get_max_id
  strSQL = "SELECT max(ID_no) as max_id FROM activities;"
  set result = execute_sql(strSQL)
  get_max_id = result("max_id")
end function

'init
uploadsDirVar = Server.MapPath("../../uploads")
Set Upload = New FreeASPUpload
Upload.Upload
name = Upload.Form("name")
date_ = Upload.Form("date")
time_ = Upload.Form("time")
venue = Upload.Form("venue")
content = Upload.Form("content")
id_ = Upload.Form("id")

if id_="" then
  strSQL = "INSERT INTO activities ([name], [date], [time], [venue], [content]) VALUES ('" & name & "', '" &  date_ & "', '" & time_ & "', '" & venue & "', '" &  content & "');"
  set result = execute_sql(strSQL)
  
  if Upload.UploadedFiles.Exists("image1") or Upload.UploadedFiles.Exists("attach1") then
    'do not know why need twice to get the real max id, but it works anyway. Long. Feb 13, 2012.
    id_ = get_max_id()
    id_ = get_max_id()
    update_image()
    update_attach()
    Upload.Save(uploadsDirVar)
  end if
else
  'editing
  strSQL = fmt("UPDATE activities SET [name] = '%x', [date] = '%x', [time] = '%x', [venue] = '%x', [content] = '%x' where ID_no = %x;", Array(name, date_, time_, venue, content, id_))
  set result = execute_sql(strSQL)
    
  update_image()
  update_attach()
  Upload.Save(uploadsDirVar)
end if

set result = Nothing
Response.Redirect "activities.asp"
%>
