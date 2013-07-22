<!--#include file="../includes/login_filter.inc"-->	
<!--#include file="../includes/function_custom.inc"-->	
<!--#include file="../includes/freeASPUpload.asp" -->

<%
'declare
dim uploadsDirVar, result, strSQL, name, date_, content, id_, Upload, image_file_extension, splitted_image_file_name, image_dest_name, splitted_attach_file_name, attach_dest_name, attach_origin_name

'could be more efficiency if update images in one sql transaction
sub update_image(index)
  if Upload.UploadedFiles.Exists("image" & index) then
    'remove old image
    strSQL = "SELECT * FROM news where ID_no = " & id_
    set result = execute_sql(strSQL)
    rm_in_uploads(result("image" &  index & "_file_name"))
    
    splitted_image_file_name = Split(Upload.UploadedFiles("image" & index).FileName, ".")
    image_file_extension = splitted_image_file_name(UBound(splitted_image_file_name))
    image_dest_name = "news_" & id_ & "_image" &  index & "." & image_file_extension
  
    Upload.UploadedFiles("image" & index).FileName = image_dest_name

    strSQL = fmt("update news set [image%x_file_name] = '%x' where ID_no = %x", Array(index, image_dest_name, id_))
    set result = execute_sql(strSQL)
  end if
end sub

sub update_attach
  if Upload.UploadedFiles.Exists("attach1") then
    'remove old file
    strSQL = "SELECT * FROM news where ID_no = " & id_
    set result = execute_sql(strSQL)
    rm_in_uploads(result("attach1_file_name"))

    'save attach
    attach_origin_name = Upload.UploadedFiles("attach1").FileName
    splitted_attach_file_name = Split(Upload.UploadedFiles("attach1").FileName, ".")
    attach_file_extension = splitted_attach_file_name(UBound(splitted_attach_file_name))
    attach_dest_name = "news_" & id_ & "_attach1." & attach_file_extension

    Upload.UploadedFiles("attach1").FileName = attach_dest_name
  
    strSQL = fmt("update news set [attach1_file_name] = '%x', [attach1_origin_name] = '%x' where ID_no = %x", Array(attach_dest_name, attach_origin_name, id_))
    set result = execute_sql(strSQL)
  end if
end sub

function get_max_id
  strSQL = "SELECT max(ID_no) as max_id FROM news;"
  set result = execute_sql(strSQL)
  get_max_id = result("max_id")
end function

sub save_all_attachments
  update_image("1")
  update_image("2")
  update_image("3")
  update_attach()
  Upload.Save(uploadsDirVar)
end sub

'init
uploadsDirVar = Server.MapPath("../../uploads")
Set Upload = New FreeASPUpload
Upload.Upload
name = Upload.Form("name")
date_ = Upload.Form("date")
summary = Upload.Form("summary")
content = Upload.Form("content")
content = replace(content, "'", "''")
id_ = Upload.Form("id")

if id_="" then
  strSQL = "INSERT INTO news ([name], [date], [summary], [content]) VALUES ('" & name & "', '" &  date_ & "', '" & summary & "', '" &  content & "');"
  set result = execute_sql(strSQL)
  
  if Upload.UploadedFiles.Exists("image1") or Upload.UploadedFiles.Exists("image2") or Upload.UploadedFiles.Exists("image3") or Upload.UploadedFiles.Exists("attach1") then
    'do not know why need twice to get the real max id, but it works anyway. Long. Feb 13, 2012.
    id_ = get_max_id()
    id_ = get_max_id()
    save_all_attachments()
  end if
else
  'editing
  strSQL = fmt("UPDATE news SET [name] = '%x', [date] = '%x', [summary]='%x', [content] = '%x' where ID_no = %x;", Array(name, date_, summary, content, id_))
  set result = execute_sql(strSQL)
    
  save_all_attachments()
end if

set result = Nothing
Response.Redirect "news_list.asp"
%>
