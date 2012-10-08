<!--#include file="function_fmt.inc"-->	
<!--#include file="function_execute_sql.inc"-->	

<%
sub rm(file_name)
  dim file_sys
  set file_sys = CreateObject("Scripting.FileSystemObject")
  if file_sys.FileExists(file_name) then
    file_sys.DeleteFile(file_name)
  end if
end sub

sub rm_in_uploads(file_name)
  file_name = Server.MapPath("../../uploads") & "\" & file_name
  rm(file_name)
end sub

sub round_bar(title)
%>
<div class="round-bar">
  <div class="left"><img src="img/nav-left.png"></div>
  <div class="right"><img src="img/nav-right.png"></div>
  <span><%=title%></span>
</div>
<%
end sub
%>


