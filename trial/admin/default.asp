<!--#include file="../includes/admin_header.inc"-->	

<%
dim Validated
Validated = "OK"
if Request.Cookies("ValidUser") <> Validated then
  'Construct the URL for the current page.
  dim s
  s = "http://"
  s = s & Request.ServerVariables("HTTP_HOST")
  s = s & Request.ServerVariables("URL")
  if Request.QueryString.Count > 0 then
	  s = s & "?" & Request.QueryString 
  end if
  'Redirect unauthorized users to the logon page.
  Response.Redirect "login.asp?from=" &Server.URLEncode(s)
end if
%>

<div class="grid_4">Menu Panel</div>
<div class="grid_12">添加新的活動</div>

<!--#include file="../includes/footer.inc"-->	
