<%
dim Validated
Validated = "IBCVOK"
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
