<!--#include file="../includes/admin_header.inc"-->	

<%
dim Username,Password,Validated
Username = "ibcv"
Password = "6002"
Validated = "IBCVOK"
if Strcomp(Request.Form("User"),Username,1)=0 AND Request.Form("password") = Password then
  'Set the validation cookie and redirect the user to the original page.
  Response.Cookies("ValidUser") = Validated
  'Check where the users are coming from within the application.
  if (Request.QueryString("from")<>"") then
    Response.Redirect Request.QueryString("from")
  else
	  'If the first page that the user accessed is the Logon page,
    'direct them to the default page.
     Response.Redirect "default.asp"
  end if
else
  'Only present the failure message if the user typed in something.
  if Request.Form("User") <> "" then
    Response.Write "<div class='error'>用户名或密碼错误</div>"
  end if
end if
%>

<div class="grid_16">
  <FORM ACTION=<%Response.Write "admin/login.asp?"&Request.QueryString%> method="post">
    <h2>用户名: </h2>
    <INPUT TYPE="text" NAME="User" VALUE='' size="20"></INPUT>
    <h2>密碼: </h2>
    <INPUT TYPE="password" NAME="password" VALUE='' size="20"></INPUT>
    <INPUT TYPE="submit" VALUE="登入網站管理"></INPUT>
  </FORM>
</div>

<!--#include file="../includes/footer.inc"-->	