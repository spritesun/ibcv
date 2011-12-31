<%
Dim adoCon, result, strSQL

Set adoCon = Server.CreateObject("ADODB.Connection")
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../_private/ibcv.mdb")

strSQL = "INSERT INTO activities ([name], [date], [time], [venue], [content]) VALUES ('" & Request.Form("name") & "', '" &  Request.Form("date") & "', '" & Request.Form("time") & "', '" & Request.Form("venue") & "', '" &  Request.Form("content") & "');"

Set result = adoCon.Execute(strSQL)

Set rsAddComments = Nothing
Set adoCon = Nothing

Response.Redirect "activities.asp"
%>