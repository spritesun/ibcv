<!--#include file="../includes/login_filter.inc"-->	
<%
Dim adoCon, rsGuestbook, strSQL

Set adoCon = Server.CreateObject("ADODB.Connection")
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../_private/ibcv.mdb")

strSQL = "alter table activities add column category TEXT(50);"

'strSQL = "CREATE TABLE news ([ID_no] AUTOINCREMENT, [name] TEXT(255), [date] TEXT(255), [summary] MEMO, [content] MEMO, image1_file_name TEXT(255), image2_file_name TEXT(255), image3_file_name TEXT(255), attach1_file_name TEXT(255), attach1_origin_name TEXT(255));"

'strSQL = "alter table activities add column XXXX"
'strSQL = "INSERT INTO activities ([name], [date], [time], [venue], [content]) VALUES ('test_name', 'test_date', 'test_time', 'test_venue', 'test_content');" 
'drop and create activities table
'strSQL = "DROP TABLE activities;"
'strSQL = "CREATE TABLE activities ([ID_no] AUTOINCREMENT, [name] TEXT(255), [date] TEXT(255), [time] TEXT(255), [venue] TEXT(255), [summary] MEMO, [content] MEMO, image1_file_name TEXT(50), attach1_file_name TEXT(50), attach1_origin_name TEXT(50));"

Set rsGuestbook = adoCon.Execute(strSQL)

'For Each f In rsGuestbook.Fields
'    Response.Write (f.name)
'Next

'Response.Write(rsGuestbook.GetString)
'Response.Write "success"

'Do While not rsGuestbook.EOF
' 
' 'Write the HTML to display the current record in the recordset
' Response.Write ("<br>")
' Response.Write (rsGuestbook("Name"))
' Response.Write ("<br>")
' Response.Write (rsGuestbook("Comments"))
' Response.Write ("<br>")
'
' 'Move to the next record in the recordset
' rsGuestbook.MoveNext
'
'Loop

'rsGuestbook.Close
Set rsGuestbook = Nothing
Set adoCon = Nothing

%>