<%
function execute_sql(strSQL)
  dim adoCon

  set adoCon = Server.CreateObject("ADODB.Connection")
  adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/_private/ibcv.mdb")

  set execute_sql = adoCon.Execute(strSQL)
  set adoCon = Nothing
end function
%>
