Function FindLabel ( [PARCELID] )
Dim strPrclQry, strInfo, strInfo2
strPrclQry = "SELECT VW_ZONINGHISTORYPARCELS.PARCELID, VW_ZONINGHISTORYPARCELS.CASENUMCOMBINED FROM LANDUSE.VW_ZONINGHISTORYPARCELS WHERE PARCELID = '" & [PARCELID] & "'"
Dim ADOConn
set ADOConn = createobject("ADODB.Connection")
Dim rsPrcl
set rsPrcl = createObject("ADODB.Recordset")
ADOConn.Open "Provider=msdaora;User ID=<userID>;Password=<password>;Persist Security Info=true;Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = <host>)(PORT = <port>)))(CONNECT_DATA =(SERVICE_NAME = <service_name>)))"
ADOConn.CursorLocation = 3
rsPrcl.Open strPrclQry, ADOConn, 3, 1, 1
Select Case rsPrcl.RecordCount
Case -1, 0
strInfo = ""
strInfo2 = ""
strLabel = strInfo & strInfo2
Case Else
strLabel = ""
For i = 0 To rsPrcl.RecordCount -1
strInfo2 = rsPrcl.Fields("CASENUMCOMBINED").Value
strLabel = strLabel &  strInfo2 & chr(13)
rsPrcl.MoveNext
Next
End Select
rsPrcl.Close
ADOConn.Close
Set rsPrcl = Nothing
Set ADOConn = Nothing
FindLabel = strLabel
End Function