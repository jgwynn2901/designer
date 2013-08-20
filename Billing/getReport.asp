<%
Response.Buffer = False
Response.Expires = -1


dim cFileID, cFilePath, oExcel

cFileID = Request.QueryString("FILEID")
cFilePath = Request.QueryString("FILEPATH")
Set oExcel = Server.CreateObject("ExcelClass.XLSClass")
oExcel.cDownloadLocation = cFilePath
oExcel.cDestinationFileName = cFileID
oExcel.cBackground = "#d6cfbd"
oExcel.sendFile
Set oExcel = Nothing
'
%>
