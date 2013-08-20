<!--#include file="..\lib\genericSQL.asp"-->
<%

	SQL_UPDATE_DIVIDE = Split(Request.Form("TxtUpdateData") , chr(130))
	SQL_INSERT_DIVIDE = Split(Request.Form("TxtInsertData") , chr(130))
	SQL_DELETE_DIVIDE = Split(Request.Form("TxtDeleteData") , chr(130))

	For J = 0 to CLng(Request.Form("UPCOUNT"))-1 Step 1
		UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_UPDATE_DIVIDE(J), Chr(128), Chr(129), "UPDATE", "ATTR_INSTANCE", "ATTR_INSTANCE_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
	Next
	
	For a = 0 to CLng(Request.Form("INCOUNT"))-1 Step 1
		InsertSQL = ""
		InsertSQL = BuildSQL(SQL_INSERT_DIVIDE(a), Chr(128), Chr(129), "INSERT", "ATTR_INSTANCE", "ATTR_INSTANCE_ID", NextPkey("ATTR_INSTANCE", "ATTR_INSTANCE_ID"))		 
		Set RSUpdate = Conn.Execute(InsertSQL)
	Next
	
	For x = 0 to CLng(Request.Form("DELCOUNT"))-1 Step 1
		DeleteSQL = ""
		DeleteSQL = BuildSQL("", "", "", "DELETE", "ATTR_INSTANCE", "ATTR_INSTANCE_ID", SQL_DELETE_DIVIDE(x))		 
		Set RSUpdate = Conn.Execute(DeleteSQL)
	Next
	
 %>
<HTML>
<HEAD>

</HEAD>
<BODY>

</BODY>
