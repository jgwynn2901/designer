<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->

<HTML>
<HEAD>
<SCRIPT>
<%
   On Error Resume Next
    ACTION = CStr(Request.Form("TxtAction"))
	SQL_STRING = Request.Form("TxtSaveData")
	
	'  CoverageCode AND VendorDesignator HAS TO BE UNIC 
	    DIM Coverage ,rsRecordCount,RS,VendorDesignator, SQLST2,rsHolder
            Coverage   = Request.Form("txtCoverageCode")
            VendorDesignator =Request.Form("txtVendorDesignator")
      if   Coverage <> "" and  VendorDesignator <> "" then 
          Set Conn = Server.CreateObject("ADODB.Connection")
		    Conn.Open CONNECT_STRING
		    SQLST2 = ""
		
		   SQLST2 = "SELECT COVERAGECODE_CONVERSION_ID,ACCNT_HRCY_STEP_ID,"
		   SQLST2 = SQLST2 & "COVERAGE_CODE,VENDOR_DESIGNATOR,DESCRIPTION "
		   SQLST2 = SQLST2 & "FROM COVERAGECODE_CONVERSION "
		   SQLST2 = SQLST2 & "WHERE COVERAGE_CODE = '" & Coverage & "'"
		   SQLST2 = SQLST2 & " AND VENDOR_DESIGNATOR='" & VendorDesignator & "'"
		   Set RS = Conn.Execute(SQLST2)
             if not (RS.EOF OR RS.BOF) then
                rsRecordCount= RS("COVERAGECODE_CONVERSION_ID")
              else
              rsRecordCount=0
            end if
         end if
    rsHolder=rsRecordCount  
	
	If ACTION = "UPDATE" Then
	   UpdateSQL = ""
		UpdateSQL = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "COVERAGECODE_CONVERSION", "COVERAGECODE_CONVERSION_ID", "")		 
		Set RSinsert = Conn.Execute(UpdateSQL)
		strError = CheckADOErrors(Conn,"COVERAGECODE_CONVERSION " & ACTION)
	Elseif ACTION = "INSERT" AND  rsHolde = 0 Then
	    InsertSQL = ""
	     NewXREFID = NextPkey("COVERAGECODE_CONVERSION","COVERAGECODE_CONVERSION_ID")
	      If NewXREFID > 0  Then
			InsertSQL = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "COVERAGECODE_CONVERSION", "COVERAGECODE_CONVERSION_ID", NewXREFID)		 
			Set RSUpdate = Conn.Execute(InsertSQL)
			
			strError = CheckADOErrors(Conn,"COVERAGECODE_CONVERSION " & ACTION)
			If strError = "" Then Response.write("parent.frames('WORKAREA').UpdateXREFID ('" & NewXREFID &  "');")	
		Else
			strError = "Unable to obtain next primary key for COVERAGECODE_CONVERSION table."
		End If			
End If
	
	
	Conn.Close
	
	If strError <> "" Then
		LogStatusGroupBegin
		LogStatus S_ERROR, strError, "COVERAGECODE_CONVERSION", "", 0, ""
		LogStatusGroupEnd
		%>
		parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
		parent.frames('WORKAREA').SetDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);
<%	Else
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('WORKAREA').UpdateStatus('Update successful.');
		parent.frames('WORKAREA').ClearDirty();
		parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);		
<%	End If  %>
</SCRIPT>
</HEAD>
<BODY>
</BODY>
