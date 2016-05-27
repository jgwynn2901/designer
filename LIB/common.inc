<%

Const BODYBGCOLOR = "#d6cfbd"
Const adUseClient = 3
Const adParamInput = &H0001
Const adInteger = 3
Const adOpenForwardOnly = 0
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adStateOpen = &H00000001

Dim CONNECT_STRING
CONNECT_STRING=Session("ConnectionString")

Function renderCell( cValue )
	if IsNull(cValue) or Len(Trim(cValue)) <= 0 then
		renderCell = "&nbsp;"
	else
		renderCell = cValue
	end if
End Function

Sub Getdefault
	'****************************************
    ' DMS: 2/25/00 
	' Instead of hard coding the values, they can be grabbed from the database
	' The 3 entries will need to be added to the database.
	'****************************************
	set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	s_SQL2 = "SELECT name, value " &_
	         "  FROM setting      " &_
	         " WHERE name in ('QAFAX', 'QAPRINTER' , 'PRODUCTIONPRINTER','QAEMAIL', 'QALEGACYEMAIL') "
	
	Conn.Open ConnectionString
	set RS2 = Conn.Execute(s_SQL2)
	do while not RS2.EOF
	    select case RS2("NAME")
		    case "QAFAX":
	           s_QAFax             = RS2("VALUE")		    
			case "QAPRINTER" :
			   s_QAPrinter         = RS2("VALUE")		    
			case "PRODUCTIONPRINTER":
			   s_ProductionPrinter = RS2("VALUE")		
			case "QAEMAIL":
				s_QAEMAIL = RS2("VALUE")    
			case "QALEGACYEMAIL":
				s_QALEGACYEMAIL		= RS2("VALUE")
		end select
	    RS2.movenext  
	loop

	RS2.Close
	Conn.Close
	SET RS2  = NOTHING
	SET Conn = NOTHING
	
	If (IsNull(s_QAFAX) or IsEmpty(s_QAFAX)) then s_QAFAX             = "5017475675"
    If (IsNull(s_QAPRINTER) or IsEmpty(s_QAPRINTER)) then s_QAPRINTER = "\\Cha0s2t\CHA4SI"
	If (IsNull(s_PRODUCTIONPRINTER) or IsEmpty(s_PRODUCTIONPRINTER)) then 
	   s_PRODUCTIONPRINTER = "\\Cha0s00t\OPER_HP5SI_A"
	End If
	If (IsNull(s_QAEMAIL) or IsEmpty(s_QAEMAIL)) then s_QAEMAIL = "FNSINET@CONCENTRAMC.COM"
	If (IsNull(s_QALEGACYEMAIL) or IsEmpty(s_QALEGACYEMAIL)) then s_QALEGACYEMAIL = "FNSINET@CONCENTRAMC.COM"
	'response.write s_QAFax,s_QAPrinter, s_ProductionPrinter,s_qaemail 
end sub

function getInstanceName
dim cSQL, oConn, oRS

set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
cSQL = "Select sys_context('USERENV','DB_NAME') as INSTANCE from dual"	
set oRS = oConn.Execute(cSQL)
getInstanceName = oRS.fields("INSTANCE").value
oRS.Close
oConn.Close
set oRS  = nothing
set oConn = nothing
end function
%>
<!--#include file="..\lib\logincheck.asp"-->