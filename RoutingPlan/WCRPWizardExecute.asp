<!--#include file="..\lib\common.inc"-->
<%	
	dim ldoActionInfunction, bErrors, Item, bHaveConn, oConn, cSQL, cAHSid
	dim strAHSID, strLob_cd, strGroup_id, strState
	dim strDescription, strDestType, strEnableruleid, strEnabledflg, strInputSystemName
	dim strSequence, strRetrycount, strRetrywait_time, strDestString, strAltDestString, strTransType
	dim strseq1, strseq2, strseq3, strOdid1, strOdid2, strOdid3 
	

	bErrors = True
    strDescription      = Request.form("DESCRIPTION")
    strDestType         = Request.form("DESTINATION_TYPE")
    strEnableruleid     = Request.form("ENABLEDRULE_ID")
    
    if strEnableruleid = "" THEN
       strEnableruleid = ""  
    ELSE
       strEnableruleid = clng(strEnableruleid)
    end if
    
    strEnabledflg       = Request.form("ENABLED_FLG")
    
    if strEnabledflg = "on" then
       strEnabledflg = "Y"
    else 
       strEnabledflg = "N" 
    end if
    
    strInputSystemName  = Request.form("INPUT_SYSTEM_NAME")
    strAHSID            = Request.form("ACCNT_HRCY_STEP_ID")
    
    if not isnull(strAHSID) and not isempty(strAHSID) then 
      strAHSID = clng(strAHSID)
    end if
    
    strLob_cd           = Request.form("LOB_CD")
    strState            = Request.form("STATE")
	strSequence         = Request.form("SEQUENCE")
	strRetrycount       = Request.form("RETRY_COUNT")

	if not isempty(strSequence) and not isnull(strSequence)then 
       strSequence = cint(strSequence)
    end if
    
    if not isempty(strRetrycount) and not isnull(strRetrycount) then 
       strRetrycount = cint(strRetrycount)
    end if
   	
	strRetrywait_time    = Request.form("RETRY_WAIT_TIME")
    
	if not isempty(strRetrywait_time) and not isnull(strRetrywait_time) then 
       strRetrywait_time = cint(strRetrywait_time)
    end if
	
	strDestString        = Request.form("DESTINATION_STRING")
	strAltDestString     = Request.form("ALT_DESTINATION_STRING")
	strTransType         = Request.form("TRANSMISSION_TYPE_ID")
    
	
	if not isempty(strTransType) and not isnull(strTransType) then 
       strTransType = cint(strTransType)
    end if
	
	strseq1              = Request.form("SEQUENCE1")
	strOdid1             = Request.form("OUTPUTDEF_ID1")
	strseq2              = Request.form("SEQUENCE2")
	strOdid2             = Request.form("OUTPUTDEF_ID2")
	strseq3              = Request.form("SEQUENCE3")
	strOdid3             = Request.form("OUTPUTDEF_ID3")
	
	if strseq2 = "" then
	   strseq2 = 0
	else
	   strseq2 = cint(strseq2)
	end if
	if  strseq3 = "" then
	   strseq3 = 0
	else
	   strseq3 = cint(strseq3)
	end if
	
	if  strOdid2 = "" then
	   strOdid2 = 0
	else
	   strOdid2 = clng(strOdid2)
	end if
	if  strOdid3 = "" then
	   strOdid3 = 0
	else
	   strOdid3 = clng(strOdid3)
	end if
		
	arrState = split(strState, ",")
	'if isarray(arrState) then 
	'   Response.Write("its an array")
	'end if
	
	strDescription1 = left(strdescription,8) + " " 
	strDescription3 = " " + mid(strdescription, 9, len(strdescription)) 
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	if not (oConn)is Nothing then
		oConn.Open CONNECT_STRING
	ELSE
	   Response.Write("open conn failed")
	END IF
		
	if not (oConn)is Nothing then
	    bHaveConn = True

	    For each item in arrstate
	        strDescription  = strDescription1 +  item +   strDescription3
			cSQL	=		"{call SP_WCRPWIZARD('"	&	_
							strDescription & "','"	&	_
							strDestType	& "','"	&	_    
							strInputSystemName	& "',"		&	_
							strAHSID	& ",'"		&	_        
							strLob_cd	& "','"		&	_        
							trim(item) 	& "','"		&	_            
							strEnabledflg	& "',"		&	_    
							strEnableruleid	& ","	 &	_        
							strTransType	& ",'"	&	_        
							strDestString	& "','"		&	_    
							strAltDestString	& "',"		&	_
							strSequence	& ","		&	_        
							strRetrycount	& ","		&	_    
							strRetrywait_time	& ","		&	_
							strseq1	& ","		&	_            
							strOdid1	& ","		&	_        
							strseq2	& ","		&	_            
							strOdid2	& ","		&	_        
							strseq3	& ","		&	_            
							strOdid3	& ")}"	          
			oConn.Execute cSQL
			
			'strError = CheckADOErrors(oConn,"SP_WCRPWIZARD")
			if oConn.Errors.Count > 0 then 
			   bErrors = True
			else
			   bErrors = false
			end if
		next
		oConn.close
		set oConn = nothing
	else
	   bHaveConn = False
	End If
%>


<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="../fnsnet.css">
</HEAD>
<BODY  class="ModalDialog" onclick="onTimeExpired();">
<%
If bErrors = True Then
%>
	<p align="center" style="color:#FF0000"><strong>
	Wizard unsuccessful.
	</strong></p>
<%
elseIf bHaveConn = False Then
%>
	<p align="center" style="color:#FF0000"><strong>
	Unable to get database connection.
	</strong></p>
<%
else
	AHSID = Request.form("AHSid")
%>
	<SCRIPT LANGUAGE='VBScript'>
	self.location.href = "..\AH\NodeSummary.asp?AHSID=<%=AHSID%>"
	</SCRIPT>
<%
End If 
%>

<SCRIPT LANGUAGE='JScript'>


function onTimeExpired()
{
<%	If bErrors = True Then %>
	window.returnValue = false;
<%	Else  %>
	window.returnValue = true;
<%	End If %>
	window.close();
}


function onLoad()
{
<%	If bErrors = False Then %>
	setTimeout("onTimeExpired()", 3*1000);
<%	Else  %>
	setTimeout("onTimeExpired()", 10*1000);
<%	End If %>
}

</SCRIPT>
</BODY>
</HTML>


