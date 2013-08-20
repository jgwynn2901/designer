<%@ LANGUAGE="VBSCRIPT" %>
<%  Option Explicit
	on error resume next
	Response.expires=0
	
	dim strList ' Global
	If Session("CurrentUser").CanUserDo("Account Maintenance","MODIFY") <> true  Then 
		Response.redirect "../AccessDenied.asp"
	End If
	

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../fnsnet.css">
<TITLE>Call Flow Cache list</TITLE>
</head>
<script language="VBScript">

sub onClickSubmit()
   
   
   frmCache.CallFlowId.value = frmCache.inpEntry.options(frmCache.inpEntry.selectedIndex).CALLFLOWID
   frmCache.CallStFlg.value  = frmCache.inpEntry.options(frmCache.inpEntry.selectedIndex).CALLSTFLG
   frmCache.Client.value     = frmCache.inpEntry.options(frmCache.inpEntry.selectedIndex).CLIENT
   frmCache.LOB.value        = frmCache.inpEntry.options(frmCache.inpEntry.selectedIndex).LOB
   frmCache.DType.value      = frmCache.inpEntry.options(frmCache.inpEntry.selectedIndex).DTYPE
   frmCache.LastLoad.value   = frmCache.inpEntry.options(frmCache.inpEntry.selectedIndex).LASTLOAD
   
   if msgbox("Are you sure you want to release " & chr(13) & frmCache.Client.value & "-" & frmCache.LOB.value & "-" & frmCache.DType.value & "   which was last loaded at " & frmCache.LastLoad.value & " from the list?",36,"FNS Net")=6 Then 
      frmCache.ConfirmedFlg.value = "Y"
   else
      frmCache.ConfirmedFlg.value = "N"
   end if

<% 

 If Request("status")  = "ok" and Request("ReleaseFlg") = "N" and Request("ConfirmedFlg") = "Y" Then
     
      dim CTCall   	
    
      set CTCall = nothing
      Set CTCall = Server.CreateObject("Call.Call.1")
      CTCall.ReloadFlowInList Request("CallFlowId"), Request("CallStFlg")
      set CTCall = nothing
   end if
   %>
   
  
end sub

sub onClickSubmitAll()

dim strMsg, arrList, strTemp, arrList2
dim strClient,strAhs,strLOB,strcallSt,strCallflowid,strLastLoadTime,strCallStDisp
   frmCache.ReleaseFlg.value = "Y"
   frmCache.StringList.value = frmCache.inpEntry.options(frmCache.inpEntry.selectedIndex).STRINGLIST
   
   
   strMsg = " Are you sure you want to release the following entries from the list ?" + Chr(13)
	  
   arrList=Split(frmCache.StringList.value,";")
   if IsArray(arrList) Then
      For Each strTemp In arrList
	     arrList2      = split(strTemp, ",")
		 strClient     = arrList2(0)
		 strAhs        = arrList2(1)
		 strLOB        = arrList2(2)
		 strCallSt     = arrList2(3)
		 strCallflowid = arrList2(4)
		 strLastLoadTime = arrList2(5)
		 
         if strCallSt = "Y" Then
		    strCallStDisp = "Call Start"
		 else
		    strCallStDisp = "CLaim Entry"
		 end if
	        strMsg = strMsg + chr(13) + strClient & "-" & strLOB & "-" & strCallStDisp & "-" & strLastLoadTime 
	     Next
   End if
   if msgbox(strMsg ,36,"FNS Net")=6 Then 
      frmCache.ConfirmedFlg.value = "Y"
   else
      frmCache.ConfirmedFlg.value = "N"
   end if
   
<% 
If Request("status")  = "ok" and Request("ReleaseFlg") = "Y" and Request("ConfirmedFlg") = "Y" Then
     
     ' dim CTCall   	
    
      set CTCall = nothing
      Set CTCall = Server.CreateObject("Call.Call.1")
      CTCall.ReleaseFlowLists
      set CTCall = nothing
   end if
   %>
   
  
end sub
</script>
<body CLASS="DataFrame" align=center leftmargin=0>

<% 

function getlist()

dim CTCall, arrList1,arrList2, bFirst, strTemp
dim strClient ,strAhs,strLOB,strCallSt,	strCallflowid, strLastLoadTime, strCallStDisp

	set CTCall = nothing
    Set CTCall = Server.CreateObject("Call.Call.1")
	
	strList = CTCall.GetCachedCallFlowList
    
	bFirst = true
	
	if strList =  "" Then
	   response.write (" No Entry found in Cache List ")
	Else
	   arrList1 = Split(strList,";")
	   if IsArray(arrList1) Then
	     
		 For Each strTemp In arrList1
		    arrList2      = split(strTemp, ",")
			strClient     = arrList2(0)
			strAhs        = arrList2(1)
			strLOB        = arrList2(2)
			strCallSt     = arrList2(3)
			strCallflowid = arrList2(4)
			strLastLoadTime = arrList2(5)
            
			
			   if strCallSt = "Y" Then
			      strCallStDisp = "Call Start"
			   else
			      strCallStDisp = "Claim Entry"
			   end if
			   If bFirst Then %>
			      <TABLE Width="100%">
                  <TR>
                  <TD Width="100%" BGCOLOR="#008b8b" CLASS="LABEL"><FONT COLOR="WHITE">Call Flow Cache List Maintenance</TD>
                  </TR>
                  </TABLE cellspacing=50 >
				  <br>
				  <table cellspacing=10>
				   <tr><td class="label">  Call Flow Cache Entries </td>
			      <td><select name="inpEntry" >
                 <option value="<%=strCallflowid%>,<%=strCallst%>" STRINGLIST="<%=strList%>" CALLFLOWID="<%=strCallflowid%>" CALLSTFLG="<%=strCallst%>" CLIENT="<%=strClient%>"  LOB="<%=strLOB%>" DTYPE="<%=strCallstDisp%>" LASTLOAD="<%=strLastLoadTime%>" SELECTED> <%=strClient%>-<%=strLOB%>-<%=strCallstDisp%>-<%=strLastLoadTime%> </option>
				 <% bFirst = False
			    Else 
				   'if bcontinue = false then%>
				   <option value="<%=strCallflowid%>,<%=strCallst%>" STRINGLIST="<%=strList%>" CALLFLOWID="<%=strCallflowid%>" CALLSTFLG="<%=strCallst%>" CLIENT="<%=strClient%>"  LOB="<%=strLOB%>" DTYPE="<%=strCallstDisp%>" LASTLOAD="<%=strLastLoadTime%>"> <%=strClient%>-<%=strLOB%>-<%=strCallstDisp%>-<%=strLastLoadTime%></option>
			 <% End if
		  Next
	   End If
	   response.write("</select></td></tr></table>")
	 End If

     Set CTCall = nothing
end function

'	response.write Request("listaction")

%>

<form name="frmCache" method="post" action="AM_CacheListMaintenance.asp" >
<p></p>
<% 	getList()
    If strList <>  "" Then %>
       <table cellspacing=10 align=left> 
       <tr></tr>       	   
	   <tr></tr>
	   <tr></tr>
	   <tr>  
       <td><input type=submit name="listaction" value="Release Selected Entry" onClick='onClickSubmit()'></td>
	   <td><input type=submit name="listaction1" value="Release All Entries" onClick= 'onClickSubmitAll()'></td>
	   <td><input type="hidden" name="status" value="ok"></td>
	   <td><input type="hidden" name="OK" ></td>
	   <td><input type="hidden" name="CallFlowId" ></td>
	   <td><input type="hidden" name="CallStFlg" ></td>
	   <td><input type="hidden" name="ReleaseFlg" value="N" ></td>
	   <td><input type="hidden" name="ConfirmedFlg" value="N" ></td>
	   <td><input type="hidden" name="Client" ></td>
	   <td><input type="hidden" name="LOB" ></td>
	   <td><input type="hidden" name="DType" ></td>
	   <td><input type="hidden" name="StringList" ></td>
	   <td><input type="hidden" name="LASTLOAD" ></td>
	   </tr>
	   </table>
    <%end if%>

</form>


</body>
</html>
