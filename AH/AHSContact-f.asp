<!-- #include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!-- #include file="..\lib\RenderTextinc.asp"-->


<%
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	Dim ContainerType, Mode
	Dim s_SQL2, Conn, ConnectionString, rs, RS2, s_Select, ClientCode, s_lob, COID
	Dim s_SeqStepID,s_DestinationID,s_SeqNumber ,s_RetryCount ,s_RetryWaitTime ,s_DestinationStr,s_Transmission 

	AccountTextLen = 28
	'ContainerType = "Modal"
	ClientCode = Request.QueryString("ClientCode")
   '** Check for appropriate privs...	
   'If HasModifyPrivilege("FNSD_SPECIFICDESTINATION", SECURITYPRIV) <> True Then MODE = "RO"
   ' Mode in this case is either EDIT or NEW
   '******

    Mode = Request.Querystring("MODE")
    COID	= CStr(Request.QueryString("COID"))
	's_accnt_hrcy_step_id =  62
	's_account_name = ""
	' If EDIT was clicked then display data for that COID etc..
	' Only 1 row shd. be retrieved since we r specifying AHSCID

	IF Request.QueryString("MODE") <> "NEW" THEN
		SET Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = CONNECT_STRING
		if ClientCode <> "SED" then
		s_Select = "Select ahsc.ahs_contact_id, " &_
		           "       ahsc.contact_id, " &_
				   "	   ahsc.accnt_hrcy_step_id,"  &_ 
				   "	   to_char(ahsc.active_start_dt,'MM/DD/YYYY') as active_start_dt, " &_
				   "	   to_char(ahsc.active_end_dt,'MM/DD/YYYY') as active_end_dt, " &_
	               "       ahs.name  "  &_
				   "  From ahs_contact ahsc, account_hierarchy_step ahs " &_
		           " Where ahs.accnt_hrcy_step_id = ahsc.accnt_hrcy_step_id " &_
				   "   and ahs_contact_id         = " & Request.QueryString("AHSCID")
			else
			
			s_Select = "Select ahsc.ahs_contact_id, " &_
		           "       ahsc.contact_id, " &_
				   "	   ahsc.accnt_hrcy_step_id,"  &_ 
				   "	   to_char(ahsc.active_start_dt,'MM/DD/YYYY') as active_start_dt, " &_
				   "	   to_char(ahsc.active_end_dt,'MM/DD/YYYY') as active_end_dt, " &_
	               "       ahs.name,ahsc.lob_cd  "  &_
				   "  From ahs_contact ahsc, account_hierarchy_step ahs " &_
		           " Where ahs.accnt_hrcy_step_id = ahsc.accnt_hrcy_step_id " &_
				   "   and ahs_contact_id         = " & Request.QueryString("AHSCID")
				   
			end if
			
		Conn.Open ConnectionString
		'response.write(s_Select)
		

		SET rs = Conn.Execute(s_Select)	
		If not rs.EOF Then
		   s_ahs_contact_id     = rs("ahs_contact_id")
		   s_contact_id         = rs("contact_id")
		   s_accnt_hrcy_step_id = rs("accnt_hrcy_step_id")
		   s_active_start_dt    = rs("active_start_dt")
		   s_active_end_dt      = rs("active_end_dt")
		   s_account_name       = rs("name")
		   if ClientCode = "SED" then
		   s_lob								= rs("lob_cd")
		   end if
		   
		End If
		rs.Close
		Conn.Close
		SET rs = NOTHING
		SET Conn = NOTHING
	
	END IF

	
%>

<HTML>
<HEAD>
<!--#include file="..\lib\tablecommon.inc"-->
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>AHS Contacts</Title>
<Link Rel="StyleSheet" Type="text/css" Href="..\FNSDESIGN.CSS">
</HEAD>

<frameset  ROWS="0,*" border="0" framespacing="0">
   		<frame NAME="hiddenPage" SRC="ABOUT:BLANK" scrolling="No" noresize FRAMEBORDER="no" BORDER="0" framespacing="0">
		<% IF ClientCode <> "SED" then %>
		<frame NAME="WORKAREA" SRC="ContactDetailsAHSSummaryModal.asp?<%=Request.QueryString%>"   SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
		<%else %>
		<frame NAME="WORKAREA" SRC="SEDContactDetailsAHSSummaryModal.asp?<%=Request.QueryString%>"   SCROLLING="NO" FRAMEBORDER="no" NORESIZE BORDER="0" framespacing="0">
		<%end if %>
        
	</frameset>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
</BODY>

</HTML>
