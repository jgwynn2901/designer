<%@ Language=VBScript %>
<% Response.Expires=0 %>
<!--#include file="lib\common.inc"-->
<!--#include file="lib\security.inc"-->

<HTML>
<HEAD>
        <FRAMESET  COLS="35%,*">
<%	ViewTree = false
	If HasAutomaticSecurityPrivilege() = True Or HasViewPrivilege("FNSD_TREE","")= True Then ViewTree = True%>

<%	If ViewTree = True Then %>			
			<FRAMESET  ROWS="70%,*">
				<FRAME NAME="LEFT" SRC="AH/AHTree.asp?MAXRS=<%=Session("USERTREECOUNT")%>&MAXLEVEL=99&Build_Tree" SCROLLING="no">
<%	End if %>				
				<frame name="FAVORITES" src="AH/favorites.asp" marginwidth="0" marginheight="0" scrolling="NO" frameborder="no">
<%	If ViewTree = True Then %>			
			</FRAMESET>         
<%	End if %>				
			<FRAME NAME="WORK" SRC="About:Blank" SCROLLING="AUTO" FRAMEBORDER="no">
        </FRAMESET>
        
</HEAD>
<BODY>
</BODY>
</HTML>

