<%
Dim MAXRECORDCOUNT, MSG_MAXRECORDS
MAXRECORDCOUNT = Cint(Session("USERMAXRECORDS"))
MSG_MAXRECORDS = "Record count is at the maximum of " & Session("USERMAXRECORDS") & ", please refine your search criteria"
Const MSG_SEARCH = "Searching..."
Const MSG_PROMPT = "Enter criteria and choose Search"
Const MSG_FILTER_OFF = " Filter Removed"
Const MSG_FILTER_ON = " Filter Applied"
%>