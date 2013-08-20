<%Response.CacheControl="no-cache"
Response.AddHeader "Pragma","no-cache"
Response.Expires=-1%>

<!--#include file="../lib/genericSQL.asp"-->

<%
dim method, parms

method=trim(request.QueryString("method"))
parms=trim(request.QueryString("parms"))

if method = "UserExists" and parms <> "" then
    response.Write(UserExists(parms))

end if

function UserExists(username)
dim cSQL,oRS, iUserCount
    cSQL = "Select Count(USER_ID) as UserCount From USERS Where Name = '" & username & "'"
set oRS = Conn.Execute(cSQL)

with oRS
	iUserCount = .fields("UserCount")
	.close
end with
set oRS = nothing
UserExists = iUserCount

end function
%>