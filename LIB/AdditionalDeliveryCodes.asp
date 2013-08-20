<OPTION VALUE='AGNT'>Agent
<OPTION VALUE='BR'>Branch
<OPTION VALUE='CA'>Caller
<OPTION VALUE='INS'>Insured
<OPTION VALUE='MONO'>Mono. State
<OPTION VALUE='RL'>Risk Location
<OPTION VALUE='ST'>State

<% If Left(Session("ConnectionString"), 7) = "DSN=CRA" Then
	Response.Write("<OPTION VALUE='CAF'>Corp CAU w/ Form")
	Response.Write("<OPTION VALUE='CGF'>Corp CLI w/ Form")
	Response.Write("<OPTION VALUE='CPF'>Corp CPR w/ Form")
	Response.Write("<OPTION VALUE='CWF'>Corp WOR w/ Form")
	Response.Write("<OPTION VALUE='CAN'>Corp CAU w/o Form")
	Response.Write("<OPTION VALUE='CGN'>Corp CLI w/o Form")
	Response.Write("<OPTION VALUE='CPN'>Corp CPR w/o Form")
	Response.Write("<OPTION VALUE='CWN'>Corp WOR w/o Form")
	Response.Write("<OPTION VALUE='RAF'>RL CAU w/ Form")
	Response.Write("<OPTION VALUE='RGF'>RL CLI w/ Form")
	Response.Write("<OPTION VALUE='RPF'>RL CPR w/ Form")
	Response.Write("<OPTION VALUE='RWF'>RL WOR w/ Form")
	Response.Write("<OPTION VALUE='RAN'>RL CAU w/o Form")
	Response.Write("<OPTION VALUE='RGN'>RL CLI w/o Form")
	Response.Write("<OPTION VALUE='RPN'>RL CPR w/o Form")
	Response.Write("<OPTION VALUE='RWN'>RL WOR w/o Form")   
End If
%>
