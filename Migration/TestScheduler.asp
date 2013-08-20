<HTML>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<HEAD>
</HEAD>
<BODY CLASS="LABEL" BGCOLOR="#d6cfbd">
<%

	'On Error Resume Next

	Set oScheduler =  Server.CreateObject("Scheduler.SchedulingAgent.1")
	if not isobject(oScheduler) then
		response.write("Could not create Scheduler Object.")
		response.end
	end if
	Set oTask = oScheduler.Tasks.Add("DMU_TEST_SCHEDULER1")
	if err.number <> 0 then
		response.write ("Error adding Task: " & err.description & " - " & err.number)
		response.end
	end if
	oTask.SetAccountInformation "WAL_CMCS_DEV\fnsdesigner", "fnsd"
	oTask.ApplicationName = "C:\WINNT\SYSTEM32\CALC.EXE"
	oTask.creator = "ASP Page"
	oTask.Flags = 194
	oTask.Priority = 32
	oTask.TaskFlags = 0
	oTask.WorkingDirectory = "C:\WINNT\SYSTEM32\"
	
	set oTR = oTask.Triggers.Add()
	if err.number <> 0 then
		response.write ("Error adding Trigger: " & err.description & " - " & err.number)
		response.end
	end if
    
	oTR.Duration = 0
	oTR.Flags = 0
	oTR.Interval = 0
	oTR.TriggerType = 0
	oTR.BeginTime = DateAdd("n", 2, Now)

	set oScheduler = nothing
	set oTR = nothing
	set oTask = nothing
%>
		Migration Job scheduled. <%=Request.QueryString("BeginTime")%>
</BODY>
