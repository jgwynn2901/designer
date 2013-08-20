<!--#include file="..\lib\common.inc"-->
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

Function ReplaceStr(TextIn, SearchStr , Replacement)
	Dim WorkText, Pointer
    WorkText = TextIn
    Pointer = InStr(1, WorkText, SearchStr)
    Do While Pointer > 0
      WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
      Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr)
    Loop
    ReplaceStr = WorkText
End Function

Function NextPkey( TableName, ColName )
On error resume next
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName & "', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	strError = CheckADOErrors(Conn,"Users " & ACTION)
	NextPkey = NextRS("outResult") 

End Function

Function SplitString(UsrString, Delimeter)
If Len(Trim(UsrString)) < 1 Then
	SplitString = null
	Exit function
End If

Dim posd, i, Tempstr, lngstr
Dim RArray(200)
i = 1
posd = 1
lngstr = Len(UsrString)
Tempstr = Trim(UsrString)

Do While posd <> 0
    posd = InStr(Tempstr, Delimeter)
    If posd <> 0 Then
        Rarray(i) = Trim(Left(Tempstr, posd - 1))
    Else
        Rarray(i) = Trim(Tempstr)
    End If
    Tempstr = Right(UsrString, lngstr - posd)
    lngstr = Len(Tempstr)
    i = i + 1
Loop
SplitString = Rarray
End Function


Function SplitStringNoTrim(UsrString, Delimeter)
If Len(UsrString) < 1 Then
	SplitStringNoTrim = null
	Exit function
End If

Dim posd, i, Tempstr, lngstr
Dim RArray(200)
i = 1
posd = 1
lngstr = Len(UsrString)
Tempstr = UsrString

Do While posd <> 0
    posd = InStr(Tempstr, Delimeter)
    If posd <> 0 Then
        Rarray(i) = Left(Tempstr, posd - 1)
    Else
        Rarray(i) = Tempstr
    End If
    Tempstr = Right(UsrString, lngstr - posd)
    lngstr = Len(Tempstr)
    i = i + 1
Loop
SplitStringNoTrim = Rarray
End Function


'******************************MMAI-0007****************
'******************************Prashant Shekhar 05/21/2007****************
'****Insert_Update function checks whether the existing field exists in the Database for the ***************
'** Accnt_Hrcy_Step_ID and then performs an  update if it exists else it performs an  Insert.*****

sub Insert_Update(ByVal ahs_field_name,ByVal ahs_field_value)

	cExecSQL1 = "select count (*) as count from AHS_Extension where field_name = '" & ahs_field_name & "' AND ACCNT_HRCY_STEP_ID = '" & cAHS & "'"
				
	rs = Conn.Execute(cExecSQL1)
	temp = rs(0)
	
		if cint(temp) >0 then

			updateSQL = "UPDATE  AHS_EXTENSION SET FIELD_VALUE = '" & ahs_field_value & "' WHERE field_name = '" & ahs_field_name & "' AND ACCNT_HRCY_STEP_ID = '" & cAHS & "' "
			Conn.Execute(updateSQL)		
		else 
			InsertSQL = "INSERT INTO AHS_EXTENSION (ACCNT_HRCY_STEP_ID, FIELD_NAME, FIELD_VALUE) VALUES "
			InsertSQL= InsertSQL & " ('" & cAHS & "','" & ahs_field_name & "','" & ahs_field_value & "')"
			Conn.Execute(InsertSQL)	

		end if
	 
end sub

'*********************************************************



'*********************************************************

'**************************************
' UserString = NAME~BOB~1|ADDRESS~21 JUMP STREET~1|AGE~45~0
' 1 = string
' 0 = number
'
'**************************************
Function BuildSQL(UserString, FieldDelimeter, ValueDelimeter, Action, TableName, PrimaryKey, PrimaryKeyValue)
MySQLSelect = ""
MySQL = ""
MyWhere = ""

lret = SplitString(UserString, FieldDelimeter)
Select Case Action
	Case "UPDATE"
		MySQLSelect = MySQLSelect & "UPDATE " & TableName & " SET "
		for i = 1 to UBOUND(lret) step 1
			
		If lret(i) <> "" Then
		lret2 = SplitString(lret(i), ValueDelimeter)
		
			If Ucase(lret2(1)) <> Ucase(PrimaryKey) Then
				MySQL = MySQL & lret2(1) & "="
				If lret2(3) = "1" Then
					'We need the trimmed off spaces for this one field
					if(lret2(1) = "ADDITIONAL_DELIVERIES") Then
						lret2(2) = " " & lret2(2) & " "
					end if
					MySQL = MySQL & "'" & ReplaceStr(lret2(2), "'", "''") & "',"
				Else
					MySQL = MySQL & lret2(2) & ","
				End If
			Else
				MyWhere = MyWhere & " WHERE " & Ucase(PrimaryKey) & "=" 
				if len(PrimaryKeyValue) = 0 then
					If lret2(3) = "1" Then
						MyWhere = MyWhere & "'" & ReplaceStr(lret2(2), "'", "''") & "'"
					Else 
						MyWhere = MyWhere & lret2(2)
					End If
				else
					If lret2(3) = "1" Then
						MyWhere = MyWhere & "'" & ReplaceStr(PrimaryKeyValue, "'", "''") & "'"
					Else 
						MyWhere = MyWhere & PrimaryKeyValue
					End If
				end if
			End If
	End If
		Next
	BuildSQL = MySQLSelect & mid(MySQL, 1, (len(MySQL) -1)) & MyWhere
	
	Case "INSERT"
		MySQLSelect = MySQLSelect & "INSERT INTO " & TableName
		for i = 1 to UBOUND(lret) step 1
			If lret(i) <> "" Then
				lret2 = SplitString(lret(i), ValueDelimeter)
				
				If 	Ucase(lret2(1)) <> Ucase(PrimaryKey) Then
					FIELDS = FIELDS & lret2(1) & ","
					If lret2(3) = "1" Then
						VALUES= VALUES & "'" & ReplaceStr(lret2(2), "'", "''") & "',"
					Else
						VALUES= VALUES & lret2(2) & ","
					End If
				Else
					FIELDS = FIELDS & lret2(1) & ","
					If PrimaryKeyValue = "" Then
						PrimaryKeyValue = lret2(2)
					End if
					If lret2(3) = "1" Then
						VALUES= VALUES & "'" & ReplaceStr(PrimaryKeyValue, "'", "''") & "',"
					Else
						VALUES= VALUES & PrimaryKeyValue & ","
					End If
				End If

			End If
		Next
		FIELDS = " (" & mid(FIELDS,1, (len(FIELDS)-1)) & ")"
		VALUES = " VALUES (" & mid(VALUES,1, (len(VALUES)-1)) & ")"
	BuildSQL = MySQLSelect & FIELDS & VALUES
	
	Case "DELETE"
		BuildSQL = MySQL & "DELETE FROM " & TableName & " WHERE " & PrimaryKey & "=" & PrimaryKeyValue
	Case Else
	BuildSQL = "ERROR"
End Select 

End Function

Function BuildSQLNoTrim(UserString, FieldDelimeter, ValueDelimeter, Action, TableName, PrimaryKey, PrimaryKeyValue)
MySQLSelect = ""
MySQL = ""
MyWhere = ""

lret = SplitStringNoTrim(UserString, FieldDelimeter)
Select Case Action
	Case "UPDATE"
		MySQLSelect = MySQLSelect & "UPDATE " & TableName & " SET "
		for i = 1 to UBOUND(lret) step 1
			
	If lret(i) <> "" Then
		lret2 = SplitStringNoTrim(lret(i), ValueDelimeter)
		
			If Ucase(lret2(1)) <> Ucase(PrimaryKey) Then
				MySQL = MySQL & lret2(1) & "="
				If lret2(3) = "1" Then
					MySQL = MySQL & "'" & ReplaceStr(lret2(2), "'", "''") & "',"
				Else
					MySQL = MySQL & lret2(2) & ","
				End If
			Else
				MyWhere = MyWhere & " WHERE " & Ucase(PrimaryKey) & "=" 
				If lret2(3) = "1" Then
					MyWhere = MyWhere & "'" & ReplaceStr(lret2(2), "'", "''") & "'"
				Else 
					MyWhere = MyWhere & lret2(2)
				End If
			End If
	End If
		Next
	BuildSQLNoTrim = MySQLSelect & mid(MySQL, 1, (len(MySQL) -1)) & MyWhere
	
	Case "INSERT"
		MySQLSelect = MySQLSelect & "INSERT INTO " & TableName
		for i = 1 to UBOUND(lret) step 1
			If lret(i) <> "" Then
				lret2 = SplitStringNoTrim(lret(i), ValueDelimeter)
				
				If 	Ucase(lret2(1)) <> Ucase(PrimaryKey) Then
					FIELDS = FIELDS & lret2(1) & ","
					If lret2(3) = "1" Then
						VALUES= VALUES & "'" & ReplaceStr(lret2(2), "'", "''") & "',"
					Else
						VALUES= VALUES & lret2(2) & ","
					End If
				Else
					FIELDS = FIELDS & lret2(1) & ","
					If PrimaryKeyValue = "" Then
						PrimaryKeyValue = lret2(2)
					End if
					If lret2(3) = "1" Then
						VALUES= VALUES & "'" & ReplaceStr(PrimaryKeyValue, "'", "''") & "',"
					Else
						VALUES= VALUES & PrimaryKeyValue & ","
					End If
				End If

			End If
		Next
		FIELDS = " (" & mid(FIELDS,1, (len(FIELDS)-1)) & ")"
		VALUES = " VALUES (" & mid(VALUES,1, (len(VALUES)-1)) & ")"
	BuildSQLNoTrim = MySQLSelect & FIELDS & VALUES
	
	Case "DELETE"
		BuildSQLNoTrim = MySQL & "DELETE FROM " & TableName & " WHERE " & PrimaryKey & "=" & PrimaryKeyValue
	Case Else
	BuildSQLNoTrim = "ERROR"
End Select 

End Function
%>

