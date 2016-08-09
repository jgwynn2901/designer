<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>

<%
'On Error Resume Next
DIM cOriginalStatus, cOriginalParent, cError, cSQL, cExecSQL,cExecSQL1, s_SQLUpdateParent, s_ErrorUpdateParent
DIM oRS, rs_UpdateParent, cAHS,ahs_field_name,ahs_field_value

'MMAI-0007
'Prashant Shekhar 05/21/2007
'Declare the variables to be used by the functions.

Dim rs,temp,tele_claim_value,cat_loss_value,expo_ppo_value,fist_script_value,tcm_ind_value,triage_ind_value,acc_rec_value,gen_rout_value,spec_lost_time_value,spec_med_value,osha_recordable_value,longshore_value,edi_value,severity_value,monopolistic_value,selfadminindicator_value,srsclient_value,fromdatetip_value,todatetip_value,account_type_value,employer_report_level_value,ro_override_value,secure_email_value,policy_lookup_state_value,mask_ssn_value,cbpolicy_lookup_value,rentalreferral_value,customsidecar_value,staffingagency_value, textescdestination_value
Dim tele_claim_field,cat_loss_field,expo_ppo_field,fist_script_field,tcm_ind_field,triage_ind_field,acc_rec_field,gen_rout_field,spec_lost_time_field,spec_med_field,osha_recordable_field,longshore_field,edi_field,severity_field,monopolistic_field,selfadminindicator_field,srsclient_field,fromdatetip_field,todatetip_field,account_type_field,employer_report_level_field,ro_override_field,secure_email_field,policy_lookup_state_field,mask_ssn_field,cbpolicy_lookup_field,rentalreferral_field,customsidecar_field,staffingagency_field, textescdestination_field

ACTION = Request.Form("TxtAction")
cSQL = Request.Form("TxtSaveData")

'*********** MMAI 0019 Change ***********
dim checkType,clientNodeId
checkType = Request.Form("CheckType")
'clientNodeId = Request.Form("cOriginalClientNode")
'*********** MMAI 0019 Change ***********

cOriginalStatus = Request.Form("cOriginalStatus")
cOriginalParent = Request.Form("cOriginalParent")
cOriginalValidRule = Request.Form("cOriginalValidRule")
cAHS = Request.Form("AHSID")

'MMAI-0007
'Prashant Shekhar 06/13/2007
'Insert_New function performs the insert when a new risk location is created.

function Insert_New(ahs_field,ahs_value)

InsertSQL = "INSERT INTO AHS_EXTENSION (ACCNT_HRCY_STEP_ID, FIELD_NAME, FIELD_VALUE) VALUES "
			InsertSQL= InsertSQL & " ('" & NewAHSID & "','" & ahs_field & "','" & ahs_value & "')"
			Conn.Execute(InsertSQL)
end function




'MMAI-0007
'Prashant Shekhar 05/21/2007
' The Swap function takes the value of the checkbox and depending on whether
' it is checked or unchecked it assigns Y or N exept for the Account Record Indicator
' field. SwapRO fucntion covers the Account Record Indicator which stores RO when checked
' and Blank when unchecked.

function Swap(id)

	if UCase(id) = "ON" then
		 Swap = "Y"
	else
		Swap = "N"
	end if

end function

function SwapRO(id_RO)

	if UCase(id_RO) = "ON" then
		 SwapRO = "RO"
	else
		SwapRO = ""
	end if

end function

If ACTION = "UPDATE" Then

	cExecSQL = BuildSQL(cSQL,Chr(128), Chr(129), "UPDATE", "ACCOUNT_HIERARCHY_STEP", "ACCNT_HRCY_STEP_ID", "")

	Conn.Execute(cExecSQL)
	cError = CheckADOErrors(Conn,"ACCOUNT_HIERARCHY_STEP " & ACTION)
	If cError = ""  Then
		if request.form("ACTIVE_STATUS") <> cOriginalStatus then
			'	need to update all the nodes
			cExecSQL = "SELECT AHS.ACCNT_HRCY_STEP_ID, AHS_CHILD.PARENT_ID, AHS_CHILD.CHILD_ID "
			cExecSQL = cExecSQL & "FROM ACCOUNT_HIERARCHY_CHILD AHS_CHILD, ACCOUNT_HIERARCHY_STEP AHS "
			cExecSQL = cExecSQL & "WHERE AHS.ACCNT_HRCY_STEP_ID = AHS_CHILD.PARENT_ID AND (AHS.ACCNT_HRCY_STEP_ID=" & cAHS & ")"
			set oRS = Conn.Execute( cExecSQL )
			cExecSQL = "UPDATE ACCOUNT_HIERARCHY_STEP SET ACTIVE_STATUS = '" & request.form("ACTIVE_STATUS") & "' WHERE ACCNT_HRCY_STEP_ID = "
			do while not oRS.eof and cError = ""
				Conn.execute( cExecSQL & oRS.fields("CHILD_ID") )
				cError = CheckADOErrors(Conn,"ACCOUNT_HIERARCHY_STEP " & ACTION)
				oRS.movenext
			loop
			oRS.close
			set oRS = nothing
		end if

	'MMAI-0007
	'Prashant Shekhar               13/06/2007
	'The following code is used to update the ESIS flags.

	If cError = "" Then
		'MMAI-0019
		'Delete AHS Extension flags if there is an update from type Account to other
		If CheckType <> "ACCOUNT" then
		cExecSQL1 = "delete from AHS_Extension where ACCNT_HRCY_STEP_ID = '" & cAHS & "'"
		Conn.Execute(cExecSQL1)
		end if

		fromdatetip_value = Request.Form("ACC_FROM_DATE_TIP")
		if CheckType = "ACCOUNT" then
				fromdatetip_field = "CLAIM:ACCOUNT:FROM_DATE_TIP"
				todatetip_field = "CLAIM:ACCOUNT:TO_DATE_TIP"
				textescdestination_field = "CLAIM:ACCOUNT:TEXT_ESC_DESTINATION"
		elseif CheckType = "INSURED" then
				fromdatetip_field = "CLAIM:INSURED:FROM_DATE_TIP"
				todatetip_field = "CLAIM:INSURED:TO_DATE_TIP"
				textescdestination_field = "CLAIM:INSURED:TEXT_ESC_DESTINATION"
		elseif CheckType = "RISK LOCATION" then
				fromdatetip_field = "CLAIM:RISK_LOCATION:FROM_DATE_TIP"
				todatetip_field = "CLAIM:RISK_LOCATION:TO_DATE_TIP"
				textescdestination_field = "CLAIM:RISK_LOCATION:TEXT_ESC_DESTINATION"
		end if
		Insert_Update fromdatetip_field,fromdatetip_value

		todatetip_value = Request.Form("ACC_TO_DATE_TIP")

		Insert_Update todatetip_field,todatetip_value
		'REQ-2016-00467
		textescdestination_value = Request.Form("TEXT_ESC_DESTINATION")
        Insert_Update textescdestination_field,textescdestination_value

		If CheckType = "ACCOUNT" then
				tele_claim_value =  Swap(Request.Form("REVERSE_TELECLAIM_INDICATOR"))
				tele_claim_field = "CLAIM:ACCOUNT:CONCENTRA_REVTELECLAIM_FLG"
				Insert_Update tele_claim_field,tele_claim_value

				cat_loss_value = Swap(Request.Form("CONCENTRA_CAT_LOSS"))
				cat_loss_field = "CLAIM:ACCOUNT:CONCENTRA_CAT_FCM_FLG"
				Insert_Update cat_loss_field,cat_loss_value

				expo_ppo_value = Swap(Request.Form("EXPO_PPO_INDICATOR"))
				expo_ppo_field = "CLAIM:ACCOUNT:CONCENTRA_EXPO_PPO_FLG"
				Insert_Update expo_ppo_field,expo_ppo_value


				fist_script_value = Swap(Request.Form("FIRST_SCRIPT_INDICATOR"))
				fist_script_field = "CLAIM:ACCOUNT:CONCENTRA_FIRST_SCRIPT_FLG"
				Insert_Update fist_script_field,fist_script_value


				tcm_ind_value = Swap(Request.Form("CONCENTRA_TCM_INDICATOR"))
				tcm_ind_field = "CLAIM:ACCOUNT:CONCENTRA_TCM_FLG"
				Insert_Update tcm_ind_field,tcm_ind_value


				triage_ind_value = Swap(Request.Form("RN_TRIAGE_INDICATOR"))
				triage_ind_field = "CLAIM:ACCOUNT:RN_TRIAGE_FLG"
				Insert_Update triage_ind_field,triage_ind_value

				acc_rec_value = SwapRO(Request.Form("ACCOUNT_RECORD_INDICATOR"))
				acc_rec_field = "CLAIM:ACCOUNT:RO_INDICATOR"
				Insert_Update acc_rec_field,acc_rec_value


				gen_rout_value = Swap(Request.Form("GENERATE_ROUTING_RECORD"))
				gen_rout_field = "CLAIM:ACCOUNT:RO_ROUTING_FLG"
				Insert_Update gen_rout_field,gen_rout_value


				spec_med_value = Swap(Request.Form("SPECIAL_MEDICAL"))
				spec_med_field = "CLAIM:ACCOUNT:SPECIAL_MED_ONLY_FLG"
				Insert_Update spec_med_field,spec_med_value


				spec_lost_time_value = Swap(Request.Form("SPECIAL_LOST_TIME"))
				spec_lost_time_field= "CLAIM:ACCOUNT:SPECIAL_LOST_TIME_FLG"
				Insert_Update spec_lost_time_field,spec_lost_time_value

				'*********** MMAI-0023 ***********

				osha_recordable_value = Swap(Request.Form("CH_OSHA_RECORDABLE"))
				osha_recordable_field = "CLAIM:ACCOUNT:OSHA_RECORDABLE"
				Insert_Update osha_recordable_field,osha_recordable_value


				longshore_value = Swap(Request.Form("CH_LONGSHORE"))
				longshore_field = "CLAIM:ACCOUNT:LONGSHORE"
				Insert_Update longshore_field,longshore_value

				edi_value = Swap(Request.Form("CH_EDI"))
				edi_field = "CLAIM:ACCOUNT:GENERATE_EDI"
				Insert_Update edi_field,edi_value

				severity_value = Swap(Request.Form("CH_SEVERITY"))
				severity_field = "CLAIM:ACCOUNT:CUSTOM_SEVERITY_FLG"
				Insert_Update severity_field,severity_value

				monopolistic_value = Swap(Request.Form("CH_MONOPOLISTICSTATE"))
				monopolistic_field = "CLAIM:ACCOUNT:MONOPOLISTIC_STATE_HANDLIN"
				Insert_Update monopolistic_field,monopolistic_value

				'*********** MMAI-0385 ***********
				selfadminindicator_value = Swap(Request.Form("CH_SELFADMIN_INDICATOR"))
				selfadminindicator_field = "CLAIM:ACCOUNT:SELF_ADMIN_FLG"
				Insert_Update selfadminindicator_field,selfadminindicator_value

                '*********** BCAB-0379 ***********
				ro_override_value = Swap(Request.Form("CH_RO_OVERRIDE"))
				ro_override_field = "CLAIM:ACCOUNT:RO_OVERRIDE"
				Insert_Update ro_override_field,ro_override_value

				'*********** KKHU-0107 ***********
				srsclient_value = Swap(Request.Form("CH_SRS_CLIENT"))
				srsclient_field = "CLAIM:ACCOUNT:SRS_CLIENT"
				Insert_Update srsclient_field,srsclient_value
				
				'*********** MMAI-0516 ***********
				secure_email_value = Swap(Request.Form("CH_SECURE_EMAIL"))
				secure_email_field = "CLAIM:ACCOUNT:SECURE_EMAIL"
				Insert_Update secure_email_field,secure_email_value

				'*********** KFAB-0808 ***********

				account_type_value = Request.Form("ACCOUNT_TYPE")
				account_type_field = "CLAIM:ACCOUNT:ACCOUNT_TYPE"
				Insert_Update account_type_field,account_type_value

				'***********  END of KFAB-0808 ***********

				employer_report_level_value = Request.Form("SEL_EMPLOYER_REPORT_LEVEL")
				employer_report_level_field= "CLAIM:ACCOUNT:EMPLOYER_REPORT_LEVEL"
				Insert_Update employer_report_level_field,employer_report_level_value

				'*********** END Of MMAI-0023 ***********
				
				'*********** BCAB-0772 ***********

				policy_lookup_state_value = Request.Form("POLICY_LOOKUP_STATE")
				policy_lookup_state_field = "CLAIM:ACCOUNT:POLICY_LOOKUP_STATE"
				Insert_Update policy_lookup_state_field,policy_lookup_state_value

				'***********  END of BCAB-0772 ***********
				
				'*********** BCAB-0906 ***********
				mask_ssn_value = Swap(Request.Form("CH_MASK_SSN"))
				mask_ssn_field= "CLAIM:ACCOUNT:MASK_SSNO"
				Insert_Update mask_ssn_field,mask_ssn_value
				'*********** END BCAB-0906 ***********
				
				'*********** MMAI-0684 ***********
				cbpolicy_lookup_value = Swap(Request.Form("CH_CBPOLICY_LOOKUP"))
				cbpolicy_lookup_field = "CLAIM:ACCOUNT:CUSTOM_POLICY_LU"
				Insert_Update cbpolicy_lookup_field,cbpolicy_lookup_value
				'*********** END MMAI-0684 ***********
				
				'*********** MMAI-0708 ***********
				rentalreferral_value = Swap(Request.Form("CH_RENTAL_REFERRAL"))
				rentalreferral_field = "CLAIM:ACCOUNT:RENTAL_REFERRAL_FLG"
				Insert_Update rentalreferral_field,rentalreferral_value
				'*********** END MMAI-0708 ***********
				
				'*********** MMAI-0726 ***********
				customsidecar_value = Swap(Request.Form("CH_CUSTOM_SIDECAR"))
				customsidecar_field = "CLAIM:ACCOUNT:CUSTOM_SIDECAR_FLG"
				Insert_Update customsidecar_field,customsidecar_value
				'*********** END MMAI-0726 ***********	
				
				'*********** MMAI-0756 ***********
				staffingagency_value = Swap(Request.Form("CH_STAFFING_AGENCY"))
				staffingagency_field = "CLAIM:ACCOUNT:STAFFING_AGENCY_FLG"
				Insert_Update staffingagency_field,staffingagency_value
				'*********** END MMAI-0756 ***********					
		end if
	end if


		IF request.form("PARENT_NODE_ID") <> cOriginalParent THEN
			If cOriginalParent = "" Then cOriginalParent = "NULL"
			s_SQLUpdateParent = "{Call Account_Maintenance.ReBuild_AccountHierarchyChild(" & cAHS & ", " & cOriginalParent & ", {resultset 1, StatusNum})}"
			Set rs_UpdateParent = Conn.execute(s_SQLUpdateParent)
			IF Conn.Errors.Count = 0 Then
				If rs_UpdateParent("StatusNum") <> "0" Then
					s_ErrorUpdateParent = "ReBuild Account_Hierarchy_Child Failed.  Error(" & rs_UpdateParent("StatusNum") & ")."
				End IF
				Response.write("parent.frames('WORKAREA').UpdateOriginalParent('" & request.form("PARENT_NODE_ID") &  "');")
				rs_UpdateParent.Close
				Set rs_UpdateParent = Nothing
			ELSE
				s_ErrorUpdateParent = Conn.Errors(1).Description
			END IF
		END IF
		IF request.form("VALID_RULE_ID") <> cOriginalValidRule then
			IF len(request.form("VALID_RULE_ID")) = 0 then
				'	rule has been removed
				cExecSQL = "Delete from AHS_VALID_RULES Where ACCNT_HRCY_STEP_ID = " & cAHS
				Conn.Execute(cExecSQL)
			elseif len(cOriginalValidRule) = 0 then
				cExecSQL = "Insert into AHS_VALID_RULES Values(" & cAHS & "," & Request.Form("VALID_RULE_ID") & ")"
				Conn.Execute(cExecSQL)
			else
				cExecSQL = "Update AHS_VALID_RULES Set RULE_ID = " & request.form("VALID_RULE_ID") & " Where ACCNT_HRCY_STEP_ID = " & cAHS
				Conn.Execute(cExecSQL)
			end if
		 end if

	end if
Elseif ACTION = "INSERT" Then

	NewAHSID = CLng(NextPkey("ACCOUNT_HIERARCHY_STEP","ACCNT_HRCY_STEP_ID"))
	If NewAHSID > 0 Then
		cExecSQL = BuildSQL(cSQL, Chr(128), Chr(129), "INSERT", "ACCOUNT_HIERARCHY_STEP", "ACCNT_HRCY_STEP_ID", NewAHSID)
		Conn.Execute(cExecSQL)
		cError = CheckADOErrors(Conn,"AHS " & ACTION)
		If cError = ""  Then
			if len(Request.Form("VALID_RULE_ID")) <> 0 then
				cExecSQL = "Insert into AHS_VALID_RULES Values(" & NewAHSID & "," & Request.Form("VALID_RULE_ID") & ")"
				Conn.Execute(cExecSQL)
			end if

		fromdatetip_value = Request.Form("ACC_FROM_DATE_TIP")
		
		if CheckType = "ACCOUNT" then
				fromdatetip_field = "CLAIM:ACCOUNT:FROM_DATE_TIP"
				todatetip_field = "CLAIM:ACCOUNT:TO_DATE_TIP"
				textescdestination_field = "CLAIM:ACCOUNT:TEXT_ESC_DESTINATION"
		elseif CheckType = "INSURED" then
				fromdatetip_field = "CLAIM:INSURED:FROM_DATE_TIP"
				todatetip_value = "CLAIM:INSURED:TO_DATE_TIP"
				textescdestination_field = "CLAIM:INSURED:TEXT_ESC_DESTINATION"
		elseif CheckType = "RISK LOCATION" then
				fromdatetip_field = "CLAIM:RISK_LOCATION:FROM_DATE_TIP"
				todatetip_value = "CLAIM:RISK_LOCATION:TO_DATE_TIP"
				textescdestination_field = "CLAIM:RISK_LOCATION:TEXT_ESC_DESTINATION"
		end if
		Insert_New fromdatetip_field,fromdatetip_value

		todatetip_value = Request.Form("ACC_TO_DATE_TIP")

		Insert_New todatetip_field,todatetip_value
		'REQ-2016-00467
		textescdestination_value = Request.Form("TEXT_ESC_DESTINATION")
        Insert_New textescdestination_field,textescdestination_value

	'MMAI-0007
	'Prashant Shekhar   13/06/2007
	'The following code is used to Insert the ESIS flags into the AHS_Extension when a new account is created.

		If eError = "" then

			'*********** MMAI 0019 Change ***********
			If CheckType = "ACCOUNT" then
				'If clientNodeId = 202 then
					tele_claim_value =  Swap(Request.Form("REVERSE_TELECLAIM_INDICATOR"))
					tele_claim_field = "CLAIM:ACCOUNT:CONCENTRA_REVTELECLAIM_FLG"
					Insert_New tele_claim_field,tele_claim_value

					cat_loss_value = Swap(Request.Form("CONCENTRA_CAT_LOSS"))
					cat_loss_field = "CLAIM:ACCOUNT:CONCENTRA_CAT_FCM_FLG"
					Insert_New cat_loss_field,cat_loss_value

					expo_ppo_value = Swap(Request.Form("EXPO_PPO_INDICATOR"))
					expo_ppo_field = "CLAIM:ACCOUNT:CONCENTRA_EXPO_PPO_FLG"
					Insert_New expo_ppo_field,expo_ppo_value


					fist_script_value = Swap(Request.Form("FIRST_SCRIPT_INDICATOR"))
					fist_script_field = "CLAIM:ACCOUNT:CONCENTRA_FIRST_SCRIPT_FLG"
					Insert_New fist_script_field,fist_script_value


					tcm_ind_value = Swap(Request.Form("CONCENTRA_TCM_INDICATOR"))
					tcm_ind_field = "CLAIM:ACCOUNT:CONCENTRA_TCM_FLG"
					Insert_New tcm_ind_field,tcm_ind_value


					triage_ind_value = Swap(Request.Form("RN_TRIAGE_INDICATOR"))
					triage_ind_field = "CLAIM:ACCOUNT:RN_TRIAGE_FLG"
					Insert_New triage_ind_field,triage_ind_value


					acc_rec_value = SwapRO(Request.Form("ACCOUNT_RECORD_INDICATOR"))
					acc_rec_field = "CLAIM:ACCOUNT:RO_INDICATOR"
					Insert_New acc_rec_field,acc_rec_value


					gen_rout_value = Swap(Request.Form("GENERATE_ROUTING_RECORD"))
					gen_rout_field = "CLAIM:ACCOUNT:RO_ROUTING_FLG"
					Insert_New gen_rout_field,gen_rout_value


					spec_med_value = Swap(Request.Form("SPECIAL_MEDICAL"))
					spec_med_field = "CLAIM:ACCOUNT:SPECIAL_MED_ONLY_FLG"
					Insert_New spec_med_field,spec_med_value


					spec_lost_time_value = Swap(Request.Form("SPECIAL_LOST_TIME"))
					spec_lost_time_field= "CLAIM:ACCOUNT:SPECIAL_LOST_TIME_FLG"
					Insert_New spec_lost_time_field,spec_lost_time_value

					'*********** MMAI-0023 ***********

					osha_recordable_value = Swap(Request.Form("CH_OSHA_RECORDABLE"))
					osha_recordable_field = "CLAIM:ACCOUNT:OSHA_RECORDABLE"
					Insert_New osha_recordable_field,osha_recordable_value


					longshore_value = Swap(Request.Form("CH_LONGSHORE"))
					longshore_field = "CLAIM:ACCOUNT:LONGSHORE"
					Insert_New longshore_field,longshore_value


					edi_value = Swap(Request.Form("CH_EDI"))
					edi_field = "CLAIM:ACCOUNT:GENERATE_EDI"
					Insert_New edi_field,edi_value

					severity_value = Swap(Request.Form("CH_SEVERITY"))
					severity_field = "CLAIM:ACCOUNT:CUSTOM_SEVERITY_FLG"
					Insert_New severity_field,severity_value

					monopolistic_value = Swap(Request.Form("CH_MONOPOLISTICSTATE"))
					monopolistic_field = "CLAIM:ACCOUNT:MONOPOLISTIC_STATE_HANDLIN"
					Insert_New monopolistic_field,monopolistic_value

					'*********** MMAI-0385 ***********
					selfadminindicator_value = Swap(Request.Form("CH_SELFADMIN_INDICATOR"))
					selfadminindicator_field = "CLAIM:ACCOUNT:SELF_ADMIN_FLG"
					Insert_New selfadminindicator_field,selfadminindicator_value

                    '*********** BCAB-0379 ***********
					ro_override_value = Swap(Request.Form("CH_RO_OVERRIDE"))
					ro_override_field = "CLAIM:ACCOUNT:RO_OVERRIDE"
					Insert_New ro_override_field,ro_override_value

					'*********** KKHU-0107 ***********
					srsclient_value = Swap(Request.Form("CH_SRS_CLIENT"))
					srsclient_field = "CLAIM:ACCOUNT:SRS_CLIENT"
					Insert_New srsclient_field,srsclient_value
					
					'*********** MMAI-0516 ***********
					secure_email_value = Swap(Request.Form("CH_SECURE_EMAIL"))
					secure_email_field = "CLAIM:ACCOUNT:SECURE_EMAIL"
					Insert_New secure_email_field,secure_email_value

					'*********** KFAB-0808 ***********

					account_type_value = Request.Form("ACCOUNT_TYPE")
					account_type_field = "CLAIM:ACCOUNT:ACCOUNT_TYPE"
					Insert_New account_type_field,account_type_value

					'***********  END of KFAB-0808 ***********


					employer_report_level_value = Request.Form("SEL_EMPLOYER_REPORT_LEVEL")
					employer_report_level_field= "CLAIM:ACCOUNT:EMPLOYER_REPORT_LEVEL"
					Insert_New employer_report_level_field,employer_report_level_value

					'*********** END Of MMAI-0023 ***********
					
					'*********** BCAB-0772 ***********

					policy_lookup_state_value = Request.Form("POLICY_LOOKUP_STATE")
					policy_lookup_state_field = "CLAIM:ACCOUNT:POLICY_LOOKUP_STATE"
					Insert_New policy_lookup_state_field,policy_lookup_state_value

					'***********  END of BCAB-0772 ***********
					
					'*********** BCAB-0906 ***********
					mask_ssn_value = Swap(Request.Form("CH_MASK_SSN"))
					mask_ssn_field = "CLAIM:ACCOUNT:MASK_SSNO"
					Insert_New mask_ssn_field,mask_ssn_value
					'*********** END BCAB-0906 ***********
					
					'*********** MMAI-0684 ***********
					cbpolicy_lookup_value = Swap(Request.Form("CH_CBPOLICY_LOOKUP"))
					cbpolicy_lookup_field = "CLAIM:ACCOUNT:CUSTOM_POLICY_LU"
					Insert_New cbpolicy_lookup_field,cbpolicy_lookup_value
					'*********** END MMAI-0684 ***********
					
					'*********** MMAI-0708 ***********
					rentalreferral_value = Swap(Request.Form("CH_RENTAL_REFERRAL"))
					rentalreferral_field = "CLAIM:ACCOUNT:RENTAL_REFERRAL_FLG"
					Insert_New rentalreferral_field,rentalreferral_value
					'*********** END MMAI-0708 ***********
					
					'*********** MMAI-0726 ***********
					customsidecar_value = Swap(Request.Form("CH_CUSTOM_SIDECAR"))
					customsidecar_field = "CLAIM:ACCOUNT:CUSTOM_SIDECAR_FLG"
					Insert_New customsidecar_field,customsidecar_value
					'*********** END MMAI-0726 ***********					
				
					'*********** MMAI-0756 ***********
					staffingagency_value = Swap(Request.Form("CH_STAFFING_AGENCY"))
					staffingagency_field = "CLAIM:ACCOUNT:STAFFING_AGENCY_FLG"
					Insert_New staffingagency_field,staffingagency_value
					'*********** END MMAI-0756 ***********				
				'end if
			end if
		end if

			Response.write("parent.frames('WORKAREA').UpdateAHSID('" & NewAHSID &  "');")
		End If
	Else
		cError = "Unable to obtain next primary key for AHS table."
	End If
ElseIf ACTION = "DELETE" Then
	cExecSQL = "Delete from AHS_VALID_RULES Where ACCNT_HRCY_STEP_ID = " & cAHS
	Conn.Execute(cExecSQL)
	cError = CheckADOErrors(Conn,"Attribute " & ACTION)
	cExecSQL = BuildSQL("", "", "", "DELETE", "ACCOUNT_HIERARCHY_STEP", "ACCNT_HRCY_STEP_ID", cSQL)
	Conn.Execute(cExecSQL)
	cError = CheckADOErrors(Conn,"Attribute " & ACTION)
ElseIf ACTION = "RESET" Then
	cSQL = "{call Designer_3.resetMC(" & cAHS & ", {resultset 1, cStatusMsg, nStatusCode})}"
	Set oRS = Conn.Execute(cSQL)
	If oRS("nStatusCode") <> "0" Then
		cError = oRS("cStatusMsg")
	end if
	oRS.close
	set oRS = nothing
End If
Conn.Close
If cError <> "" Then
	LogStatusGroupBegin
	LogStatus S_ERROR, cError, "AHS", "", 0, ""
	LogStatusGroupEnd
%>
	parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update Unsuccessful, Check Status Report.");
	parent.frames("WORKAREA").SetDirty();
	parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true);
<%
ElseIf s_ErrorUpdateParent <> "" Then
%>
	parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> " & s_ErrorUpdateParent);
	parent.frames("WORKAREA").SetDirty();
	parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true);
<%
Else
	LogStatusGroupBegin
	LogStatusGroupEnd
	if ACTION = "RESET" Then
%>
	parent.frames("WORKAREA").UpdateStatus("Update Successful (MC).");
<%
	else
%>
	parent.frames("WORKAREA").UpdateStatus("Update Successful.");
<%
	end if
%>
	parent.frames("WORKAREA").ClearDirty();
	parent.frames("WORKAREA").SetStatusInfoAvailableFlag(false);
<%
	If ACTION	= "DELETE" Then
%>
		parent.frames("WORKAREA").UpdateScreenOnDelete();
<%
	End If
%>

<%
End If
 If ACTION = "UPDATE" and request.form("ACTIVE_STATUS") <> cOriginalStatus Then
%>
	parent.frames('WORKAREA').document.all.cOriginalStatus.value = "<%=request.form("ACTIVE_STATUS")%>"
	window.setTimeout("top.frames('WORKAREA').document.frames('LEFT').location.reload()",1500);
<%
end if
%>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
</html>