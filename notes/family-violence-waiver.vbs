'STATS GATHERING=============================================================================================================
name_of_script = "NOTES - Family Violence Waiver.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Required for statistical purposes===========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block==========================================================================================================

'Case number dialog
BeginDialog, 0, 0, 191, 105, "Family Violence Waiver"
  ButtonGroup ButtonPressed
    OkButton 45, 80, 50, 15
    CancelButton 105, 80, 50, 15
  EditBox 80, 10, 85, 15, MAXIS_case_number
  DropListBox 80, 30, 85, 15, "Approval"+chr(9)+"Closure", action_taken
  Text 15, 15, 50, 10, "Case Number:"
  Text 15, 30, 60, 10, "Action Taken:"
  DropListBox 80, 50, 85, 15, "Pre 60 Month"+chr(9)+"Post 60 Month", time_status
  Text 15, 50, 50, 10, "TIME Status:"
EndDialog

'THE SCRIPT==================================================================================================================

'Connects to BlueZone
EMConnect ""
'Checks Maxis for password prompt
'CALL check_for_MAXIS(True)

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------
DO
	err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
	Dialog
	IF ButtonPressed = cancel THEN StopScript          'If the user pushes cancel, stop the script

	IF IsNumeric(MAXIS_case_number) = FALSE  THEN err_msg = err_msg & vbNewLine & "* You must type a valid numeric case number."     'MAXIS_case_number should be mandatory in most cases. Bulk or nav scripts are likely the only exceptions
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."     '
LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!
'End dialog section-----------------------------------------------------------------------------------------------
memo_check = checked 'We default this to send the memo'
tikl_check = checked
'Call up the correct dialog based on case TYPE
IF action_taken = "Approval" THEN
'Dialog for pre-60 approval'
BeginDialog, 0, 0, 286, 200, "Family Violence Waiver Approval"
  ButtonGroup ButtonPressed
    OkButton 150, 175, 50, 15
    CancelButton 205, 175, 50, 15
  EditBox 75, 5, 110, 15, ES_plan_date
  Text 20, 10, 55, 10, "ES Plan Date:"
  EditBox 75, 25, 200, 15, verif_on_file
  Text 10, 30, 60, 10, "Verification on file:"
  EditBox 75, 45, 20, 15, approval_footer_month
  EditBox 110, 45, 20, 15, approval_footer_year
  Text 100, 45, 5, 15, "/"
  Text 10, 45, 60, 20, "Month approval begins: (MM/YY)"
  IF time_status = "Pre 60 Month" THEN CheckBox 10, 70, 70, 15, "MEMI Updated", MEMI_check
  IF time_status = "Pre 60 Month" THEN CheckBox 100, 70, 60, 15, "TIME Updated", TIME_check
  CheckBox 10, 90, 120, 15, "New MAXIS approval completed", approval_check
  CheckBox 10, 105, 220, 20, "Check here to have the script set a TIKL for three month review", TIKL_check
  CheckBox 10, 125, 240, 15, "Check here to have the script send a SPEC/MEMO about the approval", memo_check
	EditBox 150, 155, 100, 15, worker_signature
	Text 80, 155, 65, 15, "Worker Signature:"
EndDialog

	DO
		err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
		Dialog
		IF ButtonPressed = cancel THEN StopScript          'If the user pushes cancel, stop the script

		IF worker_signature = ""  THEN err_msg = err_msg & vbNewLine & "* You must sign your case note."     '
		IF ES_plan_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the date the employment service plan was received."
		IF isnumeric(approval_footer_month) = false or isnumeric(approval_footer_year) = false THEN err_msg = err_msg & vbNewLine & "* You must enter the footer month and year of approval for the waiver."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!
approval_date = cdate(approval_footer_month & "/01/" & approval_footer_year) 'convert the closure month into a date format for the notice
END IF


'dialog for pre-60 closure'
IF action_taken = "Closure" and time_status = "Pre 60 Month" THEN
BeginDialog , 0, 0, 286, 200, "Family Violence Waiver Approval"
  ButtonGroup ButtonPressed
    OkButton 150, 175, 50, 15
    CancelButton 205, 175, 50, 15
  EditBox 75, 10, 110, 15, ES_plan_date
  Text 10, 10, 55, 15, "Status Update Received:"
  EditBox 75, 35, 200, 15, closure_reason
  Text 10, 35, 60, 15, "Reason for closure:"
  EditBox 75, 65, 20, 15, closure_footer_month
  EditBox 110, 65, 20, 15, closure_footer_year
  Text 100, 65, 5, 15, "/"
  Text 10, 60, 60, 30, "Closure month: (first counted TANF month)"
  CheckBox 10, 105, 120, 15, "New MAXIS approval completed", approval_check
  CheckBox 10, 125, 240, 15, "Check here to have the script send a SPEC/MEMO about the approval", memo_check
  EditBox 150, 155, 100, 15, worker_signature
  Text 80, 155, 65, 15, "Worker Signature:"
EndDialog

DO
	err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
	Dialog
	IF ButtonPressed = cancel THEN StopScript          'If the user pushes cancel, stop the script
	IF ES_plan_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the date the status update was received from ES."
	IF worker_signature = ""  THEN err_msg = err_msg & vbNewLine & "* You must sign your case note."     'MAXIS_case_number should be mandatory in most cases. Bulk or nav scripts are likely the only exceptions
	IF isnumeric(closure_footer_month) = false or isnumeric(closure_footer_year) = false THEN err_msg = err_msg & vbNewLine & "* You must enter the footer month and year of closure."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."     '
LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!
closure_date = cdate(closure_footer_month & "/01/" & closure_footer_year)
END IF

IF action_taken = "Closure" and time_status = "Post 60 Month" THEN
	BeginDialog, 0, 0, 286, 170, "Closure Post 60th Month"
  	ButtonGroup ButtonPressed
    	OkButton 145, 150, 50, 15
    	CancelButton 200, 150, 50, 15
  	EditBox 120, 5, 80, 15, ES_plan_date
  	Text 10, 10, 100, 10, "Status update received date:"
  	EditBox 80, 25, 195, 15, verif_on_file
  	Text 10, 30, 65, 10, "Reason for closure:"
  	EditBox 120, 45, 20, 15, closure_footer_month
  	EditBox 155, 45, 20, 15, closure_footer_year
  	Text 145, 45, 5, 15, "/"
  	Text 10, 45, 110, 20, "Closure month: (first counted TANF month):"
  	CheckBox 10, 110, 240, 15, "Check here to have the script send a SPEC/MEMO about the approval", Check5
  	DropListBox 120, 65, 35, 15, "Yes"+chr(9)+"No", extension_available
  	Text 10, 70, 105, 10, "Other extension available?"
  	EditBox 80, 85, 195, 15, extension_details
  	Text 10, 90, 65, 10, "Extension details:"
  	EditBox 145, 130, 100, 15, worker_signature
  	Text 65, 135, 70, 10, "Worker Signature:"
	EndDialog

	DO
		err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
		Dialog
		IF ButtonPressed = cancel THEN StopScript          'If the user pushes cancel, stop the script
		IF ES_plan_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the date the status update was received from ES."
		IF worker_signature = ""  THEN err_msg = err_msg & vbNewLine & "* You must sign your case note."     'MAXIS_case_number should be mandatory in most cases. Bulk or nav scripts are likely the only exceptions
		IF isnumeric(closure_footer_month) = false or isnumeric(closure_footer_year) = false THEN err_msg = err_msg & vbNewLine & "* You must enter the footer month and year of closure."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."     '
	LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!
closure_date = cdate(closure_footer_month & "/01/" & closure_footer_year) 'convert the closure month into a date format for the notice

END IF



'Checks Maxis for password prompt
'CALL check_for_MAXIS(True)
'Navigate to SPEC/MEMO and send the appropriate note to the client
IF memo_check = checked THEN

'Navigating to SPEC/MEMO
call navigate_to_MAXIS_screen("SPEC", "MEMO")

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
	arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	PF5                                                     'PF5s again to initiate the new memo process
END IF
'Checking for SWKR
row = 4                             'Defining row and col for the search feature.
col = 1
EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
IF row > 4 THEN                     'If it isn't 4, that means it was found.
	swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	PF5                                           'PF5s again to initiate the new memo process
END IF
EMWriteScreen "x", 5, 12                                       'Initiates new memo to client
IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
transmit                                                        'Transmits to start the memo writing process

	'Writes the MEMO based on action and time status.
	IF action_taken = "Approval" and time_status = "Pre 60 Month" THEN
		call write_variable_in_SPEC_MEMO("You have been approved for a family violence waiver effective "& approval_date & "." &_
		" While you are on this waiver, your MFIP months will not count towards the 60 month life time limit. You are required to review your safety plan with your job" &_
		" counselor every three months to maintain this waiver. If you are not in compliance with your safety plan, the waiver could end.")
	END IF
	IF action_taken = "Approval" and time_status = "Post 60 Month" THEN
		call write_variable_in_SPEC_MEMO("You have been approved for a family violence waiver effective "& approval_date & "." &_
		" This waiver is your basis of extension for receiving MFIP beyond the 60 month lifetime limit.  You are required to review your safety plan with your job " &_
		"counselor every three months to maintain this waiver. If you are not in compliance with your safety plan or fail to review it with you job counselor, your MFIP could close.")
	END IF
	IF action_taken = "Closure" and time_status = "Pre 60 Month" THEN
		call write_variable_in_SPEC_MEMO("Your family violence waiver is ending. Your MFIP months will begin counting towards the 60-month lifetime limit again effective " & closure_date & "." &_
		" If you have any questions, please contact your job counselor.")
	END IF
	IF action_taken = "Closure" and time_status = "Post 60 Month" THEN
		call write_variable_in_SPEC_MEMO("Your family violence waiver is ending. This was your basis for extension of the 60 month lifetime limit. You will receive a " &_
		"separate notice regarding the status of your benefits. If you believe this waiver end is in error, please contact your job counselor." &_
		" If you have any questions concerning the status of your benefits, please contact the county. ")
	END IF
	PF4 'exit the memo'
END IF 'Close out the memo section'

'Write a TIKL if requested'
IF action_taken = "Approval" AND TIKL_check = checked THEN
	call navigate_to_MAXIS_screen("DAIL", "WRIT")
	call create_maxis_friendly_date(dateadd("m", 3, approval_date), 82, 5, 18)
	call write_variable_in_TIKL("Family Violence Waiver approaching 3 months, please review.")
	transmit
	PF3
END IF
'Now it navigates to a blank case note
start_a_blank_case_note

IF action_taken = "Approval" THEN
 	IF time_status = "Pre 60 Month" THEN CALL write_variable_in_case_note("* Family Violence Waiver Approved *")
	IF time_status = "Post 60 Month" THEN CALL write_variable_in_case_note("* Family Violence Waiver Approved for Extension *")
	CALL write_bullet_and_variable_in_case_note( "Safety plan received: ", ES_plan_date)
	CALL write_bullet_and_variable_in_case_note( "Verification on file: ", verif_on_file)
	Call write_variable_in_case_note("Family Violence Waiver Approved beginning: " & approval_date)
	IF memo_check = checked THEN call write_variable_in_case_note("Notice sent to client via SPEC/MEMO regarding the approval.")
END IF
IF action_taken = "Closure" THEN
	CALL write_variable_in_case_note("* Family Violence Waiver Ending *")
	CALL write_bullet_and_variable_in_case_note("Status Update Received: ", ES_plan_date)
	CALL write_bullet_and_variable_in_case_note("Reason for closure: ", closure_reason)
	CALL write_variable_in_case_note("Family Violvence Wavier ending effective: " & closure_date)
	IF memo_check = checked THEN call write_variable_in_case_note("Notice sent to client via SPEC/MEMO regarding the waiver end.")
	IF extension_available = "YES" THEN call write_bullet_and_variable_in_case_note("Client switching to new extension: ", extension_details)
	IF extension_available = "NO" THEN call write_variable_in_case_note("* No other extension is documented a this time.")
END IF

IF approval_check = checked THEN CALL write_variable_in_case_note("New MAXIS approval completed.")
IF MEMI_check = checked THEN CALL write_variable_in_case_note("MEMI updated.")
IF TIME_check = checked THEN CALL write_variable_in_case_note("TIME updated.")
IF TIKL_check = checked THEN call write_variable_in_case_note("TIKL set for three month review.")

'...and a worker signature.
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")
