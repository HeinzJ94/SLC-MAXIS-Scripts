'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - COLA OTHER INCOME.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 345         'manual run time in seconds
STATS_denomination = "C"       'C is for each MEMBER
'END OF stats block==============================================================================================

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "David Courtright, St Louis")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Go grab the info from UNEA!
'GOING TO STAT
EMSendKey "s"
transmit
EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")
EMWriteScreen "UNEA", 20, 71
Transmit




'DIALOGS----------------------------------------------------------------------------------------------
BeginDialog Dialog1, 0, 0, 281, 140, "Other Retirement COLA"
  ButtonGroup ButtonPressed
    OkButton 170, 120, 50, 15
    CancelButton 225, 120, 50, 15
  CheckBox 5, 75, 180, 10, "Enter a TIKL for followup.", TIKL_check
  EditBox 70, 10, 205, 15, income_source
  EditBox 45, 55, 230, 15, other_notes
  Text 5, 15, 60, 10, "Income Source:"
  Text 5, 60, 25, 10, "Notes:"
  CheckBox 5, 30, 275, 20, "Documentation exists showing this income does not have annual COLA changes.", unchanging_check
  EditBox 170, 95, 105, 15, worker_signature
  Text 95, 100, 65, 10, "Worker Signature:"
EndDialog

'Show dialog
do
	Do
		Dialog
		cancel_confirmation
		MAXIS_dialog_navigation
	Loop until ButtonPressed = -1
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Navigates back to DAIL
Do
	EMReadScreen DAIL_check, 4, 2, 48
	If DAIL_check = "DAIL" then exit do
	PF3
Loop until DAIL_check = "DAIL"

'Navigates to case note
EMSendKey "n"
transmit

'Creates blank case note
PF9
transmit



'Writes that the message is unreported, and that the proofs are being sent/TIKLed for.
call write_variable_in_case_note("COLA for other retirement income from " & income_source & ".")
IF unchanging_check = checked THEN
	call write_variable_in_case_note("Documentation in case file shows this income does not change annually. No updates to case.")
ELSE
	call write_variable_in_case_note("* Sent DHS-2919 requesting verification of COLA changes from this income source.")
	If TIKL_check = checked then call write_variable_in_case_note("* TIKLed for 10 days for return.")
END IF
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature & ", using automated script.")
PF3
PF3

'If TIKL_checkbox is unchecked, it needs to end here.
If TIKL_check = unchecked then script_end_procedure("Success! Don't forget to send DHS2919 requesting verification of income changes if necessary. The income is from: " & income_source & ".")



IF TIKL_check = checked THEN
  'Navigates to TIKL
  EMSendKey "w"
  transmit

  'The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
  call create_MAXIS_friendly_date(date, 11, 5, 18)

  'Setting cursor on 9, 3, because the message goes beyond a single line and EMWriteScreen does not word wrap.
  EMSetCursor 9, 3

  'Sending TIKL text.
  call write_variable_in_TIKL("Verification of " & income_source & " income COLA change should have been returned. If not received and processed, take appropriate action. (TIKL auto-generated from script).")

  'Submits TIKL
  transmit
  'Exits TIKL
  PF3
End IF

'Writing the case note'
'Exits script and logs stats if appropriate
script_end_procedure("Success! Case note made, and a TIKL has been sent for Jan 1. A DHS-2919 should now be sent. The income is from: " & income_source & ".")
