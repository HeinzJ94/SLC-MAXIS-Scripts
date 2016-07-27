'Required for statistical purposes===============================================================================
name_of_script = "BULK - EXP SNAP REVIEW.vbs"
start_time = timer
STATS_counter = 1                           'sets the stats counter at one
STATS_manualtime = 29                       'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
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

'DIALOGS-------------------------------------------------------------------------------------------------------------
BeginDialog EXP_SNAP_review_dialog, 0, 0, 286, 185, "EXP SNAP review "
  EditBox 75, 10, 205, 15, worker_number
  CheckBox 25, 50, 140, 10, "Check here to run this query county-wide ", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 175, 45, 50, 15
    CancelButton 230, 45, 50, 15
  Text 5, 15, 65, 10, "Worker(s) to check:"
  Text 10, 90, 270, 20, "This script will create a list of cases that should be reviewed for expedited SNAP eligibilty from REPT/PND1 and REPT/PND2."
  Text 5, 30, 275, 10, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  GroupBox 5, 75, 280, 105, "BULK - Expedited SNAP review"
  Text 10, 115, 270, 30, "* The REPT/PND1 list of cases will include ALL cases that do not have a case note that identifies the case as not expedited. This includes cases that are not pending for SNAP since REPT/PND1 does not make the distinction."
  Text 10, 145, 270, 25, "*The REPT/PND2 list of cases will identify ALL cases that are pending for SNAP (or MFIP if SNAP isn't active) that do not have a case note that identifies the case as not expedited."
EndDialog

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
'Determining specific county for multicounty agencies & connects to BlueZone
get_county_code
EMConnect ""

'Shows dialog
DO 
	Do 
		err_msg = ""
    	Dialog EXP_SNAP_review_dialog
    	If buttonpressed = cancel then script_end_procedure("")
		If worker_number = "" then err_msg = err_msg & vbNewLine & "* You must enter at least one worker number."
		If worker_number <> "" AND all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* You must select either a worker number(s) or agency-wide, not both."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect						
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'sets up PRIV case array
Dim PRIV_case_array()
Redim PRIV_case_array(5, 0)

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>PND1 information
'Sets up the array to store all the information for each client'
Dim PND1_array ()
ReDim PND1_array (5, 0)
entry_record = 0

'Sets constants for the array to make the script easier to read (and easier to code)
Const work_num     = 1     			
Const case_num     = 2		'Each of the case numbers will be stored at this position	
Const clt_name     = 3
Const app_date     = 4
Const days_pending = 5

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "PND1")
	EMWriteScreen worker, 21, 13
	transmit
	
	'Skips workers with no info
	EMReadScreen has_content_check, 8, 7, 3
	If has_content_check <> "        " then
		'Grabbing each case number and case information 
		Do
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3		 'Reading case number
				MAXIS_case_number = trim(MAXIS_case_number)
				EMReadScreen worker_basket, 7, 21, 13
				EMReadScreen client_name, 25, MAXIS_row, 13			 'Reading client name
				client_name = trim(client_name)
				EMReadScreen appl_date, 8, MAXIS_row, 41		     'Reading application date
				appl_date = replace(appl_date, " ", "/")
				EMReadScreen nbr_days_pending, 4, MAXIS_row, 54		 'Reading nbr days pending
					
				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				If MAXIS_case_number = "" and client_name = "" then exit do			'Exits do if we reach the end
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
				
				'Adding client information to the array'
				ReDim Preserve PND1_array(5, entry_record)	'This resizes the array based on the number of rows in the Excel File
				'The client information is added to the array
				PND1_array (work_num,     entry_record) = worker_basket
				PND1_array (case_num,	  entry_record) = MAXIS_case_number		
				PND1_array (clt_name,  	  entry_record) = client_name
				PND1_array (app_date, 	  entry_record) = appl_date
				PND1_array (days_pending, entry_record) = nbr_days_pending
					
				entry_record = entry_record + 1			'This increments to the next entry in the array
				MAXIS_row = MAXIS_row + 1	
				STATS_counter = STATS_counter + 1       'adds one instance to the stats counter	
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'Now the script goes into CASENOTE and searches for evidence that EXP screening has
For item = 0 to UBound(PND1_array, 2)
	MAXIS_case_number = PND1_array(case_num, item)	'Case number for each loop from the array
	appl_date = PND1_array(app_date, item)			'appl date for each loop from the array
		
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen MAXIS_case_number, 18, 43
	Call navigate_to_MAXIS_screen("CASE", "NOTE")
	
	'Checking for PRIV cases
	EMReadScreen priv_check, 6, 24, 14 			'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN 				'Delete priv cases from excel sheet, save to a list for later
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
		PF3
		appears_exp = True
	ELse
		'starting at the 1st case note, checking the headers for the NOTES - EXPEDITED SCREENING text	
		MAXIS_row = 5
		Do 
			EMReadScreen case_note_date, 8, MAXIS_row, 6
			If case_note_date = "        " then 
				appears_exp = True
				exit do
			End if 
			If case_note_date => appl_date then 
				EMReadScreen case_note_header, 55, MAXIS_row, 25
				case_note_header = trim(case_note_header)	
				IF instr(case_note_header, "client appears expedited") then				
					appears_exp = True 
					exit do
				Elseif instr(case_note_header, "client does not appear expedited") then
					appears_exp = FALSE
					exit do
				Else 
					appears_exp = True			'defaults all other cases to true, to be addded to the list 
				END IF
			END IF 
			MAXIS_row = MAXIS_row + 1
		LOOP until case_note_date < appl_date
	END If 
				
	'if cases are pending for MFIP or SNAP and appear to be EXP based on not having a EXP screening, or EXP screening shows they appear exp, then the cases will be added to Excel.
	If appears_exp = True then 
		add_to_excel = True
	ELSEif appears_exp = False then 
		add_to_excel = False
	END IF 
NEXT		

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "Case information"
ObjExcel.ActiveSheet.Name = "PND1 cases"

'adding information to the Excel list
ObjExcel.Cells(1, 1).Value = "Worker"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Client name"
ObjExcel.Cells(1, 4).Value = "APPL date"
objExcel.Columns(4).NumberFormat = "mm/dd/yy"	'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 5).Value = "# day pending"
ObjExcel.Cells(1, 6).Value = "NOTES"

'formatting the cells 
FOR i = 1 to 6		
	objExcel.Cells(1, i).Font.Bold = True		'bold font
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT	

'Addded the potentially EXP SNAP cases to Excel starting at row 2
excel_row = 2

For item = 0 to UBound(PND1_array, 2)
	If add_to_excel = True then 
		objExcel.Cells(excel_row, 1).Value = PND1_array (work_num,   	item)	'Adding worker number
		objExcel.Cells(excel_row, 2).Value = PND1_array (case_num,	 	item)	'Adding case number
		objExcel.Cells(excel_row, 3).Value = PND1_array (clt_name, 	   	item)	'Addubg client name
		objExcel.Cells(excel_row, 4).Value = PND1_array (app_date, 	   	item)	'Adding application date
		objExcel.Cells(excel_row, 5).Value = PND1_array (days_pending, 	item)	'Adding number of days pending
		excel_row = excel_row + 1
	End If
Next

FOR i = 1 to 6		'formatting the cells
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT
	
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>PND2 information 
'Sets up the array to store all the information for each client'
Dim PND2_array ()
ReDim PND2_array (5, 0)
entry_record = 0

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "PND2")
	EMWriteScreen worker, 21, 13
	transmit
	
	'Skips workers with no info
	EMReadScreen has_content_check, 8, 7, 3
	If has_content_check <> "        " then
		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			Do
				EMReadScreen SNAP_pending_status, 1, MAXIS_row, 62
				IF SNAP_pending_status = "P" then add_to_PND2_array = True
				If SNAP_pending_status <> "P" then 
					EMReadScreen CASH_pending_status, 1, MAXIS_row, 54
					If CASH_pending_status = "P" then 
						add_to_PND2_array = True
					Else
						add_to_PND2_array = FALSE
					End if 
				END IF 
					
				EMReadScreen worker_basket, 7, 21, 13				
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5		 'Reading case number
				EMReadScreen appl_date, 8, MAXIS_row, 38		     'Reading application date
				appl_date = replace(appl_date, " ", "/")
				EMReadScreen nbr_days_pending, 4, MAXIS_row, 49		 'Reading nbr days pending
				EMReadScreen client_name, 22, MAXIS_row, 16			 'Reading client name
				
				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and (instr(all_case_numbers_array, MAXIS_case_number) <> 0 and client_name <> " ADDITIONAL APP       ") then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end
				
				'If additional application is rec'd then the excel output is the client's name, not ADDITIONAL APP
				if trim(client_name) = "ADDITIONAL APP" then
					EMReadScreen alt_client_name, 22, MAXIS_row - 1, 16
					client_name = trim(alt_client_name)
				END IF
				
				'Adding client information to the array'
				If add_to_PND2_array = True and MAXIS_case_number <> "" then 
					ReDim Preserve PND2_array(5, entry_record)	'This resizes the array based on the number of rows in the Excel File'
					'The client information is added to the array'
					PND2_array (work_num,     entry_record) = worker_basket
					PND2_array (case_num,	  entry_record) = MAXIS_case_number		
					PND2_array (clt_name,  	  entry_record) = client_name
					PND2_array (app_date, 	  entry_record) = appl_date
					PND2_array (days_pending, entry_record) = nbr_days_pending
					
					entry_record = entry_record + 1			'This increments to the next entry in the array'
					STATS_counter = STATS_counter + 1
				END IF
				MAXIS_row = MAXIS_row + 1	
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'Now the script goes into CASENOTE and searches for evidence that EXP screening has
For item = 0 to UBound(PND2_array, 2)
	MAXIS_case_number = PND2_array(case_num, item)	'Case number for each loop from the array
	appl_date = PND2_array(app_date, item)			'appl date for each loop from the array
	
	appears_exp = True
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen MAXIS_case_number, 18, 43
		
	'Checking for PRIV cases
	EMReadScreen priv_check, 6, 24, 14 			'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN 				'Delete priv cases from excel sheet, save to a list for later
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
		PF3
		appears_exp = True
	ELse
		'checking for ACTIVE SNAP 
		Call navigate_to_MAXIS_screen("STAT", "PROG")
		EMReadScreen SNAP_status, 4, 10, 74
		If SNAP_status = "ACTV" then 
			appears_exp = false
		Elseif SNAP_status <> "PEND" then 
			'Checking for ACTIVE MFIP
			MAXIS_row = 6
			Do 
				EMReadScreen cash_status, 2, MAXIS_row, 67
				If cash_status = "MF" then 
					EMReadScreen program_status, 4, MAXIS_row, 74
					If program_status <> "PEND" then 
						appears_exp = false
					elseif cash_status = "  " then 
						appears_exp = true
					END IF 
				END IF 
				MAXIS_row = MAXIS_row + 1
			LOOP until MAXIS_row = 	8
		END IF 
	
		'Because some cases don't have HCRE dates listed, so when you try to go past PROG the script gets caught up. Do...loop handles this instance.
		PF3		'exits PROG to prompt HCRE if HCRE isn't complete
		Do
			EMReadscreen HCRE_panel_check, 4, 2, 50
			If HCRE_panel_check = "HCRE" then
				PF10	'exists edit mode in cases where HCRE isn't complete for a member
				PF3
			END IF
		Loop until HCRE_panel_check <> "HCRE"		'repeats until case is not in the HCRE panel
		'starting at the 1st case note, checking the headers for the NOTES - EXPEDITED SCREENING text	
		MAXIS_row = 5
		Do 
			EMReadScreen case_note_date, 8, MAXIS_row, 6
			If case_note_date = "        " then 
				appears_exp = True
				exit do
			End if 
			
			If case_note_date => appl_date then 
				EMReadScreen case_note_header, 55, MAXIS_row, 25
				case_note_header = trim(case_note_header)	
				IF instr(case_note_header, "client appears expedited") then				
					appears_exp = True 
					exit do
				Elseif instr(case_note_header, "client does not appear expedited") then
					appears_exp = FALSE
					exit do
				Else 
					appears_exp = True			'defaults all other cases to true, to be addded to the list 
				END IF
			END IF
			MAXIS_row = MAXIS_row + 1
		LOOP until case_note_date < appl_date
	END if	
			
	'if cases are pending for MFIP or SNAP and appear to be EXP based on not having a EXP screening, or EXP screening shows they appear exp, then the cases will be added to Excel.
	If appears_exp = True then 
		add_to_excel = True
	ELSEif appears_exp = False then
		add_to_excel = False
	END IF  
NEXT	

'Adding another sheet 
ObjExcel.Worksheets.Add().Name = "PND2 cases"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Client name"
ObjExcel.Cells(1, 4).Value = "APPL date"
objExcel.Columns(4).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 5).Value = "# day pending"
ObjExcel.Cells(1, 6).Value = "NOTES"

FOR i = 1 to 6		'formatting the cells
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Addded the potentially EXP SNAP cases to 
excel_row = 2		'Setting the excel_row to start writing data on

For item = 0 to UBound(PND2_array, 2)
	If add_to_excel = True then 
		objExcel.Cells(excel_row, 1).Value = PND2_array (work_num,   	item)	'Adding worker number
		objExcel.Cells(excel_row, 2).Value = PND2_array (case_num,	 	item)	'Adding case number
		objExcel.Cells(excel_row, 3).Value = PND2_array (clt_name, 	   	item)	'Addubg client name
		objExcel.Cells(excel_row, 4).Value = PND2_array (app_date, 	   	item)	'Adding application date
		objExcel.Cells(excel_row, 5).Value = PND2_array (days_pending, 	item)	'Adding number of days pending
		excel_row = excel_row + 1
	End If
Next

FOR i = 1 to 6		'formatting the cells
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT	

'setting col to use to start writing run time information into to Excel
col_to_use = 8

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Please review the PND1 and PND2 lists for potential EXP SNAP processing.")