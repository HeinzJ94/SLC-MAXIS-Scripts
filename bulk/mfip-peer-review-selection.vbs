'Required for statistical purposes===============================================================================
name_of_script = "BULK - TARGETED MFIP REVIEW SELECTION.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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

'Defining classes-----------------------------
Class case_attributes 'This class holds case-specific data
	public MAXIS_case_number
	public SNAP_status
	public worker_number
	public benefit_level
	public total_income
	public snap_grant
	public inactive_date
	public failure_reason
	public inactive_reason
END Class

case_percentage = "10" 'Setting the percent of cases to select to 10% by default, can be changed in dialog'

'DIALOGS----------------------------------------------------------------------
BeginDialog targeted_snap_review_dialog, 0, 0, 286, 170, "Targeted SNAP Review Selection"
  EditBox 150, 20, 130, 15, worker_number
  CheckBox 70, 90, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 10, 20, 40, 10, "PAR", Active_check
  CheckBox 10, 50, 40, 10, "CAPER", CAPER_check
  ButtonGroup ButtonPressed
    OkButton 175, 140, 50, 15
    CancelButton 230, 140, 50, 15
  GroupBox 5, 5, 60, 90, "Case Types"
  Text 70, 25, 65, 10, "Worker(s) to check:"
  Text 70, 115, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 80, 5, 125, 10, "Targeted SNAP Review Selection"
  Text 70, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 20, 65, 35, 15, "(Closed / Denied)"
  Text 20, 35, 35, 15, "(Active)"
  EditBox 230, 65, 20, 15, MAXIS_footer_month
  EditBox 260, 65, 20, 15, MAXIS_footer_year
  Text 70, 60, 150, 20, "Footer Month and Year of actions to review: (mm/yy)"
EndDialog

'New dialog for including ES reviews / selecting case numbers at initial dialog'
'BeginDialog targeted_snap_review_dialog, 0, 0, 286, 195, "Targeted SNAP Review Selection"
'  EditBox 150, 20, 130, 15, worker_number
'  CheckBox 70, 125, 150, 10, "Check here to run this query county-wide.", all_workers_check
'  CheckBox 10, 20, 40, 10, "PAR", Active_check
'  CheckBox 10, 50, 40, 10, "CAPER", CAPER_check
'  ButtonGroup ButtonPressed
'    OkButton 170, 170, 50, 15
'    CancelButton 225, 170, 50, 15
'  GroupBox 5, 5, 60, 90, "Case Types"
'  Text 70, 25, 65, 10, "Worker(s) to check:"
'  Text 70, 145, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
'  Text 75, 5, 125, 10, "Targeted MFIP Review Selection"
'  Text 70, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
'  Text 20, 65, 35, 15, "(Closed / Denied)"
'  Text 20, 35, 35, 15, "(Active)"
'  EditBox 230, 65, 20, 15, MAXIS_footer_month
'  EditBox 260, 65, 20, 15, MAXIS_footer_year
'  Text 70, 60, 150, 20, "Footer Month and Year of actions to review: (mm/yy)"
'  EditBox 130, 85, 25, 15, Active_Cases_To_Select
'  EditBox 230, 85, 30, 15, ES_Reviews_To_Select
'  Text 160, 85, 65, 15, "ES Reviews To Select:"
'  Text 70, 85, 50, 20, "MFIP Reviews To Select:"
'EndDialog


BeginDialog cases_to_select_dialog, 0, 0, 176, 125, "Cases to Select"
  ButtonGroup ButtonPressed
    OkButton 30, 105, 50, 15
    CancelButton 85, 105, 50, 15
  Text 15, 10, 160, 20, "Cases to audit based on the total number of cases meeting selection criteria:  "
  EditBox 85, 35, 20, 15, cases_to_select
  EditBox 85, 55, 20, 15, caper_cases_to_select
  Text 50, 35, 30, 15, "Active:"
  Text 45, 55, 35, 20, "CAPER (Inactive):"
  Text 15, 80, 150, 20, "Note: reducing these numbers will reduce the overall accuracy of your case audit."
EndDialog




'DECLARE VARIABLES

'THE SCRIPT-------------------------------------------------------------------------

'Determining specific county for multicounty agencies...
get_county_code
'two_digit_county_code = "36"
'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog targeted_snap_review_dialog
If buttonpressed = cancel then stopscript


'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_password(false)


'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(x1_number)		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Prepare the arrays and counters to begin case collection

	sa_count = 0
	DIM snap_active_array()

IF CAPER_check = checked then

dim caper_array()
	ca_count = 0
END IF
active_criteria_total = 0
caper_criteria_total = 0
excel_row = 2


'First, we check REPT/ACTV.  Must be done on ACTIVE and CAPER checks'
For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "actv")
	EMWriteScreen worker, 21, 13
	transmit
	EMReadScreen user_worker, 7, 21, 71		'
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7

			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12		'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 21		'Reading client name
				EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
				EMReadScreen cash_status, 9, MAXIS_row, 51		'Reading SNAP status

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)

				If instr(cash_status, "MF A") > 0 or instr(cash_status, "DW A") > 0 and Active_check = checked then
					redim preserve SNAP_active_array(sa_count)
					set SNAP_active_array(sa_count) = new case_attributes
					SNAP_active_array(sa_count).MAXIS_case_number = MAXIS_case_number
					IF instr(cash_status, "MF A") > 0 THEN SNAP_active_array(sa_count).SNAP_status = "MF"
					IF instr(cash_status, "DW A") > 0 THEN SNAP_active_array(sa_count).SNAP_status = "DW"
					SNAP_active_array(sa_count).worker_number = worker
					sa_count = sa_count+1
				END IF

				If instr(cash_status, "MF I") = true and CAPER_check = checked then
					redim preserve caper_array(ca_count)
						set caper_array(ca_count) = new case_attributes
					caper_array(ca_count).MAXIS_case_number = MAXIS_case_number
					caper_array(ca_count).SNAP_status = SNAP_status
					caper_array(ca_count).worker_number = worker
					ca_count = ca_count + 1
				END IF

				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	'Now check REPT/INAC (caper only)
	IF CAPER_check = checked Then

		back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
		Call navigate_to_MAXIS_screen("rept", "inac")
		EMWriteScreen worker, 21, 16
		transmit
				'Skips workers with no info
		EMReadScreen has_content_check, 1, 7, 8
		If has_content_check <> " " then
			'EMReadScreen user_worker, 7, 21, 71		'
			'EMReadScreen p_worker, 7, 21, 13
			'IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV
		DO
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3		'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 14		'Reading client name

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end
				redim preserve caper_array(ca_count)
				set caper_array(ca_count) = new case_attributes
				caper_array(ca_count).MAXIS_case_number = MAXIS_case_number
				caper_array(ca_count).SNAP_status = SNAP_status
				caper_array(ca_count).worker_number = worker
				ca_count = ca_count + 1
				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
		END IF
	END IF
next


'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True
Set objWorkbook = objExcel.ActiveWorkbook

'--------Collecting CAPER CASES--------'
'First, check caper cases
IF CAPER_check = checked THEN
'Add a worksheet for CAPER - denials, label the columns'
ObjExcel.Worksheets.Add().Name = "denials"
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "INACTIVE REASON"
ObjExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "CLOSURE TYPE"
ObjExcel.Cells(1, 4).Font.Bold = TRUE
'Add a worksheet for CAPER - closures, label the columns'
ObjExcel.Worksheets.Add().Name = "closures"
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "INACTIVE REASON"
ObjExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "CLOSURE TYPE"
ObjExcel.Cells(1, 4).Font.Bold = TRUE





excel_row = 2
denial_row = 2
closure_row = 2
caper_denial_total = 0
caper_closure_total = 0

if isarray(caper_array) = true THEN
For c = 0 to ubound(caper_array)
	MAXIS_case_number = caper_array(c).MAXIS_case_number
	'Make sure in correct footer month, sometimes we drop back a month
	MAXIS_footer_month = datepart("m", date)
	IF len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
	MAXIS_footer_year = right(datepart("YYYY", date), 2)
	call navigate_to_MAXIS_screen("CASE", "CURR") 'Case/curr first, to find inactive date and reason
	EMReadScreen priv_check, 10, 24, 14
	IF priv_check = "PRIVILEGED" THEN 'Check for privileged cases that cannot be accessed'
		EMWriteScreen "____", 21, 70 'We need to blank out the command line to prevent an error on the next case'
	ELSE 'not privileged, go ahead and enter the case'

		EMWriteScreen "x", 4, 9

		Transmit
		'Check MFIP here'
		EMWriteScreen "MF", 3, 19
		Transmit
		row = 7
		col = 1
		EMSearch "INACTIVE", row, col
		IF row <> 0 THEN ' sometimes cases show up here with active status due to expedited
			EMReadScreen inactive_date, 8, row, 18
			EMReadScreen inactive_reason, 6, row, 60
			caper_array(c).inactive_date = inactive_date
			caper_array(c).inactive_reason = "" 'Default blank for now
			IF inactive_reason = "DENIED" Then caper_array(c).inactive_reason = "denial"
			IF inactive_reason = "INELIG" or inactive_reason = "SUSPEN" THEN caper_array(c).inactive_reason = "CLOSED"
			IF inactive_reason = "NO HRF" THEN caper_array(c).inactive_reason = "no hrf"
			If datediff("m", caper_array(c).inactive_date, date) <= 1 Then
		 		IF caper_array(c).inactive_reason = "CLOSED" OR inactive_reason = "denial" Then

		'	'IF caper_array(c).failure_reason = "Verification" or caper_array(c).failure_reason = "PACT" THEN 'Add cases that meet criteria to excel'
						IF caper_array(c).inactive_reason = "denial" THEN
							ObjExcel.Worksheets("denials").cells(denial_row, 1).value = caper_array(c).worker_number
							ObjExcel.Worksheets("denials").cells(denial_row, 2).value = caper_array(c).MAXIS_case_number
							ObjExcel.Worksheets("denials").cells(denial_row, 3).value = caper_array(c).failure_reason
							ObjExcel.Worksheets("denials").cells(denial_row, 4).value = caper_array(c).inactive_reason
		  				denial_row = denial_row + 1
		  				caper_denial_total = caper_denial_total + 1
						END IF
						IF caper_array(c).inactive_reason = "CLOSED" THEN
							ObjExcel.Worksheets("closures").cells(closure_row, 1).value = caper_array(c).worker_number
							ObjExcel.Worksheets("closures").cells(closure_row, 2).value = caper_array(c).MAXIS_case_number
							ObjExcel.Worksheets("closures").cells(closure_row, 3).value = caper_array(c).failure_reason
							ObjExcel.Worksheets("closures").cells(closure_row, 4).value = caper_array(c).inactive_reason
							closure_row = closure_row + 1
							caper_closure_total = caper_closure_total + 1
						END IF
				END IF
			END IF
		END IF
	END IF
Next
END IF
END IF
sa_count = 0
'Now it steps through each case in the array and determines whether to add it to the spreadsheet
IF Active_check = checked THEN
'Add a worksheet for ACTIVE cases, label the columns'
ObjExcel.Worksheets.Add().Name = "active cases"
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE

excel_row = 2
	For n = 0 to ubound(SNAP_active_array)
		'Make sure in correct footer month, sometimes we drop back a month
		MAXIS_footer_month = datepart("m", date)
		IF len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
		MAXIS_footer_year = right(datepart("YYYY", date), 2)
		MAXIS_case_number = SNAP_active_array(n).MAXIS_case_number
		call navigate_to_MAXIS_screen ("ELIG", snap_active_array(n).SNAP_status)
		EMReadScreen priv_check, 10, 24, 14
		IF priv_check = "PRIVILEGED" THEN 'Check for privileged cases that cannot be accessed'
		EMWriteScreen "____", 21, 70 'We need to blank out the command line to prevent an error on the next case'
		ELSE 'not privileged, go ahead and enter the case'

		EMReadScreen version, 3, 2, 17 'Finding most recent approved version
		version = cint(version)

		For approved = version to 0 Step -1
			EMReadScreen approved_check, 8, 3, 3
			If approved_check = "APPROVED" then Exit FOR
			version = version -1
			EMWriteScreen version, 20, 79
			transmit
		Next
		EMReadScreen approval_date, 8, 3, 14
		If datepart("m", approval_date) = datepart("m", (dateadd("m", -1, date))) THEN 'If this was approved in current month minus one, add to spreadsheet


				objExcel.cells(excel_row, 1).value = snap_active_array(n).worker_number
				objExcel.cells(excel_row, 2).value = MAXIS_case_number
				objExcel.cells(excel_row, 3).value = snap_active_array(n).total_income
				objExcel.cells(excel_row, 4).value = snap_active_array(n).snap_grant
				active_criteria_total = active_criteria_total + 1
				excel_row = excel_row + 1
			END IF
			sa_count = sa_count + 1
		END IF
	NEXT
END IF


col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

'add a sheet for audit cases and Stats
ObjExcel.Worksheets.Add().Name = "audit cases"
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 3).Value = "Auditor                "
objexcel.cells(1, 3).Font.Bold = TRUE
objexcel.cells(1, 4).Value = "Case Outcome"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.cells(1, 5).Value = "Notes on errors                                                                                   "
objExcel.Cells(1, 5).Font.Bold = TRUE


'=====THIS SECTION SELECTS RANDOM CASES FOR AUDIT==========='
'Determining how many cases to select for AUDIT
'----These inequalities create a statistically significant value based on the number of cases meeting the selections criteria.
'----The numbers selected give a lower than standard confidence in the sample in order to keep the total number of audits
'----feasible within time constraints.
IF active_check = checked THEN
	IF active_criteria_total < 2 THEN cases_to_select = active_criteria_total
	if active_criteria_total < 20 THEN cases_to_select = 2
	if active_criteria_total >= 20 AND active_criteria_total < 70 THEN cases_to_select = 3
	if active_criteria_total >= 70 AND active_criteria_total < 200 THEN cases_to_select = 4
	if active_criteria_total >= 200 AND active_criteria_total < 450 THEN cases_to_select = 5
	if active_criteria_total >= 450 THEN cases_to_select = 6 'This is open ended, as larger sample sizes see little change in results from increased sampling'
	IF all_workers_check = checked then cases_to_select = 10 ' select 10 for county-wide
	cases_to_select = cstr(cases_to_select) 'change to a string so it displays in dialog
END IF

IF caper_check = checked Then
	caper_criteria_total = caper_denial_total + caper_closure_total
	IF caper_criteria_total < 2 THEN caper_cases_to_select = caper_criteria_total
	If caper_criteria_total >= 2 AND caper_criteria_total < 20 THEN caper_cases_to_select = 2
	If caper_criteria_total >= 20 AND caper_criteria_total < 70 THEN caper_cases_to_select = 3
	If caper_criteria_total >= 70 AND caper_criteria_total < 200 THEN caper_cases_to_select = 4
	If caper_criteria_total >= 200 AND caper_criteria_total < 450 THEN caper_cases_to_select = 5
	If caper_criteria_total >= 450 THEN caper_cases_to_select = 6
	IF all_workers_check = checked then caper_cases_to_select = 10 ' Select 10 cases for each county
	caper_cases_to_select = cstr(caper_cases_to_select) 'change to a string so it displays in dialog
END IF

Dialog cases_to_select_dialog
IF buttonpressed = cancel then stopscript

audit_row = 2 'reset the row for the audit sheet
'Selecting random cases and pasting into the new worksheet
IF active_check = checked THEN
	IF active_criteria_total > 0 THEN
		objWorkbook.Worksheets("audit cases").cells(audit_row, 1).Value = "ACTIVE / PAR CASES"
		objWorkbook.Worksheets("audit cases").cells(audit_row, 1).Font.Bold = true
		audit_row = audit_row + 1
		'Make sure we don't try to sample less than all cases
		IF cint(cases_to_select) >= active_criteria_total THEN
		'Here we copy / paste the whole list
			objWorkbook.worksheets("active cases").Range("A2:B" & active_criteria_total + 1).copy
			objWorkbook.worksheets("audit cases").Range("A3").PasteSpecial
			audit_row = audit_row + active_criteria_total
		ELSE'We need a random selection of cases
			Set active_selection_list = CreateObject("Scripting.dictionary") 'create a dictionary object to prevent duplicating cases'
			active_selection_list(1) = 0 'entering row 1, so it is consistently there for future use.  We never have a case on row 1, we will be able to ignore'
			DO
				Randomize
				row_to_select = Int(active_criteria_total*Rnd)
				active_selection_list(row_to_select) = 0 '0 is just placeholder, only using keys
			LOOP UNTIL active_selection_list.count = cases_to_select + 1 'plus 1 to account for row 1 always there
			For each select_this_case in active_selection_list.keys
				IF select_this_case <> 1 THEN 'ignore row 1
					select_this_case = "A" & select_this_case & ":B" & select_this_case
					objWorkbook.worksheets("active cases").Range(select_this_case).copy
					objWorkbook.worksheets("audit cases").Range("A" & audit_row).PasteSpecial
					audit_row = audit_row + 1
				END IF
			Next
		END IF
		audit_row = audit_row + 1 'adding an extra row to separate case types
	END IF
END IF

'Selecting random caper cases and pasting into the new worksheet
If caper_check = checked THEN
'Determing totals of denials / closures, attempt to create a 50/50 ratio
	IF isnumeric(caper_cases_to_select) = true and caper_cases_to_select > 0 THEN
		denials_to_select = cint(caper_cases_to_select / 2) 'divide total by two, and round to integer
		closures_to_select = caper_cases_to_select - denials_to_select 'subtract from total to account for the rounding
		'THese conditionals reapportion the totals for all possible scenarios to prevent selecting more than total cases'
		IF caper_denial_total < denials_to_select AND caper_closure_total >= (caper_cases_to_select - caper_denial_total) THEN
			 	denials_to_select = caper_denial_total 'make sure we don't select more than we have
			closures_to_select = caper_cases_to_select - caper_denial_total 'reset the other value to keep the total the same
		END IF
		IF caper_denial_total < denials_to_select AND caper_closure_total < (caper_cases_to_select - caper_denial_total) Then
			denials_to_select = caper_denial_total
			closures_to_select = caper_closure_total
		END IF
		IF caper_closure_total < closures_to_select AND caper_denial_total >= (caper_cases_to_select - caper_closure_total) THEN
			closures_to_select = caper_closure_total
			denials_to_select = caper_cases_to_select - closures_to_select
		END IF
	END IF
	'Here, handle the denial sheet

	objWorkbook.Worksheets("audit cases").cells(audit_row, 1).Value = "CAPER CASES"
	objWorkbook.Worksheets("audit cases").cells(audit_row, 1).Font.Bold = true
	audit_row = audit_row + 1
	'Make sure we don't try to sample less than all cases
	IF caper_denial_total > 0 THEN
		IF denials_to_select >= caper_denial_total THEN
		'Here we copy / paste the whole list
			objWorkbook.worksheets("denials").Range("A2:B" & caper_denial_total + 1).copy
			objWorkbook.worksheets("audit cases").Range("A" & audit_row).PasteSpecial
			audit_row = audit_row + caper_denial_total
		ELSE'We need a random selection of cases
			Set denial_selection_list = CreateObject("Scripting.dictionary") 'create a dictionary object to prevent duplicating cases'
			denial_selection_list(1) = 0 'entering row 1, so it is consistently there for future use.  We never have a case on row 1, we will be able to ignore'
			DO
				Randomize
				row_to_select = Int(caper_denial_total*Rnd)
				denial_selection_list(row_to_select) = 0 '0 is just placeholder, only using keys
			LOOP UNTIL denial_selection_list.count = denials_to_select + 1 'plus 1 to account for row 1 always there
			For each select_this_case in denial_selection_list.keys
				IF select_this_case <> 1 THEN 'ignore row 1
					select_this_case = "A" & select_this_case & ":B" & select_this_case
					objWorkbook.worksheets("denials").Range(select_this_case).copy
					objWorkbook.worksheets("audit cases").Range("A" & audit_row & ":B" & audit_row).PasteSpecial
					audit_row = audit_row + 1
				END IF
			Next
		END IF
	END IF
	IF caper_closure_total > 0 THEN
	IF closures_to_select >= caper_closure_total THEN
	'Here we copy / paste the whole list
		objWorkbook.worksheets("closures").Range("A2:B" & caper_closure_total + 1).copy
		objWorkbook.worksheets("audit cases").Range("A" & audit_row).PasteSpecial
		audit_row = audit_row + caper_closure_total
	ELSE'We need a random selection of cases
		Set closure_selection_list = CreateObject("Scripting.dictionary") 'create a dictionary object to prevent duplicating cases'
		closure_selection_list(1) = 0
		DO
			Randomize
			row_to_select = Int(caper_closure_total*Rnd) + 1 'plus one, as we start counting at row 2,
			closure_selection_list(row_to_select) = 0 '0 is just placeholder, only using keys
		LOOP UNTIL closure_selection_list.count = closures_to_select + 1 'plus one because we skip row 1'
		For each select_this_case in closure_selection_list.keys
			IF select_this_case <> 1 THEN
				select_this_case = "A" & select_this_case & ":B" & select_this_case
				objWorkbook.worksheets("closures").Range(select_this_case).copy
				objWorkbook.worksheets("audit cases").Range("A" & audit_row & ":B" & audit_row).PasteSpecial
				audit_row = audit_row + 1
			END IF
		Next
	END IF
	END IF
END IF
'Put in some shit here to pull ES cases.  Then edit this comment.'
'IF isnumeric(ES_Reviews_To_Select) = true Then
 	'ES cases stuff goes here.  Should pull post 60 cases only
'END IF

'Query stats
stats_row = 3
objExcel.Cells(1, 6).Font.Bold = TRUE
objExcel.Cells(2, 6).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 7).Value = now
ObjExcel.Cells(2, 6).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 7).Value = timer - query_start_time
IF active_check = checked THEN
	ObjExcel.Cells(3, 6).Value = "Total active cases sampled:"
	ObjExcel.Cells(3, 7).Value = sa_count
	ObjExcel.Cells(4, 6).Value = "Percent of cases meeting criteria:"
	ObjExcel.Cells(4, 7).NumberFormat = "0.00%"
	if sa_count <> 0 THEN ObjExcel.Cells(4, 7).Value = active_criteria_total / sa_count
	stats_row = 5
END IF
IF caper_check = checked then
	ObjExcel.Cells(stats_row, 6).Value = "Total CAPER cases sampled:"
	ObjExcel.Cells(stats_row, 7).Value = ca_count
	ObjExcel.Cells(stats_row + 1, 6).Value = "Percent of cases meeting criteria:"
	ObjExcel.Cells(stats_row + 1, 7).NumberFormat = "0.00%"
	if ca_count <> 0 THEN ObjExcel.Cells(stats_row + 1, 7).Value = (caper_closure_total + caper_denial_total) / ca_count
END IF
'Formatting dropdowns for the outcome fields
'First create a hidden list of values
	ObjExcel.Cells(1, 16).Value = "Technical"
	ObjExcel.Cells(2, 16).Value = "Eligibility"
	ObjExcel.Cells(3, 16).Value = "Correct"
	ObjExcel.Cells(1, 16).entireColumn.hidden = true

For row_to_format = 2 to audit_row
	with objExcel.cells(row_to_format, 4).Validation
			.Add 3, 1, 1, "=$P$1:$P$3"
			.IgnoreBlank = True
			.InCellDropdown = True
			.InputTitle = ""
			.ErrorTitle = ""
			.InputMessage = ""
			.ErrorMessage = ""
			.ShowInput = True
			.ShowError = True
		end With
Next
'Autofitting columns
For col_to_autofit = 1 to 7
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("")
