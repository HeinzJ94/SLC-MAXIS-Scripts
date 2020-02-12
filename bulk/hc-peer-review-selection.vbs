'Required for statistical purposes===============================================================================
name_of_script = "BULK - TARGETED SNAP REVIEW SELECTION.vbs"
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
	public special_case_type
END Class

case_percentage = "10" 'Setting the percent of cases to select to 10% by default, can be changed in dialog'

'DIALOGS----------------------------------------------------------------------
BeginDialog dialog1, 0, 0, 236, 185, "Targeted HC Review Selection"
  EditBox 85, 20, 130, 15, worker_number
  CheckBox 10, 60, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 105, 155, 50, 15
    CancelButton 170, 155, 50, 15
  Text 10, 25, 65, 10, "Worker(s) to check:"
  Text 10, 75, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 80, 5, 125, 10, "Targeted HC Review Selection"
  Text 10, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 10, 100, 45, 10, "Date Range:"
  EditBox 65, 95, 55, 15, begin_date
  EditBox 150, 95, 65, 15, end_date
  Text 125, 100, 15, 10, "to"
EndDialog


'DECLARE VARIABLES

'THE SCRIPT-------------------------------------------------------------------------

'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog
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
	DIM HC_active_array()
active_criteria_total = 0

excel_row = 2


'First, we check REPT/ACTV.  Must be done on ACTIVE and CAPER checks'
For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "actv")
	if worker = "X169SHC" then msgbox "Here it is"
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
				EMReadScreen HC_status, 1, MAXIS_row, 64		'Reading SNAP status


				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)

				If HC_status = "A" then
					redim preserve HC_active_array(sa_count)
					set HC_active_array(sa_count) = new case_attributes
					HC_active_array(sa_count).MAXIS_case_number = MAXIS_case_number
				''	msgbox sa_count & " " & SNAP_active_array(sa_count).MAXIS_case_number & " " & ubound(SNAP_active_array)
					HC_active_array(sa_count).SNAP_status = SNAP_status
					HC_active_array(sa_count).worker_number = worker
					sa_count = sa_count + 1
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

next


'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True
Set objWorkbook = objExcel.ActiveWorkbook



excel_row = 2

sa_count = 0
'Now it steps through each case in the array and determines whether to add it to the spreadsheet

'Add a worksheet for ACTIVE cases, label the columns'
ObjExcel.Worksheets.Add().Name = "active cases"
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
excel_row = 2
	For n = 0 to ubound(HC_active_array) 'loop through every active HC case
		'Make sure in correct footer month, sometimes we drop back a month
		MAXIS_footer_month = datepart("m", date)
		IF len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
		MAXIS_footer_year = right(datepart("YYYY", date), 2)
		MAXIS_case_number = HC_active_array(n).MAXIS_case_number
		call navigate_to_MAXIS_screen ("ELIG", "HC")
		'This section reads through all rows of HHMM to check for programs with asset test'
		asset_test = false 'reset variable
		For member_row = 8 to 19
			EMReadScreen total_versions, 1, member_row, 65
		''	msgbox member_row
			IF isnumeric(total_versions) = true THEN
				total_versions = right(total_versions, 1)
				for version = total_versions to 1
					If len(version) = 1 then version = "0" & version
					EMReadScreen approved_status, 3, member_row, 68
					IF approved_status = "APP" THEN exit for
					EMWriteScreen version, member_row, 58
					transmit
				NEXT
				EMReadScreen program, 2, member_row, 28
				If program = "MA" THEN 'We need to go into the span to check ELIG type for MA'
					IF approved_status = "APP" THEN
						EMWriteScreen "x", member_row, 26
						transmit
						EMReadScreen autoclose_check, 4, 8, 39
						IF autoclose_check <> "Auto" THEN
							EMReadScreen process_date, 8, 2, 73
							IF begin_date <= process_date AND end_date >= process_date THEN
								EMReadScreen method, 1, 13, 21
								IF method = "B" or method = "S" then asset_test = true
								IF method = "L" then
									EMReadScreen ELIG_type, 2, 12, 17
									IF ELIG_type <> "AX" and ELIG_type <> "AA" THEN asset_test = true
								END IF
								IF method = "L" then HC_active_array(n).special_case_type = "LTC"
								IF method = "S" then HC_active_array(n).special_case_type = "EW"
								EMReadScreen waiver_type, 1, 14, 21
								IF  waiver_type = "J" or waiver_type = "K" THEN HC_active_array(n).special_case_type = "EW"
								IF waiver_type = "R" or waiver_type = "S" THEN HC_active_array(n).special_case_type = "DD"
								IF waiver_type = "F" or waiver_type = "G" THEN HC_active_array(n).special_case_type = "CADI"
								'Check spendown'
							END If
						END IF
						PF3
					END If
				END IF
				IF asset_test = false and approved_status = "APP" then 'We need to check process_date only if we haven't already determined from MA
					IF program = "QM" or program = "SL" or program = "DQ" or program = "QI" THEN
						EMWriteScreen "x", member_row, 26
						transmit
						EMReadScreen autoclose_check, 4, 8, 39
						IF autoclose_check <> "Auto" THEN
							EMReadScreen process_date, 8, 2, 73
							IF begin_date <= process_date AND end_date >= process_date THEN
								transmit
								transmit
								'read the next renewal type, we don't care about 6 month renewals
								EMReadScreen next_renewal_type, 7, 13, 3
								IF next_renewal_type <> "6 Month" THEN asset_test = true
								EMReadScreen twelve_month_date, 8, 11, 34
								twelve_month_date = replace(twelve_month_date, " ", "/")
								IF datediff("m", date, twelve_month_date) > 9 THEN	asset_test = true
								'msgbox datediff("m", date, twelve_month_date)
							END IF
						END IF
						PF3
						process_date = ""
					END If
				END IF
				IF asset_test = true then exit for
			END IF
		NEXT 'move to next line of elig screen'
		IF asset_test = true THEN 'add cases with asset_test to the worksheet'
			objexcel.cells(excel_row, 1).value = HC_active_array(n).worker_number
			objexcel.cells(excel_row, 2).value = HC_active_array(n).MAXIS_case_number
			objexcel.cells(excel_row, 3).value = HC_active_array(n).special_case_type
			excel_row = excel_row + 1
			criteria_count = criteria_count + 1
		END IF
NEXT 'move to next case in array'

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

'add a sheet for audit cases and Stats
ObjExcel.Worksheets.Add().Name = "audit cases"
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 3).Value = "Auditor                         "
objexcel.cells(1, 3).Font.Bold = TRUE
objexcel.cells(1, 4).Value = "Case Outcome"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.cells(1, 5).Value = "Notes on errors                                                                                             "
objExcel.Cells(1, 5).Font.Bold = TRUE

'=====THIS SECTION SELECTS RANDOM CASES FOR AUDIT==========='
cases_to_select = 100



audit_row = 2 'reset the row for the audit sheet
'Selecting random cases and pasting into the new worksheet

objWorkbook.Worksheets("audit cases").cells(audit_row, 1).Value = "ACTIVE / PAR CASES"
objWorkbook.Worksheets("audit cases").cells(audit_row, 1).Font.Bold = true
audit_row = audit_row + 1
'Make sure we don't try to sample less than all cases
	IF cint(cases_to_select) > criteria_count THEN
	'Here we copy / paste the whole list
		objWorkbook.worksheets("active cases").Range("A2:B" & criteria_count + 1).copy
		objWorkbook.worksheets("audit cases").Range("A2").PasteSpecial
		audit_row = audit_row + criteria_count
	ELSE'We need a random selection of cases
	Set active_selection_list = CreateObject("Scripting.dictionary") 'create a dictionary object to prevent duplicating cases'
	active_selection_list(1) = 0 'entering row 1, so it is consistently there for future use.  We never have a case on row 1, we will be able to ignore'
	DO
		Randomize
		row_to_select = Int(criteria_count*Rnd)
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
			denials_to_select = caper_denial_total AND closures_to_select = caper_closure_total
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

'Query stats
stats_row = 3
objExcel.Cells(1, 10).Font.Bold = TRUE
objExcel.Cells(2, 10).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 11).Value = now
ObjExcel.Cells(2, 10).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 11).Value = timer - query_start_time
IF active_check = checked THEN
	ObjExcel.Cells(3, 10).Value = "Total active cases sampled:"
	ObjExcel.Cells(3, 11).Value = sa_count
	ObjExcel.Cells(4, 10).Value = "Percent of cases meeting criteria:"
	ObjExcel.Cells(4, 11).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 11).Value = criteria_count / sa_count
	stats_row = 5
END IF
IF caper_check = checked then
	ObjExcel.Cells(stats_row, 10).Value = "Total CAPER cases sampled:"
	ObjExcel.Cells(stats_row, 11).Value = ca_count
	ObjExcel.Cells(stats_row + 1, 10).Value = "Percent of cases meeting criteria:"
	ObjExcel.Cells(stats_row + 1, 11).NumberFormat = "0.00%"
	ObjExcel.Cells(stats_row + 1, 11).Value = (caper_closure_total + caper_denial_total) / ca_count
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
