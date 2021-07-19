'Required for statistical purposes===============================================================================
name_of_script = "BULK - MONY VNDW.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 30                               'manual run time, per line, in seconds
STATS_denomination = "I"       'I is for each ITEM
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
call changelog_update("07/02/21", "Initial version.", "Dave Courtright, SLC")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog vndw_dialog, 0, 0, 181, 115, "MONY/VNDW"
  EditBox 10, 35, 165, 15, vendor_number_string
  ButtonGroup ButtonPressed
    OkButton 70, 95, 50, 15
    CancelButton 125, 95, 50, 15
  Text 50, 5, 90, 10, "MONY/VNDW to Excel"
  Text 10, 20, 170, 10, "Vendor number(s) to check, separated by commas:"
  Text 10, 55, 290, 10, "Date Range to Report:"
  EditBox 10, 70, 65, 15, begin_date
  EditBox 110, 70, 65, 15, end_date
  Text 90, 75, 15, 10, "to"
EndDialog

'Connects to MAXIS
EMConnect ""

'Looks up an existing user for autofilling the next dialog
CALL find_variable("User: ", x_number_editbox, 7)

'defaulting the script to check all DAILS on a DAIL list
all_check = 1

'Shows the dialog. Doesn't need to loop since we already looked at MAXIS.
DO
	DO
		err_msg = ""
		dialog vndw_dialog
		cancel_confirmation
		if vendor_number_string = "" then err_msg = err_msg & vbNewLine & "Enter at least one vendor number to check."
		if isdate(begin_date) = false or isdate(end_date) = false then err_msg = err_msg & vbNewLine & "Pleaese enter a valid date range to search."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
  Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'splits the results of the editbox into an array
vendor_number_array = split(vendor_number_string, ",")
begin_date = cdate(begin_date)
end_date = cdate(end_date)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "WARRANTS BY VENDOR"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "Vendor Number"
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value = "Warrant #"
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value = "Issue Date"
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value = "Transaction Number"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "Client Name"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "Amount"
objExcel.Cells(1, 6).Font.Bold = True
objExcel.Cells(1, 7).Value = "Program"
objExcel.Cells(1, 7).Font.Bold = True
objExcel.Cells(1, 8).Value = "Status"
objExcel.Cells(1, 8).Font.Bold = True
objExcel.Cells(1, 9).Value = "Ref Nbr"
objExcel.Cells(1, 9).Font.Bold = True


'Sets variable for all of the Excel stuff
excel_row = 2

'This for...next contains each vendor number to check
For each vendor_number in vendor_number_array

	'Trims the x_number so that we don't have glitches
	vendor_number = trim(vendor_number)

	back_to_SELF
	MAXIS_case_number = ""			'Blanking this out for PRIV case handling.
	CALL navigate_to_MAXIS_screen("MONY", "VNDW")
	EMWriteScreen vendor_number, 4, 11
	transmit
		'This loop will grab the information for this particular vendor and place on the sheet if within the date range.
	read_line = 7
	DO
		EMReadScreen iss_date, 8, read_line, 14 'the issue date determines what warrants we want

		If iss_date = "        "  Then Exit Do	'Jump out if we hit the end of the list or the list is empty

		IF cdate(iss_date) > begin_date AND cdate(iss_date) < end_date THEN
			excel_row = excel_row + 1
			EMReadScreen warrant, 8, read_line, 5
			EMReadScreen trans_number, 9, read_line, 23
			EMReadScreen client_name, 24, read_line, 33
			EMReadScreen amount, 7, read_line, 58
			EMReadScreen program, 3, read_line, 67
			EMReadScreen pay_status, 1, read_line, 71
			EMReadScreen ref_nbr, 8, read_line, 73
			objExcel.Cells(excel_row, 1).Value = vendor_number
			ObjExcel.Cells(excel_row, 2).Value = warrant
			ObjExcel.Cells(excel_row, 3).Value = iss_date
			ObjExcel.Cells(excel_row, 4).Value = trans_number
			ObjExcel.Cells(excel_row, 5).Value = client_name
			ObjExcel.Cells(excel_row, 6).Value = amount
			ObjExcel.Cells(excel_row, 7).Value = program
			ObjExcel.Cells(excel_row, 8).Value = pay_status
			ObjExcel.Cells(excel_row, 9).Value = ref_nbr
		END IF
		read_line = read_line + 1
		IF read_line = 19 Then
			PF8
			EMReadScreen last_page, 4, 24, 2
			IF last_page = "THIS" then exit do
			read_line = 7
		END IF
	LOOP UNTIL iss_date < begin_date or iss_date = "        "

Next


'Formatting the column width.
FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()
NEXT



STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!")
