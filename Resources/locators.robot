
*** Variables ***


${Not_downlaod}    image:${CURDIR}${/}..${/}data${/}images${/}Not_downlaoded.png
# ${Execute_not_download}    image:${CURDIR}${/}..${/}data${/}images${/}Execute_not_download.png
${No_Value_found}  image:${CURDIR}${/}..${/}data${/}images${/}No_value_found.png
# ${Change_layout_execute}  image:${CURDIR}${/}..${/}data${/}images${/}chnage_layout_execute.png
${Payment_list_found}  image:${CURDIR}${/}..${/}data${/}images${/}Payment_list_found.png
${Payment_list_Disabled}  image:${CURDIR}${/}..${/}data${/}images${/}Payment_list_Disabled.png
${Auth_error}  image:${CURDIR}${/}..${/}data${/}images${/}Auth_error.png
# ${no_data_found}  image:${CURDIR}${/}..${/}data${/}images${/}not_data_found_htm.png


# Locators VBS
${MATHCODE_BUTTON}     wnd[0]/usr/ctxtF110V-LAUFI
${VPMP_FILLING}     wnd[1]/usr/txtF110V-LAUFI
${No_Of_Hits}     wnd[1]/usr/txtMAX_REC
${Selecting_row}     wnd[1]/usr/lbl[1,3]
# ${Payment_list}    image:${CURDIR}${/}..${/}data${/}images${/}Payment_list.png
${Local_file_f110}    image:${CURDIR}${/}..${/}data${/}images${/}Local_file_f110.png
${LOC_USERNAME}   wnd[0]/usr/txtRSYST-BNAME
${LOC_PASSWORD}   wnd[0]/usr/pwdRSYST-BCODE
# ${Transaction}      /XFT/INV_MONITOR
# Transaction Screen
${LOC_CONTAINER}  wnd[0]/usr/cntlCONTAINER/shellcont/shell
${LOC_SHELL_0}    ${LOC_CONTAINER}/shellcont[0]/shell
${LOC_SHELL_1}    ${LOC_CONTAINER}/shellcont[1]/shell

${LOC_VAR_WINDOW}  wnd[1]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell
${LOC_VAR_FILTER}  wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW


# Upload Barcode_process.robot
${F1_document_header}     image:${CURDIR}${/}..${/}data${/}images${/}F1_document_number.png
${IMG_SETFILTER}  image:${CURDIR}${/}..${/}data${/}images${/}setfilter.png
${F1_doc_value}  image:${CURDIR}${/}..${/}data${/}images${/}F1_doc_field.png
${Current_status_header}  image:${CURDIR}${/}..${/}data${/}images${/}Current_status_header.png
${Current_status_value}  image:${CURDIR}${/}..${/}data${/}images${/}Currentstatus.png
${Doc_creation}  image:${CURDIR}${/}..${/}data${/}images${/}docs.png
${Doc_File_system}  image:${CURDIR}${/}..${/}data${/}images${/}documents.png
${Export_doc}  image:${CURDIR}${/}..${/}data${/}images${/}Saving_pdf.png
${pdf_export}  image:${CURDIR}${/}..${/}data${/}images${/}pdf_file_export.png
${count_helper}  image:${CURDIR}${/}..${/}data${/}images${/}Count_helper.png
${cross_cut}  image:${CURDIR}${/}..${/}data${/}images${/}cross_cut.png
${Downloaded_file}  image:${CURDIR}${/}..${/}data${/}images${/}file_download_check.png
${execute_button}  image:${CURDIR}${/}..${/}data${/}images${/}execute.png
${hit_filter_zero}  image:${CURDIR}${/}..${/}data${/}images${/}hit_filter_zero.png
${SAP_TMP}
${Clear_filter}     image:${CURDIR}${/}..${/}${/}data${/}images${/}Clear_filter.png
${Delete_filter}     image:${CURDIR}${/}..${/}${/}data${/}images${/}Delete_filter.png
${Not_F1_Doc}  image:${CURDIR}${/}..${/}${/}data${/}images${/}Not_F1_doc.png
@{BARCODE_RECORDS}

# Locator of Upload BArcode.robot
# === IMAGE LOCATORS ===
${IMG_INVOICE}    image:${CURDIR}${/}..${/}data${/}images${/}invoice.png
${IMG_BARCODE}    image:${CURDIR}${/}..${/}data${/}images${/}barcode.png
${IMG_LAYOUT}     image:${CURDIR}${/}..${/}data${/}images${/}layout.png


${Multiple_Selection_Image}  image:${CURDIR}${/}..${/}data${/}images${/}Invoice_multiple_selection.png
${Multiple_Selection_Image_safe}  image:${CURDIR}${/}..${/}data${/}images${/}Invoice_multiple_selection_safe.png
${Perform_Search_Image}  image:${CURDIR}${/}..${/}data${/}images${/}Perform_search.png
${Hit_filter_Image}     image:${CURDIR}${/}..${/}data${/}images${/}Hit_filter_check.png
${default_setting_click}    image:${CURDIR}${/}..${/}data${/}images${/}default_setting_select.png
${LOC_Choose_LAYOUT_SHELL}  image:${CURDIR}${/}..${/}data${/}images${/}Choose_layout.png
${Layout_User_Active}    image:${CURDIR}${/}..${/}data${/}images${/}Layout_User_Active.png

# Barcode Upload
${LOC_PATH_INPUT}  wnd[2]/usr/ctxtDY_PATH
${LOC_FILE_INPUT}  wnd[2]/usr/ctxtDY_FILENAME
${LOC_UPLOAD_BTN}  wnd[2]/tbar[0]/btn[0]
${LOC_LANGUAGE}   wnd[0]/usr/txtRSYST-LANGU
${LOC_CLIENT}     wnd[0]/usr/txtRSYST-MANDT