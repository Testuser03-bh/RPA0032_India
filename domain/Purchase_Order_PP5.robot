*** Settings ***
Documentation     SAP Operations Domain Logic
Library           Process
Library           SapGuiLibrary
Library           RPA.Desktop
Library           OperatingSystem
Library           String
Library           Collections
Library           DateTime
Library           ../adapters/db_adapter.py
Library            ..${/}adapters${/}Library${/}InitAllSettingsSQL.py
Resource     ..${/}Resources${/}locators.robot
Library    RPA.Email.ImapSmtp
    

*** Variables ***
${User_name}
${Pdf_number}      1
${primary_config}

*** Keywords ***
Main Workflow RPA0032 PP5
    [Arguments]    ${DOWNLOADED_HTML_FILES}
    Initialize SAP System PP5
    Login To SAP Application PP5
    Load Transaction Screen F110 PP5
    Download All HTML Files PP5    ${DOWNLOADED_HTML_FILES}
    Log To Console With Timestamp    Downlaoding All HTM file From PP5 
    FOR    ${file}    IN    @{DOWNLOADED_HTML_FILES}
        Log To Console With Timestamp     PP5- Total Files:- ${file}
    END
    FOR    ${HTML_FILE}    IN    @{DOWNLOADED_HTML_FILES}
        ${proc_status}    ${proc_err}=    Run Keyword And Ignore Error
        ...    Process HTML File    ${HTML_FILE}
        IF    '${proc_status}' == 'PASS'
            # Process HTML File returns two values — capture them separately
            ${log_path}    ${barcode_report}=    Process HTML File    ${HTML_FILE}
            PP5 Increment Done And Technically Checked
            PP5 Increment Log Transactions
        ELSE
            Log To Console With Timestamp    System exception processing ${HTML_FILE}: ${proc_err}
            PP5 Increment Technically Checked
            PP5 Increment Log Transactions
            CONTINUE
        END

        ${date}=    Evaluate    __import__('datetime').datetime.now().strftime('%Y_%m_%d')
        IF    '${barcode_report}' != 'None'
            log to console     Here in if condition
            ${barcode_filename}=    Evaluate    os.path.basename(r'''${barcode_report}''')    os
            Log To Console With Timestamp    ${barcode_filename}
            ${destination_path_payment_BarcodeList} =     Set Variable    ${primary_config['PathArchives_Final']}${/}LogReports\\${date}\${barcode_filename}
            Run Keyword And Ignore Error    Copy File    ${barcode_report}    ${destination_path_payment_BarcodeList}
        END
        Send Email Final Report    ${log_path}
        ${destination_path_payment_SettlementList} =     Set Variable    ${primary_config['PathArchives_Final']}${/}LogReports\\${date}\\${primary_config['FileName_ReportPaymentSettlementList']}${date}.csv
        Run Keyword And Ignore Error   Copy File    ${log_path}    ${destination_path_payment_SettlementList}
    END

Moving Cover Page Template
    ${Copy_cover_page}=     Set Variable     ${primary_config['PathToReadCoverPageTemplate']}${/}CoverPageTemplate.xlsx
    
    ${target_path}=    Join Path    ${primary_config['PathTemp']}    CoverPageTemplate.xlsx
    Copy File    ${Copy_cover_page}    ${target_path}            
    Log To Console    File moved to: ${target_path}

Initialize SAP System PP5
    [Documentation]    Kills existing SAP processes and launches SAP Logon
    Close SAP
    Start Process    C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe    shell=True
    Sleep    5s
    log to console        Starting SAP Logon for PP5
    # Wait for element    image:${CURDIR}${/}..${/}data${/}images${/}sap.png    timeout=30
    log to console    SAP Logon started successfully for PP5
    Connect To Session
    Open Connection    ${primary_config['PP5'].split('|')[1]}
    Disable Screenshots On Error
    Log To Console With Timestamp    SAP system initialized PP5

Close SAP
    [Documentation]    Kills existing SAP processes and launches SAP Logon

    Run Process    taskkill    /F    /T    /IM    saplogon.exe    shell=True
    Run Process    taskkill    /F    /T    /IM    sapgui.exe    shell=True

Login To SAP Application PP5
    [Documentation]    Performs SAP login with credentials

    Log To Console With Timestamp    Initializing SAP system...
    Evaluate                __import__('dotenv').load_dotenv(r'''${EXECDIR}/.env''')
    ${ENV}=    Get Environment Variable    ENV
    ${User_Password}=           Get Environment Variable    PP5
    ${User_name}=    Set Variable   ${primary_config['PP5'].split('|')[0].split('_')[0]}
    Set Global Variable    ${User_name}
    Input Text      ${LOC_USERNAME}      ${User_name}
    Input Password    ${LOC_PASSWORD}     ${User_Password}
    Input Text        ${LOC_LANGUAGE}    ${primary_config['Language']}
    Input Text        ${LOC_CLIENT}      ${primary_config['Client']}
    Send Vkey      0
    Log To Console With Timestamp    SAP login completed PP5

Load Transaction Screen PP5
    [Documentation]    Opens the FX Invoice Monitor transaction
    Log To Console With Timestamp    Loading transaction: ${primary_config['Transaction']}
    Run Transaction    /n${primary_config['Transaction'].split('|')[1]}
    Maximize Window   
    Log To Console With Timestamp    Transaction screen loaded

Load Transaction Screen F110 PP5
    [Documentation]    Opens the FX Invoice Monitor transaction
    Log To Console With Timestamp    Loading transaction: ${primary_config['Transaction'].split('|')[0]}
    Run Transaction    ${primary_config['Transaction'].split('|')[0]}
    Log To Console With Timestamp    Transaction screen loaded

Select Custom SAP Variant
    [Documentation]    Selects and applies SAP variant using filter
    
    Click Toolbar Button    ${LOC_SHELL_0}    /XFT/VAR_AKT
    Set Focus               ${LOC_VAR_WINDOW}
    Select Table Row        ${LOC_VAR_WINDOW}    0
    Sleep    1s
    RPA.Desktop.Press Keys    ctrl    a
    Sleep    1s
    Click Toolbar Button    ${LOC_VAR_WINDOW}    &MB_FILTER
    Input Text    ${LOC_VAR_FILTER}    ${primary_config['Variant_SAP_User']}
    RPA.Desktop.Press Keys    enter
    Sleep    1s
    Set Focus               ${LOC_VAR_WINDOW}
    Select Table Row        ${LOC_VAR_WINDOW}    0
    Send Vkey    0

Upload Barcode Data To SAP
    [Documentation]    Uploads barcode list file to SAP using offset search method
    Log To Console With Timestamp    Uploading barcode data to SAP...
    Execute Barcode Upload Via Offset Search    ${primary_config['PathTemp']}
    Apply Custom Layout Filter    ${User_name}
    Log To Console With Timestamp    Barcode upload completed

Log To Console With Timestamp
    [Documentation]    Logs message to console with timestamp prefix
    [Arguments]    ${message}
    ${timestamp}=    Get Current Date    result_format=%Y-%m-%d %H:%M:%S
    Log To Console     [${timestamp}] ${message}

Download All HTML Files PP5
    [Arguments]    ${FILES}
    
    ${today}=    Get Current Date    result_format=%Y-%m-%d
    ${previous_day}=    Set Variable    ${primary_config['PreviousDay']}
    ${today_date}=    Add Time To Date    ${today}    ${previous_day} days    result_format=%d.%m.%Y
    ${IDENT_STRING}=    Set Variable    ${primary_config["PP5_Identification"]}
    ${IDENT_STRING}=    Strip String    ${IDENT_STRING}
    @{IDENTIFICATIONS}=    Split String    ${IDENT_STRING}    |
    FOR    ${identification}    IN    @{IDENTIFICATIONS}
        Log To Console With Timestamp     ${identification}
        ${status}    ${err_msg}=    Run Keyword And Ignore Error
        ...    Perform HTML Extraction    ${today_date}    ${identification}

        IF    '${status}' == 'PASS'
            ${HTML_FILE}=    Set Variable    ${err_msg}

            IF    '${HTML_FILE}' == 'None' or '${HTML_FILE}' == '${EMPTY}'
                # Business exception — no data, payment list disabled, auth error
                Log To Console With Timestamp    No data for ${identification}
                PP5 Increment Technically Checked
                PP5 Increment Log Transactions
            ELSE
                ${file_exists}=    Run Keyword And Return Status
                ...    File Should Exist    ${HTML_FILE}
                IF    ${file_exists}
                    # Success — file downloaded
                    Append To List    ${FILES}    ${HTML_FILE}
                    Log To Console With Timestamp     Downloaded ${HTML_FILE}
                    PP5 Increment Done And Technically Checked
                    PP5 Increment Log Transactions
                ELSE
                    # Business exception — extraction ran but file missing
                    Log To Console With Timestamp    File not found after extraction: ${HTML_FILE}
                    PP5 Increment Technically Checked
                    PP5 Increment Log Transactions
                END
            END
        ELSE
            Log To Console With Timestamp    System exception for ${identification}: ${err_msg}
        END
        Log To Console With Timestamp    Here for loop completed for ${identification}
    END

Perform HTML Extraction
    [Arguments]     ${run_date}     ${Identification}
    Log To Console With Timestamp       Here in Perform HTML Search
    Wait Until Keyword Succeeds    20s    2s    Element should Be Present     wnd[0]/usr/ctxtF110V-LAUFD
    Set Focus    wnd[0]/usr/ctxtF110V-LAUFD 
    Input Text   wnd[0]/usr/ctxtF110V-LAUFD      ${run_date}   
    Log To Console With Timestamp      ---------------- Here in PP5 Started -----------------
    Wait Until Keyword Succeeds    20s    2s    Element should Be Present    ${MATHCODE_BUTTON}
    Set Focus    ${MATHCODE_BUTTON}
    Log To Console With Timestamp      HEre after clicking selct checkbox
    Send Vkey     4
    Wait Until Keyword Succeeds    20s     2s     Element should Be Present    ${VPMP_FILLING} 
    Input text     ${VPMP_FILLING}     ${Identification}
    Input Text    ${No_Of_Hits}    20
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element     wnd[1]/tbar[0]/btn[0]
    ${no_data_found}=     Get Value    wnd[0]/sbar
    Log To Console With Timestamp      ${no_data_found}
    IF  'No values found' in "${no_data_found}"
        Log To Console With Timestamp    No data found for this Identification PP5:- ${Identification}
        RETURN    None
    END
    Wait Until Keyword Succeeds    20s     2s     Element should Be Present     ${Selecting_row}
    Set Focus     ${Selecting_row}
    Log To Console With Timestamp       Clicked on Seelectin row
    # here youn have to set the logic  of Table seclectiopn
    #-----
    # till here
    ${match_date}=    Set Variable      False
    FOR    ${i}    IN RANGE    3    22
        ${locator}=    Set Variable    wnd[1]/usr/lbl[1,${i}]

        ${exists}=    Run Keyword And Return Status
        ...    Get value    ${locator}

        IF    ${exists}
            ${date}=    Get Value    ${locator}

            Log To Console    Checking ${locator} -> ${date}

            IF    $date == $run_date
                Log To Console    MATCH FOUND -> ${locator}
                Click Element    ${locator}    
                BREAK
            END
        END
    END

     IF    not ${match_date}
        Log To Console      no matching date found
        RETURN    None
        
    END
    Wait Until Keyword Succeeds    20s     2s     Element should Be Present    wnd[1]/tbar[0]/btn[0]
    # Final enter button of Appeared sheet 
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element   wnd[1]/tbar[0]/btn[0]

    ${status}=     Check Payment List Status
    Log To Console With Timestamp     ${status}
    IF  not ${status}
        Log To Console With Timestamp    Payment List Disabled in This Identification PP5 :- ${Identification}
        RETURN    None
    END
    RPA.Desktop.Press Keys     alt    e
    RPA.Desktop.Press Keys     y
    RPA.Desktop.Press Keys     i
    ${key_not_pressed}=    Run Keyword And Return Status     Wait Until Keyword Succeeds    20s     5s     Element should Be Present     wnd[1]/tbar[0]/btn[0]
    IF  not ${key_not_pressed}
        RPA.Desktop.Press Keys     i
    END
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element   wnd[1]/tbar[0]/btn[0]
    ${PATH_TO_SAVE_KZ_FOLDER}=    Set Variable      ${primary_config['PathToSaveKZFolder']}
    Wait For Element      ${Payment_list_found}     10

    RPA.Desktop.Press Keys     alt    l
    RPA.Desktop.Press Keys     e
    Wait For Element     ${Local_file_f110}    10
    RPA.Desktop.Press Keys     l

    Log To Console With Timestamp     Here after Local Send
    Wait Until Keyword Succeeds    20s    2s     Element Should Be Present       wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]
    Select Radio Button    wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]    

    Wait Until Keyword Succeeds    20s    2s     Element Should Be Present     wnd[1]/tbar[0]/btn[0]
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element   wnd[1]/tbar[0]/btn[0]
    Log To Console With Timestamp     Input Text In file
    Input Text     wnd[1]/usr/ctxtDY_PATH    ${PATH_TO_SAVE_KZ_FOLDER}
    Input Text     wnd[1]/usr/ctxtDY_FILENAME    Payment List_${identification}.htm
    ${file_path}=    Set Variable    ${PATH_TO_SAVE_KZ_FOLDER}\\Payment List_${identification}.htm
    Wait Until Keyword Succeeds    20s    2s     Element Should Be Present    wnd[1]/tbar[0]/btn[11]
    Set Focus     wnd[1]/tbar[0]/btn[11]
    Click Element   wnd[1]/tbar[0]/btn[11]

    Wait For Element    ${Not_downlaod}    5
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element   wnd[1]/tbar[0]/btn[0]
    Wait For Element     ${Downloaded_file}      10
    Send Vkey    3
    RETURN    ${file_path}

Process HTML File
    [Arguments]    ${HTML_FILE}
    Log To Console With Timestamp       Processing ${HTML_FILE}
    ${PATH_ARCHIVE}=     Set Variable     ${primary_config['PathArchives']}
    Set Global Variable    ${PATH_ARCHIVE}
    Load Transaction Screen PP5
    ${value_date}    ${kz_list}    ${rem_acc}    ${data}    ${is_cgi}=    
    ...    Extract Robot Data    ${HTML_FILE}     ${primary_config['PathArchives']}
    IF    not ${is_cgi}
        Log To Console With Timestamp       SKIPPED: ${HTML_FILE} is not a 4 CGI payments file
        Log To Console With Timestamp      Here in Not CGI file Report Generation
        Remove Directory     ${primary_config['PathReportPaymentSettlementList']}    recursive=True
        ${date}=    Evaluate    __import__('datetime').datetime.now().strftime('%Y_%m_%d')
       
        Create Directory    ${primary_config['PathReportPaymentSettlementList']}
        ${report_file}=    Set Variable    ${primary_config['PathReportPaymentSettlementList']}${primary_config['FileName_ReportPaymentSettlementList']}${date}.csv
        Create File    ${report_file}    status,CR_DR,currency,amount,remitter_name,remitter_address,remitter_account,beneficiary_name,beneficiary_address,beneficiary_account,swift_name,swift_code,value_date,creation_date,kz_number,barcodes,company_code,log_message\n
        Log To Console With Timestamp     EMPTY REPORT GENERATED Payment Settlement : ${report_file}
       
        RETURN    ${report_file}    None
    END
    ${Company_code}=     Get From Dictionary    ${data}    company_code
    ${PREFFIX}=    Set Variable    ${primary_config[str(${Company_code})].split('|')[0]}
    Log To Console With Timestamp     ${Company_code} ---- ${PREFFIX}
    Load Transaction Screen PP5
    Sleep   2s
    Log To Console With Timestamp       Here after loading transaction
    Select Custom SAP Variant
    Upload Barcode Data To SAP
    ${KZ_FOLDER_PATH}     ${COVER_PAGE_PATH}=
    ...     Extract Data      ${HTML_FILE}     ${PATH_ARCHIVE}    ${PREFFIX}
    Log To Console With Timestamp   KZ path ${KZ_FOLDER_PATH}
    @{BARCODE_RECORDS}=    Create List
    # Pdf Generate Flag 
    ${pdf_generate}=     Set Variable     False
    @{Payment_settlement_Record}=    Create List
    ${vendor_outer}=    Get From Dictionary    ${data}    vendors    
    FOR    ${vendor}    IN    @{vendor_outer} 
        Log To Console With Timestamp     HEre Vendor Data is This :- ${vendor}
        Log To Console With Timestamp     ------------------------------
        FOR    ${barcode}    IN    @{vendor['barcodes']}
            Log To Console With Timestamp     Processing Barcode :- ${barcode} in Vendor Name:- ${vendor['beneficiary_name']}
            Log    Processing barcode ${barcode}
            ${pdf_generate}=    Process for Barcode    ${barcode}    ${KZ_FOLDER_PATH}

            ${pdf_file}=    Set Variable    ${KZ_FOLDER_PATH}\\${barcode}_${Pdf_number}.pdf
            ${status}=      Set Variable    Success

            # Added record of KZ to db 
            ${inserted_id}=    Insert And Update Payment Record    ${data}    ${vendor}    ${primary_config['CR_DR']}    ${status}    ${barcode}
            Log to console     Data inserted into W.R.T ${barcode}
            ${record}=    Create List
            ...    ${status}
            ...    ${primary_config['CR_DR']}
            ...    ${vendor['currency']}
            ...    ${vendor['amount']}
            ...    ${data['remitter_name']}
            ...    ${data['remitter_address']}
            ...    ${data['remmiter_account']}
            ...    ${vendor['beneficiary_name']}
            ...    ${vendor['beneficiary_address']}
            ...    ${vendor['account_number']}
            ...    ${vendor['swift_name']}
            ...    ${vendor['bank_number']}
            ...    ${data['value_date']}
            ...    ${data['value_date']}
            ...    ${vendor['KZ']}
            ...    ${barcode}
            ...    ${data['company_code']}
            ...    Processed
            Append To List    ${Payment_settlement_Record}    ${record}
            ${record_barcode_list}=      Create List    ${barcode}    ${barcode}.pdf    ${pdf_file}    ${status}
            Append To List    ${BARCODE_RECORDS}    ${record_barcode_list}
            # Clear filters (same as your code)
            Wait For Element      ${Clear_filter}   15
            RPA.Desktop.Click     ${Clear_filter}
            ${Delete_filter_not_found}=     Run Keyword And Return Status    Wait for element      ${Delete_filter}   10
            IF   not ${Delete_filter_not_found} 
                Wait For Element      ${Clear_filter}   15
                RPA.Desktop.Click     ${Clear_filter}
                Wait for element      ${Delete_filter}   10
                RPA.Desktop.Click     ${Delete_filter}
            END
            RPA.Desktop.Click     ${Delete_filter}
        END
    END
    ${barcode_report}=    Generate Barcode Report    @{BARCODE_RECORDS}
    ${log_path}=    Generate Payment Settlement    @{Payment_settlement_Record}
    Log To Console With Timestamp     ${is_cgi}
    # Log To Console With Timestamp      Here After BArcode Report
    IF   ${pdf_generate}
        Merge Cover Excel with Barcode PDFs for KZ
        ...     ${COVER_PAGE_PATH}
        ...     ${KZ_FOLDER_PATH}
        ...     ${PATH_ARCHIVE}
        ...     ${data}
    END
    RETURN    ${log_path}    ${barcode_report}

Process for Barcode
    [Arguments]     ${barcode}      ${KZ_FOLDER}
    Log To Console With Timestamp       \n Here in Barcode Processing - ${barcode}

    # Adding F1 Docuement filter barcode
    Wait for Element    ${F1_document_header}   15
    RPA.Desktop.Click    ${F1_document_header}      right_click
    Sleep   5s
    Wait for Element    ${IMG_SETFILTER}    10
    RPA.Desktop.Click     ${IMG_SETFILTER}
    Log To Console With Timestamp       Pressed Set filter in F1
    ${data_not_found}=    Run keyword And Return Status    Wait Until Keyword Succeeds    20s     2s    Element Should be Present   wnd[1]/usr/txtMESSTXT1
    Log To Console With Timestamp     ${data_not_found}
    IF  ${data_not_found}
        ${data_not_found}=    Get Value     wnd[1]/usr/txtMESSTXT1
        IF   "contain any entries" in "${data_not_found}"
            Log To Console With Timestamp     No data found for barcode ${barcode}. Skipping this barcode.
            Click Element     wnd[1]/tbar[0]/btn[0]
            RETURN   False
        END
    END
    Wait for Element     ${F1_doc_value}    10
    Wait Until Keyword Succeeds    20s     2s    Element Should be Present     wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW
    Input Text      wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW      ${barcode}
    Click Element     wnd[1]/tbar[0]/btn[0]
    Wait For Element     ${Current_status_header}    10
    RPA.Desktop.Click    ${Current_status_header}    right_click
    Wait for Element    ${IMG_SETFILTER}    10
    RPA.Desktop.Click     ${IMG_SETFILTER}
    Log To Console With Timestamp       Pressed Set filter in CurrentStatus
    # Updating Current Status
    Wait Until Keyword Succeeds    20s     2s    Element Should be Present    wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW    5
    Input Text    wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW     Posted
    Click Element     wnd[1]/tbar[0]/btn[0]
    Wait for Element    ${Doc_creation}     20
    RPA.Desktop.Click     ${Doc_creation}    left_click
    Wait For Element     ${Doc_File_system}     20
    Sleep    2s
    RPA.Desktop.Press Keys      d
    ${found_export}=    Run Keyword And Return Status    Wait For Element    ${Export_doc}    20
    IF   ${found_export}
        RPA.Desktop.Click     ${Export_doc}
        PDF Export Approach     ${KZ_FOLDER}    ${barcode}
    ELSE
        PDF Export Approach     ${KZ_FOLDER}    ${barcode}
    END     
    Wait for element      ${cross_cut}      25
    RPA.Desktop.Click    ${cross_cut}
    RETURN     True

    
PDF Export Approach
    [Arguments]    ${KZ_FOLDER}    ${BARCODE}

    Wait for Element    ${IMG_ORIGINAL}    20
    RPA.Desktop.Click    ${IMG_ORIGINAL}

    FOR     ${i}   IN RANGE    3
        RPA.Desktop.Press Keys    shift    tab
        Sleep     0.5s
    END
    RPA.Desktop.Press Keys    enter
    # Close Edge if it appears
    Run Keyword And Return Status    Wait For Element    ${Downloaded_file}      20
    Log To Console With Timestamp       Downlaod image detected 
    Close Edge PDF Windows
    # Wait until SAP creates a folder in tmp
    Wait Until Keyword Succeeds    20x    1s    List Directories In Directory    ${SAP_TMP}
    ${SAP_TMP}=    Set Variable    ${primary_config["TempSAPGUIFolder"]}
    # Create Directory    ${SAP_TMP}
    Log To Console With Timestamp      Here :- ${SAP_TMP}
    ${TEMP_PATH}=    Set Variable     C:\\Users\\%{USERNAME}\\AppData\\Local\\SAP\\SAP${Space}GUI\\tmp\\
    # Get folders with absolute paths
    @{folders}=    List Directories In Directory    ${TEMP_PATH}    absolute=True
    # Find latest created folder
    ${latest_folder}=    Evaluate    max(${folders}, key=os.path.getmtime)    modules=os
    Log To Console With Timestamp     Latest folder: ${latest_folder}
    # Wait until PDF appears
    Wait Until Keyword Succeeds    20x    1s    List Files In Directory    ${latest_folder}    *.pdf
    @{pdfs}=    List Files In Directory    ${latest_folder}    *.pdf
    ${latest_pdf}=    Evaluate    max(${pdfs}, key=lambda x: os.path.getmtime(os.path.join(r"${latest_folder}", x)))    modules=os
    ${pdf_path}=    Set Variable    ${latest_folder}\\${latest_pdf}
    Monitor Windows    ${latest_pdf}    10    1    True
    ${destination}=    Set Variable    ${KZ_FOLDER}\\${BARCODE}_${Pdf_number}.pdf
    Log To Console With Timestamp     Moving ${pdf_path} -> ${destination}
    # Copy File    ${pdf_path}    ${SAP_TMP}
    Log To Console With Timestamp      Here before Move Pdf to SAP
    # Move Pdf From Env    ${pdf_path}    ${SAP_TMP}
    # ${pdf_path}=    Set Variable    ${SAP_TMP}\\${latest_pdf}
    Copy File    ${pdf_path}    ${destination}
    # Clean tmp folder after processing
    ${Pdf_number}=     Evaluate      ${Pdf_number}+1
    Remove Directory    ${latest_folder}    recursive=True


Wait Until PDF Unlock
    [Arguments]    ${file_path}
    Log To Console With Timestamp      Here in Until PDf Unlock
    Wait Until File Free    ${file_path}


Generate Barcode Report
    [Arguments]    @{records}
   
    Remove Directory    ${primary_config['PathReportBarcodesList']}    recursive=True
    Create Directory    ${primary_config['PathReportBarcodesList']}
    ${date}=    Evaluate    __import__('datetime').datetime.now().strftime('%Y_%m_%d')
    ${report_file}=    Set Variable    ${primary_config['PathReportBarcodesList']}/report_barcode_list_${date}.csv
    Create File    ${report_file}    Barcode,File Name,Destination Path,Status\n
    FOR    ${rec}    IN    @{records}
        ${length}=    Get Length    ${rec}
        Run Keyword If    ${length} < 4    Log To Console With Timestamp     Invalid record: ${rec}
        Append To File    ${report_file}
        ...    ${rec[0]},${rec[1]},${rec[2]},${rec[3]}\n
    END
    Log To Console With Timestamp     Barcode report generated: ${report_file}
    RETURN    ${report_file}

Generate Payment Settlement
    [Arguments]    @{records}
   
    Remove Directory    ${primary_config['PathReportPaymentSettlementList']}    recursive=True
    Create Directory    ${primary_config['PathReportPaymentSettlementList']}
    ${date}=    Evaluate    __import__('datetime').datetime.now().strftime('%Y_%m_%d')
    ${report_file}=    Set Variable    ${primary_config['PathReportPaymentSettlementList']}/${primary_config['FileName_ReportPaymentSettlementList']}_${date}.csv
    Create File    ${report_file}    
    ...    status,CR_DR,currency,amount,remitter_name,remitter_address,remitter_account,beneficiary_name,beneficiary_address,beneficiary_account,swift_name,swift_code,value_date,creation_date,kz_number,barcodes,company_code,log_message\n
    FOR    ${rec}    IN    @{records}
        ${length}=    Get Length    ${rec}
        Run Keyword If    ${length} < 18    Log To Console With Timestamp     Invalid record: ${rec}
        Append To File    ${report_file}
        ...    ${rec[0]},${rec[1]},${rec[2]},${rec[3]},${rec[4]},${rec[5]},${rec[6]},${rec[7]},${rec[8]},${rec[9]},${rec[10]},${rec[11]},${rec[12]},${rec[13]},${rec[14]},${rec[15]},${rec[16]},${rec[17]}\n
    END
    Log To Console With Timestamp     Payment settlemen list Genreated at Log: ${report_file}
    RETURN    ${report_file}


Execute Barcode Upload Via Offset Search
    [Documentation]    Uploads barcodes using image recognition with 500:0 offset logic
    [Arguments]    ${barcode_file_path}

    Wait for Element    ${IMG_INVOICE}    60
    # RPA.Desktop.Move Mouse    ${IMG_INVOICE}
    ${x}    ${y}    ${top}     ${bottom}=    RPA.Desktop.Find Element    ${IMG_INVOICE}
    RPA.Desktop.Click    ${IMG_INVOICE}
    RPA.Desktop.Press Keys    tab
    RPA.Desktop.Press Keys    tab
    Sleep   0.5s
    ${found}=    Run Keyword And Return Status    Wait For Element    ${Multiple_Selection_Image}    5

    IF    ${found}
        RPA.Desktop.Click    ${Multiple_Selection_Image}
    ELSE
        RPA.Desktop.Click    ${Multiple_Selection_Image_safe}
    END
    Wait Until Keyword Succeeds    20s    2s    Element should Be Present    wnd[1]/tbar[0]/btn[23]
    Click Element     wnd[1]/tbar[0]/btn[23]

    Log To Console With Timestamp     After clicking IMg Barcode copy from file 
    Wait Until Keyword Succeeds    20s     2s    Element Should Be Present     wnd[2]/usr/ctxtDY_PATH
    Input Text    wnd[2]/usr/ctxtDY_PATH    ${barcode_file_path}
    Input Text     wnd[2]/usr/ctxtDY_FILENAME    BarcodeList.txt
    Click Element     wnd[2]/tbar[0]/btn[0]
    Wait Until Keyword Succeeds    20s    2s    Element should Be Present    wnd[1]/tbar[0]/btn[8]
    Click Element    wnd[1]/tbar[0]/btn[8]
    Wait for Element    ${IMG_INVOICE}    10
    RPA.Desktop.Click    ${IMG_INVOICE}
    Wait for Element     ${Perform_Search_Image}    5
    RPA.Desktop.Click    ${Perform_Search_Image}

    

Apply Custom Layout Filter
    [Documentation]    Applies layout filter
    [Arguments]    ${variant_name}
    
    Wait for Element     ${Hit_filter_Image}    10
    RPA.Desktop.Click    ${Hit_filter_Image}
    Wait For Element     ${LOC_Choose_LAYOUT_SHELL}     10
    RPA.Desktop.Click    ${LOC_Choose_LAYOUT_SHELL}    
    Log To Console With Timestamp       Here after loc layout sheelll
    Wait for Element    ${IMG_LAYOUT}    20
    RPA.Desktop.Click    ${IMG_LAYOUT}    right_click
    Wait for Element    ${IMG_SETFILTER}    20
    RPA.Desktop.Click    ${IMG_SETFILTER}
    ${Layout_active}=   Run Keyword And Return Status    Wait For Element    ${Layout_User_Active}    20
    IF    not ${Layout_active}
        RPA.Desktop.Press Keys     shift    tab
    END
    Type Text   /${variant_name}
    RPA.Desktop.Press Keys  tab
    Type Text  ${variant_name}
    RPA.Desktop.Press Keys    enter
    Wait for Element    ${default_setting_click}    20
    RPA.Desktop.Press Keys     enter

  

Merge Cover Excel with Barcode PDFs for KZ
    [Arguments]      ${cover_excel}    ${kz_folder}    ${path_archives}     ${data}
    ${merged_pdf}=    
    ...     Merge Cover Excel With Kz Data  ${data}     ${cover_excel}    ${kz_folder}      ${path_archives}    
    Log To Console With Timestamp     Merged PDF saved at: ${merged_pdf}


Send Email Final Report
    [Arguments]    ${log_path}

    # 2. Map Dynamic DB Parameters for the Email Configuration
    ${email_sender}=       Set Variable    ${primary_config['EmailSender']}
    ${email_recipient}=    Set Variable    ${primary_config['EmailAddress']}
    ${email_subject}=      Set Variable   RPA0032-India FX Payment Settlement Report
   
    # 3. Handle HTML Body Content
    # (Assuming Email_Body_FinalProcess contains the direct HTML string. 
    # If it's a file path instead, you would use: Get File ${primary_config['Email_Body_FinalProcess']})
    # ${body_msg}=           Set Variable    ${primary_config['Email_Body_FinalProcess']}
 
    # 4. Authorize SMTP Connection
    Authorize SMTP    
    ...    account=${email_sender}  
    ...    password=${EMPTY}    
    ...    smtp_server=${secondary_config['SMTP_Server']}   
    ...    smtp_port=${secondary_config['SMTP_Port']}
 
    Log To Console With Timestamp     Preparing to send final report to ${email_recipient} And ${log_path}...
     
    # 5. Send Message using the Mapped Parameters and Attach Log File
    Send Message    
    ...    sender=${email_sender}
    ...    recipients=${email_recipient}
   ...    subject=${email_subject}
    ...    body=The FX Payment Document Preparation Is Complete. Please Find the report Attached.
    ...    html=True
   ...    attachments=${log_path}
 
    Log To Console With Timestamp     📧 Report successfully sent with attached log!



Check Payment List Status
    ${status}=    Set Variable    NOT_FOUND

    ${status}=    Run Keyword And Return Status
    ...    Element Should Be Present    ${payment_list_locator}

    RETURN    ${status}

PP5 Increment Log Transactions
    ${lib}=    Get Library Instance    RobotProcessLibrary
    ${new_tx}=    Evaluate    int($lib.config.Log_Transactions) + 1
    Evaluate    setattr($lib.config, 'Log_Transactions', str(${new_tx}))
    Log To Console    PP5 Log_Transactions: ${new_tx}

PP5 Increment Done And Technically Checked
    ${lib}=    Get Library Instance    RobotProcessLibrary
    ${new_done}=    Evaluate    int($lib.config.Log_Done) + 1
    ${new_loop}=    Evaluate    int($lib.config.Log_Looping) + 1
    Evaluate    setattr($lib.config, 'Log_Done', str(${new_done}))
    Evaluate    setattr($lib.config, 'Log_Looping', str(${new_loop}))
    Log To Console    PP5 Completed: ${new_done} | TechnicallyChecked: ${new_loop}

PP5 Increment Technically Checked
    ${lib}=    Get Library Instance    RobotProcessLibrary
    ${new_loop}=    Evaluate    int($lib.config.Log_Looping) + 1
    Evaluate    setattr($lib.config, 'Log_Looping', str(${new_loop}))
    Log To Console    PP5 TechnicallyChecked (business exception): ${new_loop}