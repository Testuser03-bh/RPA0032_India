*** Settings ***
Documentation     SAP Operations Domain Logic
Library           Process
Library           SapGuiLibrary
Library           RPA.Desktop
Library           OperatingSystem
Library           String
Library           Collections
Library           DateTime
Resource          Purchase_Order_PP5.robot

    

*** Variables ***
${User_name}

*** Keywords ***
Workflow RPA0032 HP1
    [Arguments]    ${DOWNLOADED_HTML_FILES}

    Initialize SAP System HP1
    Login To SAP Application PP5
    Load Transaction Screen F110 HP1
    Download All HTML Files HP1    ${DOWNLOADED_HTML_FILES}
    RETURN    ${DOWNLOADED_HTML_FILES}

Close Existing SAP

    Run Process    taskkill    /F    /T    /IM    saplogon.exe    shell=True
    Run Process    taskkill    /F    /T    /IM    sapgui.exe    shell=True
    Run Process    cmd.exe    /c    rd /S /Q "%temp%\\gen_py"    shell=True


Initialize SAP System HP1
    [Documentation]    Kills existing SAP processes and launches SAP Logon

    Start Process    C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe    shell=True
    Log To Console With Timestamp     Starting SAP Logon for 
    Wait for element  image:${CURDIR}${/}..${/}data${/}images${/}sap.png  timeout=20
    Log To Console With Timestamp     SAP Logon started successfully for HP1
    Connect To Session
    Open Connection    ${primary_config['HP1'].split('|')[1]}
    Disable Screenshots On Error
    Log To Console With Timestamp     SAP system initialized PP5

Login To SAP Application PP5
    [Documentation]    Performs SAP login with credentials

    Log To Console With Timestamp     Initializing SAP system...
    Evaluate                __import__('dotenv').load_dotenv(r'''${EXECDIR}/.env''')
    ${ENV}=    Get Environment Variable    ENV
    ${User_Password}=           Get Environment Variable    HP1
    ${User_name}=    Set Variable   ${primary_config['PP5'].split('|')[0].split('_')[0]}
    Set Global Variable    ${User_name}
    Input Text      ${LOC_USERNAME}      ${User_name}
    Input Password    
    ...    ${LOC_PASSWORD}     ${User_Password}
    Input Text        ${LOC_LANGUAGE}    ${primary_config['Language']}
    Input Text        ${LOC_CLIENT}      ${primary_config['Client']}
    Send Vkey      0
    Log To Console With Timestamp     SAP login completed PP5

Load Transaction Screen F110 HP1
    [Documentation]    Opens the FX Invoice Monitor transaction
    Log To Console With Timestamp     Loading transaction: ${primary_config['Transaction'].split('|')[0]}
    Run Transaction    ${primary_config['Transaction'].split('|')[0]}
    Log To Console With Timestamp     Transaction screen loaded

Download All HTML Files HP1
    [Arguments]    ${Files}
    
    Log To Console With Timestamp    Here in HP1
    ${today}=    Get Current Date    result_format=%Y-%m-%d
    ${previous_day}=    Set Variable    ${primary_config['PreviousDay']}
    ${today_date}=    Add Time To Date    ${today}    ${previous_day} days    result_format=%d.%m.%Y
    Log To Console    Final Run Date: ${today_date}
    Log To Console With Timestamp    Today's date: ${today_date}
    ${IDENT_STRING}=    Set Variable    ${primary_config["HP1_Identification"]}
    ${IDENT_STRING}=    Strip String    ${IDENT_STRING}
    @{IDENTIFICATIONS}=    Split String    ${IDENT_STRING}    |
    FOR    ${identification}    IN    @{IDENTIFICATIONS}
        Log To Console With Timestamp    ${identification}
        ${HTML_FILE}=    Perform HTML Extraction HP1   ${today_date}    ${identification}
        
        ${exists}=    Run Keyword And Return Status
        ...    File Should Exist    ${HTML_FILE}

        IF    ${exists}
            Append To List    ${FILES}    ${HTML_FILE}
            Log To Console With Timestamp    Downloaded ${HTML_FILE}
        ELSE
            Log To Console With Timestamp    No data for ${identification}
        END
        Log To Console With Timestamp     HEre for loop empty currently

    END

 

Perform HTML Extraction HP1
    [Arguments]     ${run_date}     ${Identification}
    Log To Console With Timestamp      Here in Perform HTML Search

    Wait Until Keyword Succeeds    20s    2s    Element should Be Present     wnd[0]/usr/ctxtF110V-LAUFD
    Set Focus    wnd[0]/usr/ctxtF110V-LAUFD 
    Input Text   wnd[0]/usr/ctxtF110V-LAUFD      ${run_date}   
    Log To Console With Timestamp     Here in HP1 Processing :- ${Identification}
    Wait Until Keyword Succeeds    20s    2s    Element should Be Present    ${MATHCODE_BUTTON}
    Set Focus    ${MATHCODE_BUTTON}
    Log To Console With Timestamp     HEre after clicking selct checkbox
    Send Vkey     4
    Log To Console With Timestamp    HEre after clicking VKEy
    Log To Console With Timestamp      Filling Identification Value
    Wait Until Keyword Succeeds    20s     2s     Element should Be Present    ${VPMP_FILLING} 
    Input text     ${VPMP_FILLING}     ${Identification}
    Log To Console With Timestamp    here ater vpmp filing
    Input Text    ${No_Of_Hits}    20
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element     wnd[1]/tbar[0]/btn[0]

    ${no_data_found}=     Get Value    wnd[0]/sbar
    Log To Console With Timestamp     ${no_data_found}
    IF  'No values found' in "${no_data_found}"
        Log To Console With Timestamp     No data found for this Identification PP5:- ${Identification}
        RETURN    None
    END
    Wait Until Keyword Succeeds    20s     2s     Element should Be Present     ${Selecting_row}
    Set Focus     ${Selecting_row}
        Log To Console With Timestamp      Clicked on Selecting row
    Wait Until Keyword Succeeds    20s     2s     Element should Be Present    wnd[1]/tbar[0]/btn[0]
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element   wnd[1]/tbar[0]/btn[0]
    # Checking Payment List is Activating
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
    Sleep      2s
    ${Auth_error_Status}=     Run Keyword And Return Status     Wait For Element      ${Information_image}      10
    Log To Console With Timestamp  message=${Auth_error_Status}
    IF  ${Auth_error_Status}
        Log To Console With Timestamp     Authentication Error in This Identification :- ${Identification}
        Set Focus     wnd[1]/tbar[0]/btn[0]
        Click Element   wnd[1]/tbar[0]/btn[0]
        ${Auth_error_Status}=     Run Keyword And Return Status     Wait For Element      ${Information_image}      10
        IF  ${Auth_error_Status}
            Log To Console With Timestamp     Authentication Error in This Identification :- ${Identification}
            Set Focus     wnd[1]/tbar[0]/btn[0]
            Click Element   wnd[1]/tbar[0]/btn[0]
            RETURN    None
        END
        RETURN    None
    END

    ${PATH_TO_SAVE_KZ_FOLDER}=    Set Variable     ${primary_config['PathToSaveKZFolder']}
    Wait For Element    ${Payment_list_found}     10

    RPA.Desktop.Press Keys     alt    l
    RPA.Desktop.Press Keys     e
    Wait For Element     ${Local_file_f110}    10
    RPA.Desktop.Press Keys     l
    Log To Console With Timestamp    Here after Local Send
    Wait Until Keyword Succeeds    20s    2s     Element Should Be Present       wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]
    Select Radio Button    wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]    
    Wait Until Keyword Succeeds    20s    2s     Element Should Be Present     wnd[1]/tbar[0]/btn[0]
    Set Focus     wnd[1]/tbar[0]/btn[0]
    Click Element   wnd[1]/tbar[0]/btn[0]

    Log To Console With Timestamp    Input Text In file
    Input Text     wnd[1]/usr/ctxtDY_PATH    ${PATH_TO_SAVE_KZ_FOLDER}
    Input Text     wnd[1]/usr/ctxtDY_FILENAME    Payment List_${identification}.htm
    ${file_path}=    Set Variable    ${PATH_TO_SAVE_KZ_FOLDER}\\Payment List_${identification}.htm
    Wait Until Keyword Succeeds    20s    2s     Element Should Be Present    wnd[1]/tbar[0]/btn[11]
    Set Focus     wnd[1]/tbar[0]/btn[11]
    Click Element   wnd[1]/tbar[0]/btn[11]
    ${not_download_page}=   Run Keyword And Return Status    Wait For Element    ${Not_downlaod}    5
    IF    ${not_download_page}
        Wait For Element     wnd[1]/tbar[0]/btn[0]      timeout=5
        Click Element   wnd[1]/tbar[0]/btn[0]
    END
    # Ensuruing File Is downlaoded
    Wait For Element     ${Downloaded_file}      10
    Send Vkey    3
    RETURN    ${file_path}

