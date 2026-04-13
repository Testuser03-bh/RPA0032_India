*** Settings ***
Documentation     Environment Setup Domain Logic
Library           OperatingSystem
Resource          ..${/}resources${/}locators.robot
Library    ..${/}adapters${/}Library${/}InitAllSettingsSQL.py

*** Variables ***
${PRIMARY_PROCESS_NAME}    RPA0032-INDIA
${SECONDARY_PROCESS_NAME}    RPA
${primary_config}
${secondary_config}


*** Keywords ***
Prepare Working Environment
    [Documentation]    Initializes project folder structure and copies template files
    ${primary_fetched_config}=    Get All Settings    ${PRIMARY_PROCESS_NAME}
    ${secondary_fetched_config}=    Get All Settings    ${SECONDARY_PROCESS_NAME}
    Set Global Variable    ${primary_config}    ${primary_fetched_config}
    Set Global Variable    ${secondary_config}    ${secondary_fetched_config}
    Log To Console With Timestamp    Initializing project environment...
    Remove Directory      ${primary_config['PathTemp']}     force=True

    Create Directory    ${primary_config['PathTemp']}
    Create Directory    ${primary_config['PathToSaveKZFolder']}
    Create Directory    ${primary_config['PathReportBarcodesList']}
    Create Directory    ${primary_config['PathReportPaymentSettlementList']}
    
    Log To Console With Timestamp    Environment setup completed