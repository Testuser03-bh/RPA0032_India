*** Settings ***
Resource    ..${/}domain${/}Purchase_Order_PP5.robot
Resource    ..${/}domain${/}Purchase_order_HP1.robot
Resource    ..${/}domain${/}environment_setup.robot
Library          OperatingSystem
Library          ..${/}adapters${/}Python_helper.py
Library          ..${/}adapters${/}db_adapter.py
Library            ..${/}adapters${/}Library${/}InitAllSettingsSQL.py
Resource     ..${/}Resources${/}locators.robot

*** Variables ***
@{DOWNLOADED_HTML_FILES}
*** Keywords ***
SAP Flow Identification Process
    [Documentation]    Orchestrates the complete FX payment document preparation process
    
    # Setup Environment
    Prepare Working Environment
    # Closing All existing SAP sessions
    Close Existing SAP

    # Move  Cover Page Template
    Moving Cover Page Template

    # # Tracking all files
    ${DOWNLOADED_HTML_FILES}    Create List

    # # # # Processing and downloading HP1
    Workflow RPA0032 HP1    ${DOWNLOADED_HTML_FILES}
    # # Processing Full flow in PP5
    Main Workflow RPA0032 PP5    ${DOWNLOADED_HTML_FILES}

    Close SAP



