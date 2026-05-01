*** Settings ***
Documentation     FX Payment Workflow - High Level Business Flow
Resource    flow/SAP_process_flow.robot
Library           adapters/Library/RobotProcessLibrary.py


*** Test Cases ***
Sap autoamtion all type 
    Initialize Robot Process
    SAP Flow Identification Process
    End Robot Process