*** Settings ***
Test Teardown     Run Keyword If Test Failed    update automationcenter failure    ${TEST MESSAGE}
Library           SSHLibrary
Library           AutomationCenter
Library           WindowsAE
Library           lib\\custom_library.py
Library           SeleniumLibrary
Library           String
Library           BuiltIn
Library           Collections    #Library    RPA.Browser.Selinium    auto_close=${False}
Library           lib\\SharepointSignIn.py
Library           lib\\auditpending_status.py
Library           lib\\dataFromTextFile.py

*** Variables ***
#input
${aeuser}         ${EMPTY}
${aeparameters}    ${EMPTY}
${aeci}           ${EMPTY}
${aepassword}     ${EMPTY}
${aeprivatekey}    ${EMPTY}
${aedatastore}    ${EMPTY}
${timeout}        NONE
#result
${output}         ""
${error}          ""
${rc}             0
#${sharepointfolders}    https://skydrive3m.sharepoint.com/teams/ProjectAlpineExecution/Shared%20Documents/Forms/AllItems.aspx?csf=1&web=1&e=mjtKPr&cid=3ace0a0c%2D9027%2D47e5%2D98ca%2D99bfea0fb9d0&FolderCTID=0x0120000F402D7805F054428A3C5F00CD035802&id=%2Fteams%2FProjectAlpineExecution%2FShared%20Documents%2FService%20Management%20Office%2FQualityTools%2FQuality%20Review%20%2D%20Files%20to%20Audit&viewid=52849c25%2D869e%2D4061%2D8d08%2Da7d8051dc75b
#${browser}       edge
#${user_name}     ac5qdzz
#${password}      ZXCVzxcv2131

*** Test Cases ***
SAPtask
    [Timeout]    ${timeout}
    run
    Get Credential
    Log In
    Update AutomationCenter Output

*** Keywords ***
run
    &{aeparameters}=    AutomationCenter.json_parse    ${aeparameters}
    &{aeci}=    AutomationCenter.json_parse    ${aeci}
    set global variable    ${aeparameters}
    ${browser}=    AutomationCenter.Get Value    ${aeparameters}    Browser
    Set Global Variable    ${browser}
    ${sharepointfolders}=    AutomationCenter.Get Value    ${aeparameters}    sharepointfolders
    Set Global Variable    ${sharepointfolders}
    ${notePath}=    AutomationCenter.Get Value    ${aeparameters}    notePath
    Set Global Variable    ${notePath}
    ${orig wait} =    Set Selenium Implicit Wait    10 seconds
    Set Global Variable    ${orig wait}

Get Credential
    &{aedatastore}=    AutomationCenter.Json Parse    ${aedatastore}
    ${app_id}=    AutomationCenter.get_value    ${aeparameters}    cred_id
    &{entity_details}=    AutomationCenter.Get Entity    ${aedatastore}    ${app_id}
    &{vault_details}=    AutomationCenter.Get Entity    ${aedatastore}    ${entity_details["vaultid"]}
    Set To Dictionary    ${entity_details}    vault=${vault_details}
    ${user_name}=    AutomationCenter.Get Value    ${entity_details}    serviceAccount
    ${password}=    AutomationCenter.Resolve Sdb    ${entity_details}    ${aedatastore}
    Set Global Variable    ${user_name}
    Set Global Variable    ${password}

Log In
    ${output2}=    textFileData    ${notePath}
    ${list}=    Create List    --inprivate
    ${args}=    Create Dictionary    args=${list}
    ${desired caps}=    Create Dictionary    ms:edgeOptions=${args}
    Open Browser    ${sharepointfolders}    ${browser}    desired_capabilities=${desired caps}    #remote_url=http://localhost:9515
    Maximize Browser Window
    Sleep    10s
    sharepointAuth    ${user_name}    ${password}
    Sleep    25s
    #Click Element    //*[@class="ms-ContextualMenu-link root-229"][@name="Files"]
    #Choose File    //*[@class="ms-ContextualMenu-link root-229"][@name="Files"]    //acprd01/E/3M_CAC/SMO_AMA/mail_details.xlsx
    #Drag And Drop    //acprd01/E/3M_CAC/SMO_AMA/mail_details.xlsx    //*[@id="row455-8"]/div[2]/div[5]
    #Drag And Drop    file:${CURDIR}${/}${FILENAME}    //*[@id="row455-8"]/div[2]/div[5]
    #${count}=    Get Matching Xpath Count    //*[contains(., "QualityCheck")]
    ${count} =    Get Element Count    //*[contains(text(),"QualityCheck_Week")]    #//*[contains(., "QualityCheck")]
    Set Global Variable    ${count}
    #Log To Console    ${count}
    ${rangelimit} =    Evaluate    ${count} + 15
    ${startindex}=    Convert To String    ${output2}
    #Log To Console    ${rangelimit}
    #FOR    ${index}    IN RANGE    35    ${rangelimit}
    FOR    ${index}    IN RANGE    ${startindex}    ${rangelimit}
        ${convert}    Convert To String    ${index}
        #Log To Console    //*[contains(text(),"QualityCheck_Week_${convert}")]
        Click Element    //*[contains(text(),"QualityCheck_Week_${${convert}}")]
        sleep    10s
        Click Element    //*[@class="ms-Button-label label-129"][text()="Download"]
        sleep    15s
        Click Element    //*[contains(text(),"Quality Review - Files to Audit")]
    END
    ${output}=    Set Variable    ${rangelimit}
    Set Global Variable    ${output}
    ${output1}=    executefunc
    Set Global Variable    ${output1}

Update AutomationCenter Output
    ###-------- Define AutomationCenter Format result --------
    ${result}=    Update AutomationCenter    ${output1}    ${error}    ${rc}
