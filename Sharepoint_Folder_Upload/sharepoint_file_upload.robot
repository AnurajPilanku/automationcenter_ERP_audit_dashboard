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
Library           lib\\filemanupulation.py
Library           lib\\fileUploadWindow.py
Library           lib\\writefile.py

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
#${source_dir}    //acdev01/3M_CAC/ERP_Quality_Review/QualityReview
#${destination_dir}    //acdev01/3M_CAC/ERP_Quality_Review/sharepointUpload
#${foldername}    QualityCheck_Week
#${qualitypath}    \\\\acdev01\\3M_CAC\\ERP_Quality_Review\\sharepointUpload
#${lastweekrangefile}    \\\\acdev01\\3M_CAC\\ERP_Quality_Review\\fileupload_filerange.txt

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
    ${sharepointurl}=    AutomationCenter.Get Value    ${aeparameters}    sharepointurl
    Set Global Variable    ${sharepointurl}
    ${source_dir}=    AutomationCenter.Get Value    ${aeparameters}    source_dir
    Set Global Variable    ${source_dir}
    ${destination_dir}=    AutomationCenter.Get Value    ${aeparameters}    destination_dir
    Set Global Variable    ${destination_dir}
    ${foldername}=    AutomationCenter.Get Value    ${aeparameters}    foldername
    Set Global Variable    ${foldername}
    ${qualitypath}=    AutomationCenter.Get Value    ${aeparameters}    qualitypath
    Set Global Variable    ${qualitypath}
    ${lastweekrangefile}=    AutomationCenter.Get Value    ${aeparameters}    lastweekrangefile
    Set Global Variable    ${lastweekrangefile}
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
    ${filemanipulate}=    fileprocess    ${source_dir}    ${destination_dir}    ${foldername}    ${lastweekrangefile}
    Sleep    10s
    ${list}=    Create List    --inprivate
    ${args}=    Create Dictionary    args=${list}
    ${desired caps}=    Create Dictionary    ms:edgeOptions=${args}
    Open Browser    ${sharepointurl}    ${browser}    desired_capabilities=${desired caps}    #remote_url=http://localhost:9515
    Maximize Browser Window
    Sleep    10s
    sharepointAuth    ${user_name}    ${password}
    Sleep    25s
    ${count} =    Get Element Count    //*[contains(text(),"QualityCheck_Week")]
    Set Global Variable    ${count}
    ${rangelimit} =    Evaluate    ${count} + 15
    Set Global Variable    ${rangelimit}
    Sleep    10s
    Click Element    //*[text()="Upload"]    #//*[@class="ms-Button-label label-129"][contains(text(),"Upload")]
    sleep    5s
    Click Element    //*[text()="Folder"]    #//*[@class="ms-ContextualMenu-itemText label-394"][contains(text(),"Folder")]
    sleep    10s
    ${uploadfolder}=    sharepointfileupload    ${qualitypath}
    sleep    10s
    ${saverange}=    writetonote    ${lastweekrangefile}    ${rangelimit}
    ${output}=    Set Variable    ${rangelimit}
    Set Global Variable    ${output}

Update AutomationCenter Output
    ###-------- Define AutomationCenter Format result --------
    ${result}=    Update AutomationCenter    ${output}    ${error}    ${rc}
