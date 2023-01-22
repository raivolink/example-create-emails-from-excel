*** Settings ***
Documentation       Robot for sending emails based on input Excel
...                 Each Excel sheet will be read in and used
...                 as input for email

Library             RPA.Excel.Files
Library             Collections
Library             RPA.Email.ImapSmtp    smtp_server=smtp.gmail.com    smtp_port=587
Library             RPA.Robocorp.Vault
Library             OperatingSystem


*** Variables ***
${EXCEL_FILE}       ${CURDIR}${/}devdata${/}Employee data.xlsx
${VAULT_NAME}       RoboMail
${TEMPLATE_FILE}    ${CURDIR}${/}devdata${/}message.template
${SUBJECT}          Email notificastion from Robot Robotson
${RECIPIENT}        targetmail@mail.com


*** Tasks ***
Email Excel Data
    &{employees_from_excel}=    Read Excel Data To Dictionary
    Send All Mails    &{employees_from_excel}


*** Keywords ***
Read Excel Data To Dictionary
    Open Workbook    ${EXCEL_FILE}
    &{employees_from_excel}=    Create Dictionary
    @{sheets}=    List Worksheets
    FOR    ${sheet}    IN    @{sheets}
        &{current_employee}=    Create Dictionary
        ${current_worksheet}=    Read Worksheet    ${sheet}
        ${first_name}=    Get Cell Value    2    B
        ${last_name}=    Get Cell Value    3    B
        Set To Dictionary    ${current_employee}
        ...    First name=${first_name}
        ...    Last name=${last_name}
        Set To Dictionary    ${employees_from_excel}
        ...    ${sheet}=&{current_employee}
    END
    Close Workbook
    RETURN    &{employees_from_excel}

Send All Mails
    [Documentation]    Generates email body based on input dictionary
    ...    Sends email to recipients
    [Arguments]    &{data_dict}
    ${secret}=    Get Secret    ${VAULT_NAME}
    Authorize    account=${secret}[username]    password=${secret}[password]
    ${mail_template}=    Get File    ${TEMPLATE_FILE}
    FOR    ${employee_sheet}    IN    @{data_dict}
        ${first_name}=    Set Variable    ${data_dict['${employee_sheet}']}[First name]
        ${last_name}=    Set Variable    ${data_dict['${employee_sheet}']}[Last name]
        ${message}=    Replace Variables    ${mail_template}
        Send Message
        ...    sender=${secret}[username]
        ...    recipients=${RECIPIENT}
        ...    subject=${SUBJECT}
        ...    body=${message}
        ...    html=True
    END
