*** Comments ***
TODO
[ ] Create Excel
[ ] Build Email Template


*** Settings ***
Documentation       Robot for sending emails based on input Excel
...                 Each Excel sheet will be read in    and used
...                 as input for email

Library             RPA.Excel.Files


*** Variables ***
${EXCEL_FILE}       ${CURDIR}${/}devdata${/}Employee data.xlsx


*** Tasks ***
Minimal task
    Log To Console    \n
    Open Workbook    ${EXCEL_FILE}
    &{employees_from_excel}=    Create Dictionary
    @{sheets}=    List Worksheets
    Log To Console    ${sheets}
    ${worksheet}=    Read Worksheet
    ${first_name}=    Get Cell Value    2    B    Sheet2
    Log To Console    ${first_name}
    ${last_name}=    Get Cell Value    3    B
    Log To Console    ${last_name}
    Close Workbook
