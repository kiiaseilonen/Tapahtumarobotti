*** Settings ***
Documentation       Get events from the website and save selected ones into an Excel file

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.Excel.Files
Library             RPA.Tables
Library             Collections


*** Tasks ***
Get events from the website and save selected ones into an Excel file
    Open the event website
    Create an Excel file with headers
    Get events from the website and save selected ones into an Excel file


*** Keywords ***
Open the event website
    Open Available Browser    http://www.koirat.com/tapahtumat

Create an Excel file with headers
    Create Workbook    Tapahtumat.xlsx
    Set Worksheet Value    1    1    Event Name
    Set Worksheet Value    1    2    Event Date
    Set Worksheet Value    1    3    Location
    Set Worksheet Value    1    4    Url
    Save Workbook

Get events from the website and save selected ones into an Excel file
    Wait Until Page Contains Element    class:c528
    Open Workbook    Tapahtumat.xlsx

    Wait Until Page Contains Element    class:c528
    @{rows}=    RPA.Browser.Selenium.Get WebElements    css:table.c528 tr
    FOR    ${row}    IN    @{rows}
        @{cities}=    RPA.Browser.Selenium.Get WebElements    css:td.tar.pr.dotted

        FOR    ${city}    IN    @{cities}
            ${text}=    Get Text    ${city}

            ${name_text}=    Get Text    css:td.dotted

            ${condition}=    Evaluate    '${text}' in ['Espoo', 'Helsinki', 'Vantaa']
            IF    ${condition}
                ${row_to_add}=    Create List    ${text}    ${name_text}
                Append Rows To Worksheet    ${row_to_add}
            END
        END
    END
    Save Workbook    Tapahtumat.xlsx
    Close All Browsers
