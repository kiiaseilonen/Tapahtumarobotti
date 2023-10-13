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
    Get events and write them to Excel
    [Teardown]    Close browser


*** Keywords ***
Open the event website
    Open Available Browser    http://www.koirat.com/tapahtumat

Create an Excel file with headers
    Create Workbook    ${OUTPUT_DIR}${/}Tapahtumat.xlsx
    Set Worksheet Value    1    1    Tapahtuma
    Set Worksheet Value    1    2    Päivämäärä
    Set Worksheet Value    1    3    Paikkakunta
    Set Worksheet Value    1    4    Linkki tapahtuman sivuille
    Save Workbook

Get events and write them to Excel
    Open Workbook    ${OUTPUT_DIR}${/}Tapahtumat.xlsx
    Element Should Be Visible    class:c528

    @{tables}=    Get WebElements    css:table.c528
    ${lastRow}=    Set Variable    2
    FOR    ${i}    IN RANGE    2
        ${table}=    Get WebElement    xpath:(//table[@class='c528'])[${i+1}]
        @{rows}=    Get WebElements    xpath:(//table[@class='c528'])[${i+1}]//tr[position()>1]
        ${rowCount}=    Get Length    ${rows}

        FOR    ${j}    IN RANGE    ${rowCount}
            ${event}=    Get Text    xpath:(//table[@class='c528'])[${i+1}]//tr[${j+1}]/td[1]
            ${date}=    Get Text    xpath:(//table[@class='c528'])[${i+1}]//tr[${j+1}]/td[2]
            ${city}=    Get Text    xpath:(//table[@class='c528'])[${i+1}]//tr[${j+1}]/td[3]
            ${href}=    Get Element Attribute    xpath:(//table[@class='c528'])[${i+1}]//tr[${j+1}]/td[3]/a    href

            IF    '${city}' in ['Espoo', 'Helsinki', 'Vantaa']
                Set Worksheet Value    ${lastRow}    1    ${event}
                Set Worksheet Value    ${lastRow}    2    ${date}
                Set Worksheet Value    ${lastRow}    3    ${city}
                Set Worksheet Value    ${lastRow}    4    ${href}

                ${lastRow}=    Evaluate    ${lastRow} + 1
                RPA.Excel.Files.Auto Size Columns    A    E    width=50
                RPA.Excel.Files.Auto Size Columns    B    C    width=20
            END
        END
    END

    Save Workbook    ${OUTPUT_DIR}${/}Tapahtumat.xlsx

Close browser
    Close All Browsers
