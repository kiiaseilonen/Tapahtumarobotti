*** Settings ***
Documentation       Get events from the website and save selected ones into an Excel file

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.Excel.Files
Library             Collections
Library             RPA.Browser.Playwright


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
    Open Workbook    Tapahtumat.xlsx
    Wait Until Element Is Visible    css:table.c528
    @{events}=    RPA.Browser.Selenium.Get WebElements    class:table.c528

    FOR    ${event}    IN    @{events}
        ${event_name_element}=    Get WebElement    ${event}    class:td.dotted
        ${event_name}=    RPA.Browser.Selenium.Get Text    ${event_name_element}
        Log    ${event_name}

        ${event_date_element}=    Get WebElement    ${event}    class:tac.dotted
        ${event_date}=    RPA.Browser.Selenium.Get Text    ${event_date_element}
        ${event_location_element}=    Get WebElement    ${event}    class:tar.pr.dotted
        ${event_location}=    RPA.Browser.Selenium.Get Text    ${event_location_element}
        ${event_url_element}=    Get WebElement    ${event}    a
        ${event_url}=    RPA.Browser.Selenium.Get Element Attribute    ${event_url_element}    href
        Log ${event_name}, {event_date},{event_location}

        IF    ${event_location} in ['Espoo', 'Helsinki', 'Vantaa']
            ${row}=    Create List    ${event_name}    ${event_date}    ${event_location}    ${event_url}
            Append Rows To Worksheet    ${row}
        END
    END

    Save Workbook    Tapahtumat.xlsx
    Close All Browsers
