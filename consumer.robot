*** Settings ***
Library     Collections
Library     RPA.Browser.Selenium    auto_close=${False}
Library     RPA.Robocorp.WorkItems
Library     RPA.Tables
Library     RPA.Excel.Files
Library     OperatingSystem
Library     RPA.FileSystem
Library     DateTime
Library     RPA.Excel.Application
Library     Convert_Csv_To_Xlsx
Library     String
Library     RPA.Outlook.Application
Library     RPA.Robocorp.Vault
Library     Create_Pivot_Table
Library     Format_Table


*** Variables ***
${power_bi_url}
...                                     https://app.powerbi.com/groups/me/reports/04bfbfbd-d74e-4494-ae2b-9c993c4442d9?ctid=c315317a-5fcd-4010-9981-2f8020561f39&pbi_source=linkShare&bookmarkGuid=9387303e-d218-4b85-9f56-1a04d3a18d74
${signin_button}                        idSIButton9
${refresh_check_button}                 (//span[@class="itemName"])[1]
${report_button}                        (//span[@class="itemName"])[2]
${company_name_header}                  //div[@aria-colindex="2" and text()="Company Name"]
${refreshed_table}                      //div[@role="document" and @class="tableEx"]
${countries_header}                     //h3[@title="Location, Company Name, Resource Name"]
${company_name_filter}                  //div[contains(@aria-label, "Company Name is")]
${search_field}                         //input[@class="searchInput"]
${search_field_2}                       (//input[@class="searchInput"])[2]
${country_dropdown_arrow}               //div[@class="expandButton"]
${page_body}                            //div[@class="interactive-grid innerContainer"]
${clear_filter_button}                  //button[@aria-label="Clear filter"]
${summarized_data}                      (//label[@class="pbi-radio-button-internal"])[2]
${export_button}                        //button[@aria-label="Export"]
${refreshed_on}                         (//div[@role="gridcell"])[3]
${report_recipients}                    sdm.dwf@digitalworkforce.com    #malgorzata.worach@digitalworkforce.com    #michal.kowalczyk@digitalworkforce.com
${report_outdated_mail_recipient}       adrian.olszewski@digitalworkforce.com    #malgorzata.worach@digitalworkforce.com    #michal.kowalczyk@digitalworkforce.com
${report_outdated_cc_recipient}         sdm.dwf@digitalworkforce.com    #malgorzata.worach@digitalworkforce.com    #michal.kowalczyk@digitalworkforce.com


*** Tasks ***
Consume items
    [Documentation]    Login and then cycle through work items.
    TRY
        RPA.Outlook.Application.Open Application
        Open Power BI
        Login
        For Each Input Work Item    Handle item
    EXCEPT    AS    ${err}
        Log    ${err}    level=ERROR
        Release Input Work Item
        ...    state=FAILED
        ...    exception_type=APPLICATION
        ...    code=UNCAUGHT_ERROR
        ...    message=${err}
    FINALLY
        RPA.Outlook.Application.Quit Application
        Close All Browsers
    END


*** Keywords ***
Open Power BI
    Open Browser    ${power_bi_url}    chrome
    Maximize Browser Window
    Sleep    4s

Login
    ${vault}=    Get Secret    powerBI_credentials
    ${email_needed}=    Is Element Visible    email
    IF    ${email_needed} == ${TRUE}
        Input Text    email    ${vault}[username]
        Sleep    0.5s
        Click Button    submitBtn
        Sleep    4s
        Input Password    passwd    ${vault}[password]
        Click Button    ${signin_button}
        Sleep    3s
        Click Button    ${signin_button}
    END
    Sleep    7s

Handle item
    [Documentation]    Error handling around one work item.
    ${payload}=    Get Work Item Variables
    TRY
        Refresh check    ${payload}
        ${download_dir}=    Download report    ${payload}
        ${xlsx_file}=    Format report    ${payload}    ${download_dir}
        Send report    ${payload}    ${xlsx_file}
        Release Input Work Item    DONE
    EXCEPT    Element with locator*    type=GLOB    AS    ${err}
        ${error_message}=    Set Variable
        ...    Application error encountered: ${err}
        Log    ${error_message}    level=ERROR
        Release Input Work Item
        ...    state=FAILED
        ...    exception_type=APPLICATION
        ...    code=WEB_ELEMENT_NOT_FOUND
        ...    message=${error_message}
    EXCEPT    *!=*    type=GLOB    AS    ${err}
        ${error_message}=    Set Variable
        ...    Data for ${payload}[Customer] is outdated.
        Log    ${error_message}    level=ERROR
        Send outdated data notification    ${payload}
        Release Input Work Item
        ...    state=FAILED
        ...    exception_type=BUSINESS
        ...    code=DATA_OUTDATED
        ...    message=${error_message}
    EXCEPT    *timed out*    type=GLOB    AS    ${err}
        ${error_message}=    Set Variable
        ...    Application error encountered: ${err}
        Log    ${error_message}    level=ERROR
        Release Input Work Item
        ...    state=FAILED
        ...    exception_type=APPLICATION
        ...    code=TIMEOUT
        ...    message=${error_message}
    END

Refresh check
    [Arguments]    ${payload}
    Reload Page
    Sleep    3s
    Wait Until Element Is Visible    ${refresh_check_button}
    Click Element    ${refresh_check_button}
    Sleep    2s
    Click Element    ${refreshed_table}
    Sleep    2s
    Click Element    ${company_name_filter}
    Sleep    2s
    Wait Until Element Is Visible    ${search_field}
    Clear Element Text    ${search_field}
    Sleep    0.5s
    Input Text    ${search_field}    ${payload}[Customer]
    Sleep    2.5s
    Click Element    xpath://span[@title="${payload}[Customer]"]
    Sleep    2s
    ${refreshed_date}=    Get Text    ${refreshed_on}
    ${refreshed_date}=    Replace String    ${refreshed_date}    /    -
    ${month_list}=    Get Regexp Matches    ${refreshed_date}    ^\\d{1,2}
    ${month}=    Get From List    ${month_list}    0
    ${month}=    Convert To Integer    ${month}
    ${current_date}=    Get Current Date
    ${current_month}=    Convert Date    ${current_date}    %m
    ${current_month}=    Convert To Integer    ${current_month}
    Sleep    0.5s
    Click Button    ${clear_filter_button}
    Sleep    1s
    Click Element    ${company_name_filter}
    Sleep    0.5s
    Should Be Equal As Integers    ${month}    ${current_month}

Send outdated data notification
    [Arguments]    ${payload}
    ${email_body}=    Set Variable
    ...    <br>Hello, Adrian</br><br></br><br>Please be informed that the minutes report data in Power BI for customer ${payload}[Customer] is outdated.</br><br>Kindly investigate the issue and inform SDM team once fixed.</br><br></br><br>Please note that this email was generated by bot.</br>
    ${email_subject}=    Set Variable    Power BI: data for customer ${payload}[Customer] outdated
    Send Email
    ...    recipients=${report_outdated_mail_recipient} 
    ...    cc_recipients=${report_outdated_cc_recipient}
    ...    subject=${email_subject}
    ...    body=${email_body}
    ...    html_body=${True}

Download report
    [Arguments]    ${payload}
    ${user_home}=    Get Environment Variable    UserProfile
    ${download_dir}=    Set Variable    ${user_home}\\Downloads
    Click Element    ${report_button}
    Wait Until Element Is Visible    ${countries_header}    15s
    Click Element    ${countries_header}
    Sleep    2s
    Click Element    ${company_name_filter}
    Sleep    2s
    Wait Until Element Is Visible    ${search_field_2}
    Clear Element Text    ${search_field_2}
    Sleep    0.5s
    Input Text    ${search_field_2}    ${payload}[Customer]
    Sleep    2.5s
    Click Element    xpath://span[@title="${payload}[Customer]"]
    Sleep    2s
    Click Element    ${country_dropdown_arrow}
    Sleep    1s
    Click Element    xpath://span[@title="${payload}[Customer]"]
    Sleep    2s
    Click Element    ${page_body}
    Sleep    1s
    Press Keys    None    ALT+SHIFT+F10
    Sleep    1s
    Press Keys    None    LEFT
    Sleep    0.5s
    Press Keys    None    ENTER
    Sleep    1s
    Press Keys    None    DOWN
    Sleep    1s
    Press Keys    None    DOWN
    Sleep    0.5s
    Press Keys    None    ENTER
    Sleep    1.5s
    Click Element    ${summarized_data}
    Sleep    0.5s
    Click Button    ${export_button}
    IF    "${payload}[Customer]"=="Borealis AG"
        Sleep    20s
    ELSE    
        Sleep    15s
    END
    Click Element    xpath://span[@title="${payload}[Customer]"]
    Sleep    1s
    Click Element    ${country_dropdown_arrow}
    Sleep    0.5s
    Click Button    ${clear_filter_button}
    Sleep    1s
    Click Element    ${company_name_filter}
    Sleep    0.5s
    RETURN    ${download_dir}

Format report
    [Arguments]    ${payload}    ${download_dir}
    RPA.Excel.Files.Open Workbook    ${download_dir}${/}data.xlsx
    Delete Rows    1    2
    ${table}=    Read Worksheet As Table    header=${TRUE}
    Close Workbook
    RPA.FileSystem.Remove File    ${download_dir}${/}data.xlsx
    ${prev_month}=    Evaluate    datetime.datetime.now() - datetime.timedelta(days=30)
    ${prev_month_string}=    Convert Date    ${prev_month}    %B
    ${year}=    Convert Date    ${prev_month}    %Y
    ${year}=    Convert To Integer    ${year}
    Filter Table By Column    ${table}    End Date Time - Year    ==    ${year}
    Filter Table By Column    ${table}    End Date Time - Month    ==    ${prev_month_string}
    ${csv_file}=    Set Variable    ${payload}[Customer].csv
    ${csv_exist}=    Does File Exist    ${csv_file}
    IF    ${csv_exist}==${True}    RPA.FileSystem.Remove File    ${csv_file}
    Write table to CSV    ${table}    ${csv_file}
    ${xlsx_file}=    Set Variable    ${payload}[Customer]_Minutes_${prev_month_string}_${year}.xlsx
    Create Workbook
    Save Workbook    ${xlsx_file}
    Convert csv to xlsx    ${csv_file}    ${xlsx_file}
    Format Table    ${xlsx_file}
    Create Pivot Table    ${xlsx_file}
    RETURN    ${xlsx_file}

Send report
    [Arguments]    ${payload}    ${xlsx_file}
    ${message}=    Set Variable
    ...    <br>Hello, </br><br>Please find attached report for ${payload}[Customer].</br><br><br/><br>Please note that this email was generated by bot.</br>
    Send Email    ${report_recipients}    ${payload}[Customer] report    ${message}    ${TRUE}    ${xlsx_file}
