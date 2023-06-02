*** Settings ***
Documentation       Bot gets customer names from the Excel file and creates output work item per each row.

Library             Collections
Library             RPA.Excel.Files
Library             RPA.Robocorp.WorkItems
Library             RPA.Tables


*** Tasks ***
Produce items
    #Download Customers list    there are actions dedicated for Sharepoint and Onedrive
    Get customer data


*** Keywords ***
Get customer data
    Open Workbook    customers.xlsx
    ${table}=    Read Worksheet As Table    header=True
    FOR    ${row}    IN    @{table}
        ${variables}=    Create Dictionary
        ...    Country=${row}[Country]
        ...    Customer=${row}[Customer]
        Create Output Work Item
        ...    variables=${variables}
        ...    save=True
    END
