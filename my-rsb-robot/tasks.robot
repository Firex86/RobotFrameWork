*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF file.
Library    RPA.Browser.Selenium    auto_close=${FALSE}
Library    RPA.HTTP
Library    RPA.Excel.Files
Library    RPA.PDF

*** Keywords ***
Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/
Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass 
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the Excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=${True}


Fill and submit the form for one person
    [Arguments]    ${sales_reps}
    Input Text    firstname   ${sales_reps}[First Name] 
    Input Text    lastname    ${sales_reps}[Last Name] 
    Input Text    salesresult    ${sales_reps}[Sales] 
    Select From List By Value    salestarget    ${sales_reps}[Sales Target] 
    Click Button    Submit

Fill the form using the data from the Excel file
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=${True}
    Close Workbook    
    FOR    ${sales_reps}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_reps}
        
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png


Export the table as a PDF
    Wait Until Element Is Visible    id:sales-results
     ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
     Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Log out and close the browser
    Click Button    Log out
    Close Browser


*** Tasks ***
Insert the sales data for the week and export it as a PDF file.
    Open the intranet website
    Log in
    Download the Excel file
    Fill the form using the data from the Excel file
    Collect the results
    Export the table as a PDF
    [Teardown]    Log out and close the browser

