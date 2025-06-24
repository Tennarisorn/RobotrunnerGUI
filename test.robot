*** Settings ***
Library    ExcelHandler.py    C:/Users/t001404/OneDrive - The Siam Commercial Bank PCL/QA/RobotrunnerGUI/UDP_SIT_MVP3.xlsx    BD1
Library    SeleniumLibrary

*** Variables ***
${CSV_OUTPUT}      extracted_data.csv
${TEXT_XPATH}      //*[@id="rc-tabs-0-panel-details"]/div[2]/div/div[9]/div/span/div/span/span
${PARAM1_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[1]/span/label/span
${VALUE1_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[1]/div
${PARAM2_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[2]/span/label/span
${VALUE2_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[2]/div/span/span
${PARAM3_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[3]/span/label/span
${VALUE3_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[3]/div/span
${PARAM4_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[4]/span/label/span
${VALUE4_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[4]/div/span
${PARAM5_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[5]/span/label/span
${VALUE5_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[5]/div/span
${PARAM6_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[6]/span/label/span
${VALUE6_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[6]/div/span
${PARAM7_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[7]/span/label/span
${VALUE7_XPATH}    //*[@id="rc-tabs-0-panel-details"]/div[8]/div/div[7]/div/span
${NOTE_VALUE_XPATH}    //*[@id="app-root"]/div[4]/div/div/div[2]/div[2]/div/div[2]/div[1]/div[2]/div/div[1]/div[4]/div/div/div/div/div[2]/div/div/div[1]/div[1]/div/div[2]/div/div/div/div/span[2]

${BROWSER}         edge
${BASE_URL}        https://adb-1406497925929477.17.azuredatabricks.net/
${LOGIN_XPATH}     //*[@id="login-page"]/div/div/div[3]/a

*** Test Cases ***
Open Notebooks And Save StatusWithExtraColumns
    Open Browser    ${BASE_URL}    ${BROWSER}
    Maximize Browser Window
    Wait Until Element Is Visible    ${LOGIN_XPATH}    timeout=60s
    Click Element    ${LOGIN_XPATH}
    Sleep    30s

    ${total}=    Get Total Rows

    FOR    ${index}    IN RANGE    ${total}
        ${link}=    Get Notebook Link    ${index}
        Run Keyword If    '${link}' == '' or '${link}' == 'nan'    Continue For Loop

        Go To    ${link}
        # Sleep    10s
        Wait Until Element Is Visible    xpath=${TEXT_XPATH}    timeout=60s    error=status message error
        ${extracted_text}=    Get Text    xpath=${TEXT_XPATH}
        Log    Extracted status: ${extracted_text}
        Update Status    ${index}    ${extracted_text}
        Wait Until Element Is Visible    xpath=${PARAM1_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${VALUE1_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${PARAM2_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${VALUE2_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${PARAM3_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${VALUE3_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${PARAM4_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${VALUE4_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${PARAM5_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${VALUE5_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${PARAM6_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${VALUE6_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${PARAM7_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${VALUE7_XPATH}    timeout=10s
        Wait Until Element Is Visible    xpath=${NOTE_VALUE_XPATH}    timeout=10s
        # Params 1 to 7 extraction and add to CSV
        ${param1}=    Get Text    xpath=${PARAM1_XPATH}
        ${value1}=    Get Text    xpath=${VALUE1_XPATH}
        Log    Extracted param1: ${value1}
        Add Column With Value    ${index}    ${param1}    ${value1}

        ${param2}=    Get Text    xpath=${PARAM2_XPATH}
        ${value2}=    Get Text    xpath=${VALUE2_XPATH}
        Log    Extracted param2: ${value2}
        Add Column With Value    ${index}    ${param2}    ${value2}

        ${param3}=    Get Text    xpath=${PARAM3_XPATH}
        ${value3}=    Get Text    xpath=${VALUE3_XPATH}
        Log    Extracted param3: ${value3}
        Add Column With Value    ${index}    ${param3}    ${value3}

        ${param4}=    Get Text    xpath=${PARAM4_XPATH}
        ${value4}=    Get Text    xpath=${VALUE4_XPATH}
        Log    Extracted param4: ${value4}
        Add Column With Value    ${index}    ${param4}    ${value4}

        ${param5}=    Get Text    xpath=${PARAM5_XPATH}
        ${value5}=    Get Text    xpath=${VALUE5_XPATH}
        Log    Extracted param5: ${value5}
        Add Column With Value    ${index}    ${param5}    ${value5}

        ${param6}=    Get Text    xpath=${PARAM6_XPATH}
        ${value6}=    Get Text    xpath=${VALUE6_XPATH}
        Log    Extracted param6: ${value6}
        Add Column With Value    ${index}    ${param6}    ${value6}

        ${param7}=    Get Text    xpath=${PARAM7_XPATH}
        ${value7}=    Get Text    xpath=${VALUE7_XPATH}
        Log    Extracted param7: ${value7}
        Add Column With Value    ${index}    ${param7}    ${value7}

        ${note}=    Get Text    xpath=${NOTE_VALUE_XPATH}
        Log    Extracted note: ${note}
        Add Column With Value    ${index}    note    ${note}
        Export Row To CSV        ${index}
    END
    Save To CSV    ${CSV_OUTPUT}
    [Teardown]    Close All Browsers
