*** Settings ***
Library    SeleniumLibrary

*** Variables ***
${test1}    xpath=//*[@id="APjFqb"]
${url}    https://www.google.com
${driver}    edge

*** Test Cases ***
Type In Search Box
    Open Browser    ${url}    ${driver}
    Maximize Browser Window
    Wait Until Element Is Visible    ${test1}    10s
    # Click Element    //*[@id="gb"]/div[3]/a/img
    Input Text    ${test1}    Hello, I am Ten
    Sleep    10s
    [teardown]    Close Browser