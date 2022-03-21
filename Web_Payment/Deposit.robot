*** Settings ***
Library           SeleniumLibrary
Library           ExcellentLibrary
Library           DateTime
Library           OperatingSystem

*** Variables ***
${PW}             test
${Order_Info}     nation=${EMPTY}    ID=${EMPTY}    Price=${EMPTY}    Num=${EMPTY}    Etc=${EMPTY}

*** Test Cases ***
setup report Directory
    [Tags]    Set
    ${CurrentDate}    Get Current Date    result_format=%Y%m%d-%H%M%S
    ${logPath} =    Catenate    c:\\log\\${CurrentDate}
    Create Directory    ${logPath}
    set global variable    ${logPath}

Create Deposit ExcelFile
    [Tags]    Excel
    Create Workbook    ${logPath}\\Deposit_Result.xlsx
    ${time} =    Get Time
    Write to cell    A1    ${time}
    Write to cell    A2    법인명
    Write to cell    B2    계정명
    Write to cell    C2    주문 금액
    Write to cell    D2    주문 번호
    Write to cell    E2    취소 확인
    Write to cell    F2    비고
    Save
    Close Workbook

JPPayDeposit
    [Tags]    JPD
    #브라우저 열기 & 로그인
    Open Page    https://www.atomy.com/jp
    Login    testID    ${PW}

    #쇼핑몰에서 최저가격 상품 바로구매
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Click Element    xpath://*[@id="container"]/form/div[2]/ul[2]/li[4]/a    #낮은 가격순 정렬
    Click Element    xpath://*[@id="container"]/form/div[3]/div[1]/ul[1]/li[1]/a/img    #첫번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 600)
    Sleep    1
    Click Element    xpath://*[@id="bBuyNow"]    #바로구매 버튼 클릭

    #주문정보 입력
    Execute JavaScript    window.scrollTo(0, 800)
    input text    xpath://*[@id="tPhone1"]    010
    input text    xpath://*[@id="tPhone2"]    111
    input text    xpath://*[@id="tPhone3"]    111

    #배송정보 입력
    Click Element    xpath://*[@id="container"]/table[3]/tbody/tr[1]/td/label[1]    #주문자 정보와 동일
    input text    xpath://*[@id="tRevPhone1"]    010
    input text    xpath://*[@id="tRevPhone2"]    111
    input text    xpath://*[@id="tRevPhone3"]    1111
    input text    xpath://*[@id="tRevCellPhone1"]    010
    input text    xpath://*[@id="tRevCellPhone2"]    111
    input text    xpath://*[@id="tRevCellPhone3"]    1111
    sleep    1

    #결제정보 입력 & 결제하기
    Execute JavaScript    window.scrollTo(0, 1600)
    Click Element    xpath://*[@id="container"]/table[4]/tbody/tr/td/label[1]    # 지불방법 : 은행계좌 선택
    input text    xpath: //*[@id="txtIpgumName"]    ｱﾄﾐｼﾞｬﾊﾟﾝｺﾞｳﾄﾞｳｶﾞｲｼｬ    #입금자명 입력
    Click Element    xpath://*[@id="chkAgree"]    #동의 체크박스 클릭
    Click Element    xpath://*[@id="bPayment"]    #결제하기 클릭

    #결제화면 캡쳐 저장
    Capture Image & Save    300    JP_Deposit

    #결제정보 엑셀파일에 저장
    ${Order_Info.Num} =    Get Text    xpath://*[@id="container"]/ul[1]/li[1]/span[2]    #주문번호 가져오기
    ${Order_Info.Price} =    Get Text    xpath:/html/body/div[3]/div/ul[2]/li[4]/div    #주문금액 가져오기
    ${Order_Info.nation} =    Catenate    JP    #법인명 변수에 저장
    ${Order_Info.ID} =    Catenate    30802174    #계정명 변수에 저장
    ${Order_Info.Etc} =    Catenate    송금결제    #비고 변수에 저장
    Open Workbook    ${logPath}\\Deposit_Result.xlsx
    ${CNT} =    Get Row Count    #엑셀 열 갯수 읽기
    Set Global Variable    ${CNT}    #Save Cancel to Excel 키워드에서 사용하기 위해 전역변수 선언
    Write to cell    A${CNT+1}    ${Order_Info.nation}    #법인명 입력
    Write to cell    B${CNT+1}    ${Order_Info.ID}    #계정명 입력
    Write to cell    C${CNT+1}    ${Order_Info.Price}    #주문금액 입력
    Write to cell    D${CNT+1}    ${Order_Info.Num}    #주문번호 입력
    Write to cell    F${CNT+1}    ${Order_Info.Etc}    #결제방법 입력
    Save
    Close Workbook

    #상세페이지에서 주문확인 및 주문취소
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/div/ul/li[2]/a    #나의 애터미몰 클릭
    Execute JavaScript    window.scrollTo(0, 500)
    sleep    1
    ${Inc_Num} =    Set Variable    1
    FOR    ${i}    IN RANGE    10    #10번 반복
        ${Cf_Num} =    Get Text    xpath://*[@id="tbdList"]/tr[${Inc_Num}]/td[2]/a    #최상단부터 주문번호 비교
        Exit For Loop If    ${Cf_Num} == ${Order_Info.Num}    #비교번호가 주문번호가 같다면 루프 Exit
        Log    ${Cf_Num}
        ${Inc_Num} =    Evaluate    ${Inc_Num}+2    #다르다면 Inc_Num + 2(주문번호 List가 홀수번임)
    END
    Click Element    xpath://*[@id="tbdList"]/tr[${Inc_Num}]/td[7]/p[2]/a    # 주문취소 버튼 클릭
    sleep    2
    Handle Alert    ACCEPT    1min    # 주문취소하시겠습니까? 네
    Handle Alert    ACCEPT    1min    # 주문취소가 완료되었습니다. 확인

    #주문취소 여부 엑셀파일에 저장
    Execute JavaScript    window.scrollTo(0, 500)
    Open Workbook    ${logPath}\\Deposit_Result.xlsx
    ${check} =    Get Text    xpath://*[@id="tbdList"]/tr[${Inc_Num}]/td[7]/p    #주문현황 가져오기
    Run Keyword If    '${check}'=='注文取消し'    Run Keyword    Write to Cell    E${CNT+1}    취소 확인
    ...    ELSE    Run Keyword    Write to Cell    E${CNT+1}    취소 실패
    Save
    Close workbook

JPPayEasy
    [Tags]    JPP
    #브라우저 열기 & 로그인
    Open Browser    https://www.atomy.com/jp    chrome
    maximize browser window
    sleep    1
    Click Element    xpath://*[@id="mainNotictPopupView"]/div[2]/button    #팝업 닫기
    Click Element    xpath://*[@id="header"]/div[2]/div/ul/li[2]/a    #로그인 클릭
    input text    id:userId    testID
    input text    id:userPw    test
    Click Element    xpath://*[@id="frm"]/div/div/div[1]/p[1]/a    # 로그인 완료

    #쇼핑몰에서 최저가격 상품 바로구매
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Click Element    xpath://*[@id="container"]/form/div[2]/ul[2]/li[4]/a    #낮은 가격순 정렬
    Click Element    xpath://*[@id="container"]/form/div[3]/div[1]/ul[3]/li[1]/a/img    #3번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 600)
    Sleep    1
    Click Element    xpath://*[@id="bBuyNow"]    #바로구매 버튼 클릭

    #주문정보 입력
    Execute JavaScript    window.scrollTo(0, 800)
    input text    xpath://*[@id="tPhone1"]    010
    input text    xpath://*[@id="tPhone2"]    111
    input text    xpath://*[@id="tPhone3"]    111

    #배송정보 입력
    Click Element    xpath://*[@id="container"]/table[3]/tbody/tr[1]/td/label[1]    #주문자 정보와 동일
    input text    xpath://*[@id="tRevPhone1"]    010
    input text    xpath://*[@id="tRevPhone2"]    111
    input text    xpath://*[@id="tRevPhone3"]    1111
    input text    xpath://*[@id="tRevCellPhone1"]    010
    input text    xpath://*[@id="tRevCellPhone2"]    111
    input text    xpath://*[@id="tRevCellPhone3"]    1111
    sleep    1

    #결제정보 입력 & 결제하기
    Execute JavaScript    window.scrollTo(0, 1600)
    Click Element    xpath://*[@id="container"]/table[4]/tbody/tr/td/label[4]    # 지불방법 : 페이이지 선택
    Click Element    xpath://*[@id="chkAgree"]    #동의 체크박스 클릭
    Click Element    xpath://*[@id="bPayment"]    #결제하기 클릭
    sleep    2

    #페이이지 Window에서 애터미 Window로 Switch
    ${handles}    Get Window Handles
    switch window    ${handles}[1]    ${EMPTY}
    Close window    #페이이지 window 닫기
    sleep    1
    switch window    ${handles}[0]
    Execute JavaScript    window.scrollTo(0, 300)
    sleep    1

    #결제화면 캡쳐 저장
    Capture Image & Save    300    JP_Payeasy

    #결제정보 엑셀파일에 저장
    Save Inform to Excel    JP    testID   페이이지

    #상세페이지에서 주문정보 확인
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/div/ul/li[2]/a    #나의 애터미몰 클릭
    #페이이지는 주문취소 없음

PHPayDeposit
    [Tags]    PH
    #브라우저 열기 & 로그인
    Open Page    https://www.atomy.com/ph
    Login    test    ${PW}

    #쇼핑몰에서 최저가격 상품 바로구매
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Click Element    xpath://*[@id="container"]/form/div[2]/ul[2]/li[5]/a    #낮은 가격순 정렬
    Click Element    xpath://*[@id="container"]/form/div[3]/div[1]/ul[2]/li[1]/a/img    #2번째 상품 클릭(1번째 상품 품절)
    Execute JavaScript    window.scrollTo(0, 600)
    sleep    1
    Click Element    xpath://*[@id="bBuyNow"]    #바로구매 버튼 클릭

    #주문정보 입력
    Execute JavaScript    window.scrollTo(0, 800)
    input text    xpath://*[@id="tEmail1"]    test
    input text    xpath://*[@id="tEmail2"]    gmail.com
    input text    xpath://*[@id="tSendName"]    test

    #배송정보 입력
    Execute JavaScript    window.scrollTo(0, 1200)    #배송지 옵션 : new address(default)
    sleep    1
    input text    xpath://*[@id="tRevUserName"]    test
    input text    xpath://*[@id="tRevCellPhone"]    0101111111
    input text    xpath://*[@id="tRevAddr1"]    test
    Click Element    xpath://*[@id="bSearchCityList"]    #우편번호 찾기 클릭
    sleep    1
    Select From List By Value    id:selRegion2    ARMM
    sleep    1
    Select From List By Value    id:selLocation2    Basilan
    sleep    1
    Click Element    xpath://*[@id="tbdZipCodeList"]/tr[1]/td/a
    input text    xpath://*[@id="tRevPostNo"]    p306
    sleep    1
    Click Element    xpath: //*[@id="bPayment2"]    #계속하기 클릭

    #결제정보 입력 & 결제하기
    Execute JavaScript    window.scrollTo(0, 800)
    Click Element    xpath://*[@id="container"]/table[3]/tbody/tr/td/label[2]    # 지불방법 : 무통장입금
    Click Element    xpath://*[@id="chkAgree"]    #동의하기 클릭
    Click Element    xpath://*[@id="bPayment"]    #결제하기 클릭

    #결제화면 캡쳐 저장
    Capture Image & Save    600    PH_Deposit

    #결제정보 엑셀파일에 저장
    ${Order_Info.Num} =    Get Text    xpath://*[@id="container"]/div[5]/h3/span    #주문번호 가져오기
    ${Order_Info.Price} =    Get Text    xpath:/html/body/div[3]/div/ul[2]/li[4]/div    #주문금액 가져오기
    ${Order_Info.nation} =    Catenate    PH    #법인명 변수에 저장
    ${Order_Info.ID} =    Catenate    testID   #계정명 변수에 저장
    Set Global Variable    ${Order_Info}    #JP법인에서 사용하기 위해 전역변수 선언
    Open Workbook    ${logPath}\\Deposit_Result.xlsx
    ${CNT} =    Get Row Count    #엑셀 열 갯수 읽기
    Set Global Variable    ${CNT}    #Save Cancel to Excel 키워드에서 사용하기 위해 전역변수 선언
    Write to cell    A${CNT+1}    ${Order_Info.nation}    #법인명 입력
    Write to cell    B${CNT+1}    ${Order_Info.ID}    #계정명 입력
    Write to cell    C${CNT+1}    ${Order_Info.Price}    #주문금액 입력
    Write to cell    D${CNT+1}    ${Order_Info.Num}    #주문번호 입력
    Save
    Close Workbook

    #상세페이지에서 주문확인 및 주문취소
    Cancel Order

    #주문취소 여부 엑셀파일에 저장
    Save Cancel to Excel    Order Cancellation

IDPayDeposit
    [Tags]    ID
    #브라우저 열기 & 로그인
    Open Page    https://www.atomy.com/id
    Login    testID    ${PW}

    #쇼핑몰에서 최저가격 상품 바로구매
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Click Element    xpath://*[@id="container"]/form/div[2]/ul[2]/li[4]/a    #낮은 가격순 정렬
    Click Element    xpath://*[@id="container"]/form/div[3]/div[1]/ul[1]/li[1]/a/img    #첫번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 600)
    Click Element    xpath://*[@id="bBuyNow"]    #바로구매 버튼 클릭

    #주문정보 입력
    Execute JavaScript    window.scrollTo(0, 800)
    input text    xpath://*[@id="tSendName"]    test

    #배송정보 입력
    Click Element    xpath://*[@id="container"]/table[3]/tbody/tr[1]/td/label[1]    #배송지 옵션 : 주문자와 동일
    Execute JavaScript    window.scrollTo(0, 1500)

    #결제정보 입력 & 결제하기
    Click Element    xpath://*[@id="container"]/table[4]/tbody/tr/td/label[2]    #지불 방법 : 가상 계좌
    Click Element    xpath://*[@id="chkAgree"]
    Click Element    xpath://*[@id="bCalculate"]    #주문 확인 클릭
    Click Element    xpath://*[@id="bPayment"]    #결제하기 클릭
    Sleep    8
    Select Frame    id:ifrPG    #새프레임으로 이동
    execute javascript    document.querySelector("#optionsCENA").click();    #은행사 선택
    execute javascript    document.querySelector("#submit_button").click();    #결제하기 선택
    Unselect Frame
    sleep    1

    #결제화면 캡쳐 저장
    Capture Image & Save    400    ID_Deposit

    #결제정보 엑셀파일에 저장
    Save Inform to Excel    ID    testID

    #상세페이지에서 주문확인 & 주문취소
    Cancel Order
    
    #주문취소 여부 엑셀파일에 저장
    Save Cancel to Excel    Pesanan Dibatalkan

*** Keywords ***
Open Page
    [Arguments]    ${url}
    #브라우저 열기
    Open Browser    ${url}    chrome
    maximize browser window
    sleep    1
    Click Element    xpath://*[@id="mainNotictPopupView"]/div[2]/button    #팝업 닫기
    sleep    1
    #로그인

Login
    [Arguments]    ${ID}    ${pW}
    Click Element    xpath://*[@id="header"]/div[2]/div/ul/li[2]/a    #로그인 클릭
    input text    id:userId    ${ID}
    input text    id:userPw    ${pW}
    Click Element    xpath://*[@id="frm"]/div/div/div[1]/p[1]/a    # 로그인완료

Go Shoppingmall & BuyNow
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Click Element    xpath://*[@id="container"]/form/div[2]/ul[2]/li[3]/a    #낮은 가격순 정렬
    Click Element    xpath://*[@id="container"]/form/div[3]/div[1]/ul[1]/li[1]/a/img    #첫번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 600)
    Click Element    xpath://*[@id="bBuyNow"]    #바로구매 버튼 클릭
    

Capture Image & Save
#캡쳐 이미지 원하는 경로에 저장
    [Arguments]    ${length}    ${imgName}
    sleep    1
    Execute JavaScript    window.scrollTo(0, ${length})
    sleep    3
    Set Screenshot Directory    ${logPath}
    Capture Page Screenshot    ${imgName}.png
   

Save Inform to Excel
#결제정보 엑셀파일에 저장
    [Arguments]    ${nation}    ${ID}    ${method}= None
    ${Order_Info.Num} =    Get Text    xpath://*[@id="container"]/ul[1]/li[1]/span[2]    #주문번호 가져오기
    ${Order_Info.Price} =    Get Text    xpath:/html/body/div[3]/div/ul[2]/li[4]/div    #주문금액 가져오기
    ${Order_Info.nation} =    Catenate    ${nation}    #법인명 변수에 저장
    ${Order_Info.ID} =    Catenate    ${ID}    #계정명 변수에 저장
    ${Order_Info.Etc} =    Catenate    ${method}    #비고 변수에 저장
    Set Global Variable    ${Order_Info}    #JP법인에서 사용하기 위해 전역변수 선언
    Open Workbook    ${logPath}\\Deposit_Result.xlsx
    ${CNT} =    Get Row Count    #엑셀 열 갯수 읽기
    Set Global Variable    ${CNT}    #Save Cancel to Excel 키워드에서 사용하기 위해 전역변수 선언
    Write to cell    A${CNT+1}    ${Order_Info.nation}    #법인명 입력
    Write to cell    B${CNT+1}    ${Order_Info.ID}    #계정명 입력
    Write to cell    C${CNT+1}    ${Order_Info.Price}    #주문금액 입력
    Write to cell    D${CNT+1}    ${Order_Info.Num}    #주문번호 입력
    Write to cell    F${CNT+1}    ${Order_Info.Etc}    #결제방법 입력
    Save
    Close Workbook
    

Cancel Order
#상세페이지에서 주문 취소버튼 클릭
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/div/ul/li[2]/a    #나의 애터미몰 클릭
    Execute JavaScript    window.scrollTo(0, 500)
    sleep    1
    Click Element    xpath://*[@id="tbdList"]/tr[1]/td[7]/p[2]/a    # 주문취소 버튼 클릭
    Handle Alert    ACCEPT    1min    # 주문취소하시겠습니까? 네
    Handle Alert    ACCEPT    1min    # 주문취소가 완료되었습니다. 확인
    

Save Cancel to Excel
#주문취소 여부 엑셀파일에 저장
    [Arguments]    ${Check_str}
    ${Check} =    Get Text    xpath://*[@id="tbdList"]/tr[1]/td[7]/p    #주문현황 가져오기
    Open Workbook    ${logPath}\\Deposit_Result.xlsx
    Run Keyword If    '${Check}'=='${Check_str}'    Run Keyword    Write to Cell    E${CNT+1}    취소 확인
    ...    ELSE    Run Keyword    Write to Cell    E${CNT+1}    취소 실패    #주문취소 문구와 비교
    Save
    Close workbook