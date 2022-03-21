#tags 활용하여 run: robot -i Set -i Excel -i TR -i JP CreditCard.robot
*** Settings ***
Library           SeleniumLibrary
Library           ExcellentLibrary
Library           DateTime
Library           OperatingSystem

*** Variables ***
${PW}             test
${Order_Info}    nation=${EMPTY}    ID=${EMPTY}    Price=${EMPTY}    Num=${EMPTY}    Etc=${EMPTY}

*** Test Cases ***
setup report Directory
    [Tags]    Set
    ${CurrentDate}    Get Current Date    result_format=%Y%m%d-%H%M%S
    ${logPath} =    Catenate    c:\\log\\${CurrentDate}
    Create Directory    ${logPath}
    set global variable    ${logPath}

Create Card ExcelFile
    [Tags]    Excel
    Create Workbook    ${logPath}\\Card_Result.xlsx
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

HKPayPal
    [Tags]    HKP
    #브라우저 열기 & 로그인
    Open Page    https://www.atomy.com/hk 
    Login    testID    ${PW}
    sleep    2

    #쇼핑몰에서 최저가격 상품 바로구매
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Execute JavaScript    window.scrollTo(0, 800)
    sleep    1
    Click Element    xpath://*[@id="container"]/form/div[2]/ul/li[4]/a    #낮은 가격순 정렬
    Click Element    xpath://*[@id="container"]/form/div[3]/div[1]/ul[1]/li[1]/a/img    #첫번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 600)
    sleep    1
    Click Element    xpath://*[@id="bBuyNow"]    #바로구매 버튼 클릭
    sleep    1

    #주문정보 입력
    Execute JavaScript    window.scrollTo(0, 800)
    input text    xpath://*[@id="tSendName"]    test

    #배송정보 입력
    Click Element    xpath://*[@id="container"]/table[3]/tbody/tr[1]/td/label[1]    #배송지 옵션 : 주문자 정보와 동일
    Click Element    xpath://*[@id="bPayment2"]    # 다음
    Execute JavaScript    window.scrollTo(0, 800)
    sleep    1
    Click Element    xpath://*[@id="settleGubun3"]    #PayPal선택
    Click Element    xpath://*[@id="chkAgree"]    #동의 체크박스 클릭
    Click Element    xpath://*[@id="bPayment"]    #결제하기 클릭
    sleep    5
    input text    xpath://*[@id="cardNumber"]    0000 0000 0000 0000    #카드번호 입력
    input text    xpath://*[@id="cardExpiry"]    0124    #유효날짜
    input text    xpath://*[@id="cardCvv"]    000    #cvc
    input text    xpath://*[@id="firstName"]    test    #이름
    input text    xpath://*[@id="lastName"]    a    #성
    Execute JavaScript    window.scrollTo(0, 400)
    input text    xpath://*[@id="billingLine1"]    深水埗區    #주소입력
    input text    xpath://*[@id="phone"]    0101111111   #휴대전화번호 입력
    input text    xpath://*[@id="email"]    test@gmail.com    #이메일
    Execute JavaScript    window.scrollTo(0, 800)
    Click Element    xpath://*[@id="root"]/div/div[1]/main/div[1]/form/section[3]/div[2]/div/label    #paypal 가입 안함 선택
    Click Element    xpath://*[@id="root"]/div/div[1]/main/div[1]/form/section[3]/div[2]/fieldset/div/div/div/label    #비회원 동의 체크박스 클릭
    Click Element    xpath://*[@id="root"]/div/div[1]/main/div[1]/form/div/button    #카드 결제하기 버튼
    sleep    20

    #결제화면 캡쳐 저장
    Capture Image & Save    400    HK_Paypal

    #결제정보 엑셀파일에 저장
    Save Inform to Excel    HK    testID    paypal

    #상세페이지에서 주문확인 및 주문 취소
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/div/ul/li[1]/a
    Execute JavaScript    windows.scrollTo(0, 600)
    Click Element    xpath://*[@id="tbdList"]/tr[1]/td[7]/p[2]/a
    sleep    2
    Handle Alert    ACCEPT    1min    # 주문취소하시겠습니까? 네
    sleep    1
    Handle Alert    ACCEPT    1min    # 주문취소가 완료되었습니다. 확인
    sleep    1

    Save Cancel to Excel    取消購買
    #실결제됨

UKCreditCard
    [Tags]    UK
    #브라우저 열기 & 로그인
    Open Page    https://www.atomy.uk/uk/home 
    Login    testID   ${PW}

    #쇼핑몰에서 최저가격 상품 바로구매
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    #쇼핑몰 클릭
    Execute JavaScript    window.scrollTo(0, 800)
    sleep    1
    Click Element    xpath:/html/body/div[3]/form/div[2]/ul/li[4]/a    #낮은 가격순 정렬
    Click Element    xpath:/html/body/div[3]/form/div[3]/div[1]/ul[1]/li[1]/a/img    #첫번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 400)
    sleep    1
    Click Element    xpath://*[@id="bBuyNow"]
    sleep    2

    #주문정보 입력
    Execute JavaScript    window.scrollTo(0, 800)
    sleep    1
    input text    xpath://*[@id="tSendName"]    atomy

    #배송정보 입력
    Click Element    xpath://*[@id="runInfo0"]    #배송지 옵션 : 주문자와 동일
    sleep    1
    Execute JavaScript    window.scrollTo(0, 1800)
    sleep    1
    Click Element    xpath://*[@id="bPayment2"]
    sleep    1

    #결제정보 입력 & 결제하기
    Execute JavaScript    window.scrollTo(0, document.body.scrollHeight)
    sleep    1
    Click Element    xpath://*[@id="chkAgree"]    #동의하기
    Click Element    xpath://*[@id="bPayment"]    #결제하기
    sleep    5
    Switch Window    new    #결제창 window로 switch
    input text    xpath://*[@id="email"]    test@gmail.com    #신용카드창 이메일
    input text    xpath://*[@id="cardNumber"]    0000 0000 0000 0000
    input text    xpath://*[@id="cardExpiry"]    01 / 24
    input text    xpath://*[@id="cardCvc"]    000
    input text    xpath://*[@id="billingName"]    atomy
    Execute JavaScript    window.scrollTo(0, 400)
    Click Element    xpath://*[@id="root"]/div/div/div[2]/div/div[2]/form/div[2]/div[2]/button/div[3]    #결제하기
    sleep    4

    #결제화면 캡쳐 저장
    Switch Window    main    #애터미화면으로 switch
    Capture Image & Save    400    UK_Card

    #결제정보 엑셀파일에 저장
    Save Inform to Excel    UK    testID   

    #상세페이지에서 주문확인 & 주문취소
    Cancel Order

    #주문취소 여부 엑셀파일에 저장
    Save Cancel to Excel    Payment Cancelled
    #실결제됨(우리카드창 이동안함)

TWCreditCard
    [Tags]    TW
    #브라우저 열기 & 로그인
    Open Page    https://www.atomy.com/tw
    Login    testID    ${PW}

    #쇼핑몰에서 최저가격 상품 바로구매
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Execute JavaScript    window.scrollTo(0, 800)
    sleep    1
    Click Element    xpath://*[@id="container"]/form/div[2]/ul/li[4]/a    #낮은 가격순 정렬
    Click Element    xpath://*[@id="container"]/form/div[3]/div[1]/ul[1]/li[1]/a/img    #첫번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 600)
    sleep    1
    Click Element    xpath://*[@id="bBuyNow"]    #바로구매 버튼 클릭
    sleep    1

    #주문정보 입력
    Execute JavaScript    window.scrollTo(0, 800)
    input text    xpath://*[@id="tCellPhone"]    0101111111
    input text    xpath://*[@id="tSendName"]    atomy

    #배송정보 입력
    Select From List By Value    id:tDeliTime    2    #배송시간 선택 : 14時~18時
    input text    xpath://*[@id="tRevUserName"]    atomy
    input text    xpath://*[@id="tRevCellPhone"]    0101111111
    input text    xpath://*[@id="tRevPhone1"]    010
    input text    xpath://*[@id="tRevPhone2"]    11111111
    Click Element    xpath://*[@id="bSearchPost"]    #우편번호 찾기 클릭
    sleep    1
    Select From List By Value    id:selSido    基隆市
    sleep    1
    Select From List By Value    id:selGuname    201
    input text    xpath://*[@id="tRevAddr2"]    台北市中正區

    #결제정보 입력 & 결제하기
    Execute JavaScript    window.scrollTo(0, 1000)
    sleep    1
    Click Element    xpath://*[@id="settleGubun3"]    #카드결제 선택
    Click Element    xpath://*[@id="chkAgree"]    #동의하기
    Click Element    xpath://*[@id="bPayment"]    #결제하기 클릭
    sleep    5
    input text    xpath://*[@id="ctl00_ContentPlaceHolder1_strCardNo"]    0000000000000000    #카드번호 입력
    Select From List By Value    id:ctl00_ContentPlaceHolder1_strMM    01    #유효달
    Select From List By Value    id:ctl00_ContentPlaceHolder1_strYY    2024    #유효년도
    input text    xpath://*[@id="check_num"]    000    #cvc
    Click element    xpath://*[@id="ctl00_ContentPlaceHolder1_btn_box"]    #결제하기
    #****************이후 카드비밀번호(20152015) 입력 수동*******************
    sleep    40

    #결제화면 캡쳐 저장
    Capture Image & Save    300    TW_Card

    #결제정보 엑셀파일에 저장
    Save Inform to Excel    TW    testID

    #상세페이지에서 주문확인 & 주문취소
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/div/ul/li[1]/a    #나의 애터미몰 클릭
    Execute JavaScript    window.scrollTo(0, 500)
    sleep    1
    Click Element    xpath://*[@id="tbdList"]/tr[1]/td[7]/p[2]/a    # 주문취소 버튼 클릭
    sleep    2
    Handle Alert    ACCEPT    1min    # 주문취소하시겠습니까? 네
    Handle Alert    ACCEPT    1min    # 주문취소가 완료되었습니다. 확인

    #주문취소 여부 엑셀파일에 저장
    Save Cancel to Excel    訂單取消

*** Keywords ***
#브라우저 열기
Open Page
    [Arguments]    ${url}
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

#쇼핑몰에서 최저가격 상품 찾은 후 바로구매 클릭
Go Shoppingmall & BuyNow
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/a    # 쇼핑몰 버튼 클릭
    Execute JavaScript    window.scrollTo(0, 800)
    sleep    1
    Click Element    xpath:/html/body/div[3]/form/div[2]/ul[2]/li[4]/a    #낮은 가격순 정렬
    sleep    2
    Click Element    xpath:/html/body/div[3]/form/div[3]/div[1]/ul[1]/li[1]/a/img    #첫번째 상품 클릭
    Execute JavaScript    window.scrollTo(0, 600)
    Click element    xpath://*[@id="bBuyNow"]
    sleep    1

#캡쳐 이미지 원하는 경로에 저장
Capture Image & Save
    [Arguments]    ${length}    ${imgName}
    sleep    1
    Execute JavaScript    window.scrollTo(0, ${length})
    sleep    3
    Set Screenshot Directory    ${logPath}
    Capture Page Screenshot    ${imgName}.png

#결제정보 엑셀파일에 저장
Save Inform to Excel
    [Arguments]    ${nation}    ${ID}    ${method}= None
    ${Order_Info.Num} =    Get Text    xpath://*[@id="container"]/ul[1]/li[1]/span[2]    #주문번호 가져오기
    ${Order_Info.Price} =    Get Text    xpath:/html/body/div[3]/div/ul[2]/li[4]/div     #주문금액 가져오기
    ${Order_Info.nation} =    Catenate    ${nation}    #법인명 가져오기
    ${Order_Info.ID} =    Catenate    ${ID}            #계정명 가져오기
    ${Order_Info.Etc} =    Catenate    ${method}       #비고 가져오기
    Open Workbook    ${logPath}\\Card_Result.xlsx
    ${CNT} =    Get Row Count                          #엑셀 열 갯수 읽기
    Set Global Variable    ${CNT}                      #Save Cancel to Excel 키워드에서 사용하기 위해 전역변수 선언
    Write to cell    A${CNT+1}    ${Order_Info.nation}    #법인명 입력
    Write to cell    B${CNT+1}    ${Order_Info.ID}        #계정명 입력
    Write to cell    C${CNT+1}    ${Order_Info.Price}     #주문금액 입력
    Write to cell    D${CNT+1}    ${Order_Info.Num}       #주문번호 입력
    Write to cell    F${CNT+1}    ${Order_Info.Etc}       #결제방법 입력
    Save
    Close Workbook

#상세페이지에서 주문 취소버튼 클릭
Cancel Order
    Click Element    xpath://*[@id="gnb"]/ul/li[1]/div/ul/li[2]/a    #나의 애터미몰 클릭
    Execute JavaScript    window.scrollTo(0, 500)
    sleep    1
    Click Element    xpath://*[@id="tbdList"]/tr[1]/td[7]/p[2]/a    # 주문취소 버튼 클릭
    Handle Alert    ACCEPT    1min    # 주문취소하시겠습니까? 네
    Handle Alert    ACCEPT    1min    # 주문취소가 완료되었습니다. 확인

#주문취소 여부 엑셀파일에 저장
Save Cancel to Excel
    [Arguments]    ${Check_str}
    ${Check} =    Get Text    xpath://*[@id="tbdList"]/tr[1]/td[7]/p
    Open Workbook    ${logPath}\\Card_Result.xlsx
    Run Keyword If    '${Check}'=='${Check_str}'    Run Keyword    Write to Cell    E${CNT+1}    취소 확인
    ...    ELSE    Run Keyword    Write to Cell    E${CNT+1}    취소 실패
    Save
    Close workbook