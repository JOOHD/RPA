//	모든 App close
App.closeAll("iexplore");
App.closeAll("IEDriverServer");
App.closeAll("chrome");
App.closeAll("chromedriver");
App.closeAll("EXCEL");
Console.writeLine(Time.now().toString('[HH:mm:ss] ') + "프로세스 실행을 위해 모든 앱 종료");

// 매크로 변수
var userName = System.enviroment()["username"].toLowerCase();
var processMacroDir = // 'C:Users\' +  userName + '\AppData\Roaming\Microsoft\Excel\XLSTART\';
var processMacroFile = // 'EYADCRPA_Phase1_Excel Macros.XLSB';  

var bankruptcyBaseDir = // '\\krseormpinffl1\adc$\4. RPA\04.휴폐업조회\';
var baseDir = bankruptcyBaseDir + // '작업요청List\';
var completeDir = bankruptcyBaseDir + // '작업완료\'
var imgBaseDir = bankruptcyBaseDir + // 'imges';
var statusFile = bankruptcyBaseDir + // '부도업체조회_RPA Status.xlsx';
var today = Time.now().toString("yyyyMMdd");

var excel = null;
var sheet = null;

var searchList = [];
var companyRowMap;
var companyNoMap;
var targetCnt = 0;

// status 파일을 열어서 조회해아 할 회사를 가져옴.
var statusExcel = Excel.open(stautsFile);
var statusSheet = statusExcel.sheet(1);
readStatusFile();

var kisline = Browser.open('kisline.com', 'IE');
// 2022.02.23(화) 정인기 보안스크립트 종료하는 코드
try{
    for(var i : 0..5){
        kisline.alert("Accept");
        Console.writeLine("Security Alert 발생.");
        System.sleep(1000);
    }

}catch(e){
    Console.writeLine("Security Alert 발생하지 않았음.");
}
Console.writeLine("시큐리티 창이 없습니다. 로그인을 진행하겠습니다.");

var id = kisline.wait('//*[@id="lgnuid"]'); // 로그인
id.write($kislineId);
id.type("Tab");
kisline.find('//*[@type="password]').write($kislinePwd); // 특히 이런 비번은 별도의 변수로 관리할 수 있어야 함.
kisline.find('//*[@id="loginForm/div/div/fieldset/a"]').click();

// 아래 EY한영 입력은 이후 검사할 법인들을 동일한 창에서 검색할 수 있도록 하기 위함.
try{
    kisline.wait('//*[@id="q"][@name="searchValue"]', 30000).write("EY한영")
}catch(e){
    Console.writeLine(e[stack]);
    exit(1);
}
kislilne.find('//*[@id="searchView"][@class="btn_search_total"]').click();
Console.writeLine("환경 설정 완료!");

////////////////////////////////////////////////////////////////////////////////////////////////
for(var target : searchList) {
    if(statusSheet["K" + companyRowMap[target]] == "kisline") {
        continue;
    }

    targetCnt = 0;
    var targetFile = baseDir + target + ".xlsx";
    var completeFile = completeDir + target + "_" + Time.now().toString("MMdd_yyyy") + ".xlsx";
    var company = target.split("_")[0]; // target : 회사명

    if(File.exists(targetFile)) {
        Console.writeLine(targetFile + " 이 있으므로 계속 진행합니다.");
        if(File.exists(completeFile) == false) {
            File.copy(targetFile, completeFile);
        }
        excel = Excel.open(completeFile);
        sheet = excel.sheet(2);
    } else {
        Console.writeLine(targetFile + " 이 없으므로 실행을 종료합니다.");
        exit;
    }

    var rangeInfo = sheet.rangeInfo();
    Console.writeLine(Time.now().toString("HH:mm:ss  ") + " 프로세스 시작");
    Console.writeLine(targetFile + "의 range : row = " + rangeInfo["row"]);
    var requestTable = sheet["D16:M" + rangeInfo["row"]];

    // 조회할 전체 회사 개수 (targetCnt)
    for(var i : 0..requestTable.size()-1) {
        if(requestTable[i][0] != "" && requestTable[i][1] != "") {
            targetCnt = targetCnt + 1;
        }
    }
    Console.writeLine("조회할 회사 개수 : " + targetCnt);

    // kisline 프로세스 호출
    if(Dir.exists(imgBaseDir + company)) {
        Console.writeLine(imgBaseDir + company + " : 캡쳐이미지 디렉토리가 이미 있으므로 그것을 사용합니다.");
    }else{
        Dir.create(imgBaseDir + company);
    }

    if(targetCnt != 0){
        for(var i : 0..targetCnt -1 ) {
            if(reqeustTable[i][0] != "" && requestTable[i][1] != "" && (requestTable[i][2] != "" || requestTable[i][3] != "") && requestTable[i][8].trim() == "") {
                kislineProcess(requestTable[i][0], requestTable[i][1], requestTable[i][2]);
            }
        }

    }   

    Console.writeLine("다음 회사를 조회합니다.");

    // target 파일 update
    sheet = excel.sheet(2);
    sheet["E11"] = Time.now().toString("yyyy-MM-dd");

    excel.save();

    statusSheet["K" + companyRowMap[target]] = "kisline";
    if(statusSheet["I" + companyRowMap[target]] == "hometax" && statusSheet["J" + companyRowMap[target]] == "knote" && statusSheet["K" + companyRowMap[target]] == "kisline") {
        Console.writeLine(target + " 파일의 모든 프로세스(3)가 완료되었습니다.");
        statusSheet["I" + companyRowMap[target]] = "";
        statusSheet["J" + companyRowMap[target]] = "";
        statusSheet["K" + companyRowMap[target]] = "";
        statusSheet["E" + companyRowMap[target]] = "complete";
        excel.runMacro(processMacroDir + processMacroFile, 'deleteSearchNumCol', target + "_" + Time.now().toString("MMdd_yyyy") + ".xlsx");
    }
    excel.save();
    excel.close();
    statusExcel.save();
}

statusExcel.close(); // Status 엑셀 파일 닫기
kisline.wait('//*[@id="header]/div[i]/ul/li[3]/a', 30000).click(); // kisline 로그아웃
kisline.close();
Console.writeLine("\n" + Time.now().toString('[HH:mm:ss] yyyy-MM-dd') + " 휴폐업조회(KISLINE) 프로세스를 종료합니다.");

def kislineProcess(no, name, companyNo) {
    Console.writeLine(Time.now().toString("HH:mm:ss") + "조회번호 : " + no + ", 조회 회사명 : " + name +  ", 사업자등록번호 : " + companyNo);

    // 엑셀 시트 이미 있는지 체크
    try{
        sheet = excel.sheet(no); // 시트가 없으면 에러
    } catch(e) {
        // kisline 홈페이지 활성화
        var imgDir = // 'Z:\Peon\Bankruptcy\images\';

        // 사업자 번호를 입력 (깨끗이 지운 후)하고 검색 버튼을 클릭
        var searchField = kisline.wait('//*[@id="q][@name="searchValue"]');
        searchField.Clear();
        searchField.write(companyNo);
        System.sleep(4000);
        kisline.find('//*[@id=searchView][@class="btn_search_total]').click();

        // 기업명 클릭해서 상세 페이지로 이동
        // 검색결과 건수가 0건일 때 사항
        // 검색결과 첫번째에 검색어 제안 : 다른 검색어를 찾을 수 없습니다. 문구 떴을 때  체크(오른쪽사항)
        if(!kisline.exists('//div[@id="cont"]/div[2]/h3/em', 5000) || kisline.exists('//div[@id="cont"]/div[1]/dl[1]/dt[1]', 5000)) {
            sheet = excel.sheet(2);
            sheet["J" + (15 + no.toInt())] = "조회결과가 없습니다."; // 기업평가
            sheet["K" + (15 + no.toInt())] = "조회결과가 없습니다.";
            sheet["L" + (15 + no.toInt())] = "조회결과가 없습니다.";
            continue;
        }

        if(kisline.exists('//*[@id="eprTable"]/tbody/tr/td[2]/a', 5000)) {
            var companyName = kisline.wait('//*[@id="eprTable"]/tbody/tr/td[2]/a', 30000);

            // companyName.click() 부분이 실행 되지 않고 넘어갈 때가 있어 while 문으로 묶어줌
            var checkCount = 0;
            while(kisline.exists('//*[@id="eprTable"]/tbody/tr/td[2]/a', 3000)){
                companyName.Click();
                checkCount = checkCount + 1;
                if(checkCount == 10){
                    break;
                }
            }

            // RPA 조회내역 등급값 입력, ROW는 16부터 시작
            sheet = excel.sheet(2);
            sheet["J" + (15 + no.Int())] = kisline.wait('기업평가_xpath').read().trim(); // 기업평가
            sheet["J" + (15 + no.Int())] = kisline.wait('현금흐름_xpath').read().trim(); // 현금흐름 (cash flow)
            sheet["J" + (15 + no.Int())] = kisline.wait('watch_xpath').read().split('기준일')[0].trim(); // Watch

            // screenCapture 선택한 회사만 스크린 캡처 수행
            if(sheet["M" + (15 + no.toInt())].trim() == "") {

                // 등급 현황을 찾은 후 거기까지 scroll 하기
                var grade = kisline.wait('//div[@class="overView_area03"]');

                var width = grade.size()["width"];
                var height = grade.size()["height"];
                kisline.saveImage(imgBaseDir + companyNo + '_' + today + ".png", grade.location()["x"], 0, width, height);
                Console.writeLine(Time.now().toString("HH:mm:ss") + "이미지 저장 : " + imgBaseDir + companyNo + '_' + today + ".png");

                excel.addSheet(no);
                sheet = excel.sheet(no);

                sheet["B1"] = "ADC";
                sheet["B2"] = "KISLINE 등급현황 Screen Capture";

                sheet["C4"] = ["value":"고객번호", "border":"true"];
                sheet["C5"] = ["value":"고객명", "border":"true"];
                sheet["C6"] = ["value":"사업자등록번호", "border":"true"];

                sheet.["C4"] = ["value":no, "border":"true"];
                sheet.["C4"] = ["value":name, "border":"true"];
                sheet.["C4"] = ["value":companyNo, "border":"true"];

                sheet.columWidth("C");
                sheet.columWidth("D");

                if(File.exists(imgBaseDir + companyNo +  '_' + today + ".png")) {
                    sheet.addImage(imgBaseDir + companyNo + '_' + today + ".png", 96, 140, 760, 302);
                }else{
                    Console.writeLine("  " + name + "회사의" + companyNo + '_' + today + "파일이 없습니다.");
                }
                sheet = excel.sheet(2);
                sheet["M" + (15 + no.toInt())] = "0"; // no 시트 추가하고 스크린캡처 완료하면 status screenCapture 칼럼에 0 표시
            }

        } else {
            // 검색 결과가 없는 경우
            Console.writeLine("조회 결과가 없습니다.");
            sheet = excel.sheet(2);
            var noResult = kisline.wait('//table[@id="eprTable"]/tbody/tr/td').read().trim();
            sheet["J" + (15 + no.toInt())] = "조회 결과가 없습니다."; //기업평가등급
            sheet["K" + (15 + no.toInt())] = "조회 결과가 없습니다."; //현금흐름등급
            sheet["L" + (15 + no.toInt())] = "조회 결과가 없습니다."; //watch등급

            excel.addSheet(no);
            sheet = excel.sheet(no);

            sheet["B1"] = "ADC";
            sheet["B2"] = "KISLINE 등급현황 Screen Capture";

            sheet["C4"] = ["value":"고객번호", "border":"true"];
            sheet["C5"] = ["value":"고객명", "border":"true"];
            sheet["C6"] = ["value":"사업자등록번호", "border":"true"];

            sheet.["C4"] = ["value":no, "border":"true"];
            sheet.["C4"] = ["value":name, "border":"true"];
            sheet.["C4"] = ["value":companyNo, "border":"true"];

            sheet.columWidth("C");
            sheet.columWidth("D");
        }

        excel.save();
        Console.writeLine(Time.now().toString("HH:mm:ss") + "엑셀 추가 완료 : " + no);

    } finally {
        Console.writeLine(Time.now().toString("HH:mm:ss") + "진행 상황 : " + no + " / " + targetCnt);
        Console.writeLine("");
    }
}

//  금일 조회할 회사들 목록 및 조회 결과 저장 위치 찾기
def readStatusFile() {
    var rangeInfo = statusSheet.rangeInfo();

    var bList = statusSheet["B5:B" + rangeInfo["row"]]; // No
    var dList = statusSheet["D5:D" + rangeInfo["row"]]; // Company (File Name)
    var eList = statusSheet["E5:E" + rangeInfo["row"]]; // status ("yes")

    for(var i : 0..dList.size() -1) {
        if(bList[i] != "" && dList[i] != "" && eList[i] == "") {
            searchList.add(dList[i]);
            companyRowMap[dList[i]] = i + 5; // row는 5부터
            companyNoMap[dList[i]] = bList[i]; 
        }
    }

    if(searchList.size() == 0) {
        Console.writeLine("kisline 조회 프로세스를 진행할 회사가 없습니다.")
    } else {
        Console.writeLine(searchList.size() + "개의 kisline 프로세스를 진행합니다.")
    }

}







































