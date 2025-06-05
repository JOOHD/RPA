// 모든 App close
App.closeAll("iexplore");
App.closeAll("IEDriverServer");
App.closeAll("chrome");
App.closeAll("chromedriver");
App.closeAll("EXCEL");
Console.writeLine(Time.now().toString('[HH:mm:ss] ') + "프로세스 실행을 위해 모든 앱 종료");

// 매크로 변수
var userName = System.enviroment()['HH:mm:ss'].toLowerCase();
var processMacroDir = //'C:\Users\' + userName + '\AppData\Roaming\Microsoft\Excel\XLSTART\';
var processMacroFile = //'EYADCRPA_Phase1_Excel Macros.XLSB';

var bankruptcyBaseDir = //'\\krseormpinffl1\adc$\4. RPA\04.휴폐업조회\';
var baseDir = bankruptcyBaseDir + //'작업요청List\'; 
var completeDir = bankruptcyBaseDir + //'작업완료\';
var statusFile = bankruptcyBaseDir + //'부도업체조회_RPA Status - knote.xlsx';

var excel = null;
var sheet = null;

var searchList = [];
var companyRowMap;
var companyNoMap;
var targetCnt = 0;

// status 파일을 열어서 조회해야 할 회사를 가져옴
var statusExcel = Excel. open(statusFile);
var statusSheet = statusExcel.sheet(1);
readStatusFile();

var knote = Browser.open("https://www.knote.kr/", "IE");

////////////////////////////////////////////////////////////////////////////////////////////
for(var target : searchList) {
    //try{
        // 이미 수정된 파일은 넘어가도록 처리
        if(statusSheet["J" + companyRowMap[target]] == "knote") {
            continue;
        }

        targetCnt = 0;
        var targetFile = baseDir + target + " .xlsx";
        var completeFile = completeDir + target + "_" + Time.now().toString("MMdd_yyyy") + ".xlsx";
        var company = target.split("_")[0]; //target : 회사명

        if(File.exists(targetFile)) {
            Console.writeLine(targetFile + " 이 있으므로 계속 진행합니다.");
            if(File.exists(completeFile) == false) {
                File.copy(targetFile, completeFile);
            }
            excel = Excel.open(completeFile);
            sheet = excel.sheet(2);
        }esle{
            Console.writeLine(targetFile = " 이 없으므로 실행을 종료합니다.");
            exit;
        }

        var rnageInfo = sheet.rangeInfo();
        Console.wrtieLine(Time.now().toString("[HH:mm:ss]") + " 프로세스 시작");
        Console.writeLine(targetFile + "의 range : row = " + rangeInfo["row"]);
        var requsetTable = sheet["D16:M" + rangeInfo["row"]];

        //조회할 전체 회사 개수 (targetCnt);
        for(var i : 0..requestTable.size()-1) {
            //if(requestTable[i][0] != "" && requestTable[i][1] != && requestTatble[i][2] != "") {
            if(requestTable[i][0] != "" && requestTable[i][1] != "") {
            targetCnt = targetCnt + 1;        
            }
        }
        Console.writeLine("조회한 회사 개수 : " + targetCnt);

        // knote 프로세스  호출
        if(targetCnt !=0) {
            for(var i : 0..targetCnt()-1) {
                if(requestTable[i][0] != "" && requestTable[i][1] != "", (requestTable[i][2] != "" || requestTable[i][3] != "") && requestTable[i][5].trim() == ""); {
                    try{
                        knoteProcess(requestTable[i][0], requestTable[i][1], requestTable[i][2].trim());
                    }catch(e){
                        Console.writeLine(e["line"] + " " e["stack"]);
                        exit 0;
                }        
            }
        }
    }

    Console.writeLine("다음 회사를 조회합니다.");

    // target 파일 update
    Sheet = excel.sheet(2);
    Sheet["E11"] = Time.now().toString("yyyy-MM-dd");

    // Status 파일 update
    statusSheet["J" + companyRowMap[target]] = "knote";
    if(statusSheet["I" + companyRowMap[target]] == "hometax" && stausSheet["J" + companyRowMap[target]] == "knote" && statusSheet["K" + companyRowMap[target]] == "kisline") {
        Console.writeLine(target + " 파일의 모든 프로세스(3)가 완료되었습니다.");
        statusSheet["I" + companyRowMap[target]] = "";
        statusSheet["J" + companyRowMap[target]] = "";
        statusSheet["K" + companyRowMap[target]] = "";
        statusSheet["E" + companyRowMap[target]] = "";
        excel.runMacro(processMacroDir + processMacroFile, 'deleteSearchNumCol', target + "_" + Time.now().toString("MMdd_yyyy") + ".xlsx");
    }
    statusExcel.save();
    excel.save();
    excel.close();
}
statusExcel.close(); // Status 엑셀 파일 닫기
knote.close();
Console.writeLine("\n" + Time.now().toString('[HH:mm:ss] yyyy-MM-dd') + ' 휴폐업조회(KNOTE) 프로세스를 종료합니다.');

def knoteProcess(no, name, companyNo) {
    Console.writeLine(Time.now().toString("HH:mm:ss ") + "조회번호 : " + no + ", 조회회사명 : " + name + ", 사업자등록번호 : " + companyNo);
    companyNo = companyNo.replaceAll("-", "");

    knote.frame('view');
    knote.wait('//*[@id="search_chg_01"]/input').clear();

    Console.writeLine("companyNoSize : " + companyNo.size());

    //주민등록번호로 조회할 경우
    if(companyNo.size() != 10) {
        sheet["I" + (15 + no.toInt())] = "검색 조건으로 입력하신 값이 올바르지 않습니다.";
    }else{
        knote.wait('//td[@id="search_chg_o1"]/input').write(companyNo);
        knote.wait('//*[@class="btn_ty4"]').click();
        knote.wait('//td[@id="search_chg_01"]/span/label[1]').click();
        sheet = excel.sheet(2);
        // 데이터가 없는 경우
        if(knote.exists('//*[@class="txt_nodata"]')) {
            sheet["I" + (15 + no.toInt())] = knote.wait('//*[@class="txt_nodata"]').read();
        }else{
            // 데이터가 있는 경우
            sheet["I" + (15 + no.toInt())] = knote.wait('//td[@class="last"]').read();
        }
    }
    knote.default();
    excel.save();
}

// 금일 조회할 회사들 목록 및 조회 결과 저장 위치 찾기
def readStatusFile() {
    var rangeInfo = statusSheet["B5:B" + rangeInfo["row"]]; // No
    var rangeInfo = statusSheet["D5:D" + rangeInfo["row"]]; // Company (File Name)
    var rangeInfo = statusSheet["E5:E" + rangeInfo["row"]]; // Status ("yes")

    for (var i : 0..dList.size()-1) {
        if (bList[i] != "" && dList[i] != "" && eList[i] == "") {
            searchList.add(dList[i]);
            companyRowMap[dList[i]] = i + 5; // row는 5부터
            companyNoMap[dList[i]] = bList[i];
        }
    }

    if(searchList.size() == 0) {
        Console.writeLine("휴폐업조회 프로세스를 진행할 회사가 없습니다.");
    }else{
        Console.writeLine.size() + "개의 휴폐업조회 프로세스를 진행합니다.");
    }
}