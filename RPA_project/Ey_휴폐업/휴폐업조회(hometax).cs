package 휴폐업;
//	모든 App close
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

var statusFile = bankruptcyBaseDir + //'부도업체조회_RPA Status - hometax.xlsx';
var today = Time.now().toString("yyyyMMdd");

var excel = null;
var sheet = null;

var searchList = [];
var companyRowMap;
var companyNoMap;
var targetCnt = 0;

// status 파일을 열어서 조회해야 할 회사를 가져옴
var statusExcel = Excel.open(stausFile);
var statusSheet = statusExcel.sheet(1);
readStatusFile();
var hometax = "";
var count = 1;
while(count<=5) {
    try{
        hometax = Browser.open("hometax", "Chrome");
        System.sleep(5000);
        count = 0;
        try{
            hometax.wait('//*[@id="group1300"]').click(); //'조회/발급'
            System.sleep(5000);
            count = 0;
        }catch(e){
            Console.writeLine(e["line"] + e["stack"]);
            App.closeAll("iexplore");
            App.closeAll("IEDriverServer");
            count = coutn + 1;
            Console.writeLine("휴폐업 조회 Hometax를 종료합니다. 재시도 횟수 5회 초과");
            exit(0);
        }
    }catch(e){
        Console.writeLine(e["line"] + e["stack"]);
        App.closeAll("iexplore");
        App.closeAll("IEDriverServer");
        count = count + 1;
        Console.writeLine("휴폐업 조회 Hometax를 종료합니다. 재시도 횟수 5회 초과");
        exit(0);
    }
}
while(count<= 5) {
    try{
        hometax.frame(txppIframe);
        break;
    }catch(e) {
        Console.writeLine(e["line"] + e["stack"]);
        count = count + 1;
        if(count >=5) {
            Console.writeLine("휴폐업 조회 Hometax를 종료합니다. 재시도 횟수 5회 초과");
        }
    }
}

// '사업자등록번호로 조회' 메뉴 2020-11-12(목) 수정완료
hometax.wait('//*[@id=""]').jClick();
///////////////////////////////////////////////////////////////////////////////////////////////
for(var target : searchList) {
    //이미 수행된 파일은 넘어가도록 처리
    if(statusSheet["I" + companyRowMap[target]] == "hometax") {
        continue;
    }

    targetCnt = 0;
    var targetFile = baseDir + target + ".xlsx";
    var company = target.split("_")[0]; // target : 회사명

    if(File.exists(targetFile)) {
        Console.writeLine(Time.now().toString("[HH:mm:ss]") + targetFile + " 이 있으므로 계속 진행합니다.");
        if(File.exists(completeFile) == false) {
            File.copy(targetFile, completeFile);
        } 
        excel = Excel.open(completeFile);
        sheet = excel.sheet(2);
    }else{
        Console.writeLine(Time.now().toString("[HH:mm:ss]") + targetFile + " 이 없으므로 실행을 종료합니다.");
        exit(0);      
    }

    var rangeInfo = sheet.rangeInfo();
    Console.writeLine(Time.now().toString("[HH:mm:ss]") + "프로세스 시작");
    Console.writeLine(Time.now().toString("[HH:mm:ss]") + targetFile + "의 range : row =" + rangeInfo["row"]);
    var requestTable = sheet["D16:M" + rangeInfo["row"]];

    //조회할 전체 회사 개수 (targetCnt)
    for(var i : 0..requestTable.size()-1) {
        if(requestTable[i][0] != "" && requestTable[i][1] != "") {
            targetCnt = targetCnt + 1;
        }
    }

    Console.writeLine(Tinme.now().toString("[HH:mm:ss") + "조회할 회사 개수 :" + targetCnt);

    // hometax 프로세스 호출

    if(targetCnt !=0) {
        for(var i : 0..targetCnt -1){ 
            if(requestTable[i][0] != "" && requestTable[i][1] != "" && (requestTable[i][2]) != "" || requestTable[i][3] != "") && requestTable[i][4].trim() == "" {
                hometaxProcess(requestTable[i][0], requestTable[i][1], requestTable[i][2].trim(), requestTable[i][3]);
            }
        }     
    }

    Console.writeLine(Time.now().toString("[HH:mm:ss]") + "다음 회사를 조회합니다.");

    //target 파일 update
    sheet = excel.sheet(2);
    sheet["E11"] = Time.now().toString("yyyy-MM-dd");
 
    //status 파일 update
    statusSheet["I" + companyRowMap[target]] = "hometax";
    if(statusSheet["I" + companyRowMap[target]] == "hometax" && statusSheet["J" + companyRowMap[target]] == "konte" && statusSheet["K" + companyRowMap[target]] == "kisline") {
        Console.writeLine(target + " 파일의 모든 프로세스(3)가 완료되었습니다.");
        statusSheet["I" + companyRowMap[target]] = "";
        statusSheet["J" + companyRowMap[target]] = "";
        statusSheet["K" + companyRowMap[target]] = "";
        statusSheet["E" + companyRowMap[target]] = "Completee";
        excel.runMacro(processMacroDir + processMacroFile, 'deleteSearchNumCol', target + "_" + Time.now().toString("MMdd_yyyy") + ".xlsx");
    }
    statusExcel.save();
    excel.save();
    excel.Close();
}
statusExcel.close(); // Status 엑셀 파일 닫기
hometax.close();
Console.writeLine("\n" + Time.now().toString('[HH:mm:ss] yyyy-MM-dd') + " 휴페업조회(HOMETAX) 프로세스를 종료합니다,");

def hometaxProcess(no, name, companyNo, rrn) {
    Console.writeLine(Time.now().toString("[HH:mm:ss]") + "조회번호 : " + no + ", 조회 회사명 :" + name + ", 사업자등록번호 : " + companyNo + ", 주민등록번호 : " + rrn);
    sheet = excel.sheet(2);
    if(companyNo.trim() != "") {
        var companyNoSize = companyNo.replaceAll("-", "").size();
        if(companyNoSize != 10) {
            sheet["H" + (15 + no.toInt())] = "번호형식이 맞지 않습니다.";
            continue;
        }
        hometax.wait('//*[@id="bsno"]', 30000).clear();
        hometax.wait('//*[@id="bsno"]', 30000).write(companyNo);
        hometax.wait('//*[@id="triggers"]', 30000).click();
        count = 0;
        while(true) {
            if(count < 6){
                try{
                    sheet["H" + (15 + no.toInt())] = hometax.wait('').read();
                    break;
                }catch(e){
                    Console.writeLine(e["line"] + e["stack"]);
                    count = count + 1;
                }
            }else{
                Console.writeLine(count + "번째 시도까지 실패 중단합니다.");
                exit(1);
            }
        }
        System.sleep(7000); //바로 조회 불가        
        excel.save();
    }else{
      pause;
      //사업자등록번호대신 주민등록번호 쓰여있는 경우        
    }
}
//금일 조회할 회사들 목록 및 조회 결과 저장 위치 찾기
def readStatusFile() {
    var rangeInfo = statusSheet.rangeInfo();

    var bList = statusSheet["B5:B" + rangeInfo["row"]]; //No
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
        Console.writeLine.size() + ("개의 휴폐업조회 프로세스를 진행합니다.");
    }
} 






































