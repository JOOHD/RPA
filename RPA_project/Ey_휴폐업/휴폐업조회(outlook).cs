Console.writeLine("\n" + Time.now().toString('[HH:mm:ss] yyyy-MM-dd') + "휴폐업조회 메일링 프로세스를 시작합니다.");

//	모든 App close
App.closeAll("iexplore");
App.closeAll("IEDriverServer");
App.closeAll("chrome");
App.closeAll("chromedriver");
App.closeAll("EXCEL");
Console.writeLine(Time.now().toString('[HH:mm:ss] ') + "프로세스 실행을 위해 모든 앱 종료");

// 04.휴폐업조회 메일링
var bankruptcyBaseDir = // '\\krseormpinffl1\adc$\4. RPA\04.휴폐업조회\';
var baseDir = bankruptcyBaseDir;
var completeDir = bankruptcyBaseDir + // '작업완료\';
var imgDir = // '.\region\';
var statusFile = bankruptcyBaseDir + // '부도업체조회_RPA Status.xlsx';
var msgFile = baseDir + // 'bankruptcy_mail_template.msg';

var excel = null;
var sheet = null;

var searchList = [];
var companyRowMap;
var companyNomap;
var targetCnt = 0;
var fileName;

// status 파일을 열어서 조회해야 할 회사를 가져옴
var statusExcel = Excel.open(stausFile);
var statusSheet = statusExcel.sheet(1);
readStatusFile();

/////////////////////////////////////////////////////////////////////////////////////////
var outlook = outlook.folder();
for(var target : searchList) {
    targetCnt = targetCnt + 1;

    // 파일생성일자 status 파일에 추가
    var completeFile = completeDir + target + "_" + Time.now().toString("MMdd_yyyy") + ".xlsx";
    var company = target.split("_")[0]; // target : 회사명
    Console.writeLine(Time.now().toString("[HH:mm:ss]") + company + "(" + targetCnt + "/" + searchList.size() + ")" + " 메일 전송을 시작합니다.");
    
    if(File.exists(completeFile) {
        Console.writeLine(Time.now().toString("[HH:mm:ss]") + completeFile + "이 있으므로 계속 진행합니다.");
        excel = Excel.open(comleteFile);
        sheet = excel.sheet(2);
    } else {
        Console.writeLine(Time.now().toString("[HH:mm:ss]") + completeFile + "이 없으므로 다음 회사로 넘어갑니다.");
        continue;
    }

    // 메일 프로세스
    var mailAddList = sheet["E8:E10"]; // excel에서 추출
    excel.save();
    excel.Close();

    if(mailProcess(outlook, company, mailAddList, completeFile)) {
        // Status 파일 update
        statusSheet["G" + companyRowMap[target]] = Time.now().toString("yyyy-MM-dd");
        statusSheet["E" + companyRowMap[target]] = // "yes";
        statusExcel.save();
    }
}

statusExcel.close(); // Status 엑셀 파일 닫기
Console.writeLine(Time.now().toString('[HH:mm:ss] yyyy-MM-dd') + "휴폐업조회 메일링 프로세스를 종료합니다.");

// 메일 보내기 프로세스
def mailProcess(outlook, company, mailAddrList, completeFile) {
    var msg;
    
    // msg 파일 open
    if(File.exists(msgFile)) {
        Console.writeLine(Time.now().toString('[HH:mm:ss]') + msgFile + "이 있으므로 계속 진행합니다.");
        msg = outlook.open(msgFile);
    } else {
        Console.writeLine(Time.now().toString('[HH:mm:ss]') + msgFile + "이 없으므로 실행을 종료합니다.");
        exit -1;
    }

    // 메일앱 포커스
    var title = "Message (HTML) ";
    if(App.exists(title, 30000)) {
        App.maximize(title);
        App.focus(title);
        Console.writeLine(Time.now().toString('[HH:mm:ss]') + "MSG 파일이 열렸습니다.");
    } else {
        Console.writeLine(Time.now().toString('[HH:mm:ss]') + "MSG 파일이 열리지 않았습나다.");
        eixt -1;
    }

    msg.attach(completeFile);
    msg.to(mailAddrList[0] + ";" + mailAddrList[1] + ";" + mailAddrList[2] + ";");
    msg.subject("[ADC] RPA_휴폐업 및 당좌거래정지 조회_" + company);
    msg.CC.('Ey.Adc1@kr.ey.com');

    msg.send();
}

// 금일 조회할 회사들 목록 및 조회 결과 저장 위치 찾기
def readStatusFile() {
    var rangeInfo = statusSheet.rangeInfo();

    var bList = statusSheet["B5:B" + rangeInfo["row"]]; // no
    var bList = statusSheet["D5:D" + rangeInfo["row"]]; // Company (File Name)
    var bList = statusSheet["B5:E" + rangeInfo["row"]]; // Status ("complete")

    for (var i : 0..dList.size()-1) {
        if(bList[i] != "" && dList[i] != "" && eList[i] == "complete") {
            searchList.add(dList[i]);
            companyRowMap[dList[i]] = i + 5; // row는 5부터
            companyNoMap[dList[i] = bList[i]];
        }
    }

    if(searchList.size() == 0) {
        Console.writeLine(Time.now().toString("[HH:mm:ss]") + "휴폐업조회 메일링 프로세스를 진행할 회사가 없습니다.") 
    }else{
        Console.writeLine(Time.now().toString("[HH:mm:ss]") + searchList.size() + "개의 휴폐업조회 메일링 프로세스를 진행합니다.")
    }
}