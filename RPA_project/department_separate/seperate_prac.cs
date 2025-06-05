//	모든 App close
App.closeAll("iexplore");
App.closeAll("IEDriverServer");
App.closeAll("chrome");
App.closeAll("chromedriver");
App.closeAll("EXCEL");
Console.writeLine(Time.now().toString('[HH:mm:ss] ') + "프로세스 실행을 위해 모든 앱 종료");

//input 엑셀 파일 경로
var dir = 'C:\RPA\과제 자료\파일 분류 과제\input\';
//output 폴더 경로
var dir2 = 'C:\RPA\과제 자료\파일 분류 과제\output\';

//output 부서 폴더 
var SL_ASU = dir2 + 'ASU';
var SL_CON = dir2 + 'CON';
var SL_SaT = dir2 + 'SaT';
var SL_TAX = dir2 + 'TAX';

//폴더 생성
Dir.create(SL_ASU);
Dir.create(SL_CON);
Dir.create(SL_SaT);
Dir.create(SL_TAX);

//header 작성
var header = [
			["MK", "SL", "Invoice Number", "Eng Partner Name", "Eng Manager Name", "Eng Name", "Eng Client Number", "Eng Client Name", "Date Billed", "Due date", "Aging of inv. date", "Duedate여부", "AR_1106"]
			];

//input엑셀 파일 열기 
var excel = Excel.open(dir + 'AR Aging report_1106.xlsx');

//sheet 펼치기
var sheet = excel.sheet(1);

//sheet에 rangeInfo 적용
var range = sheet.rangeInfo(true);

//다차원 배열 이용
var data_table = sheet["A3:" + range["column"] + range["row"]]; //엑셀 전체 범위
var slList = sheet["B3:B" + range["row"]]; //부서
var pnList = sheet["D3:D" + range["row"]]; //파트너 이름

//Console.writeLine(slList);

//////////////////// 문자열로 비교 ///////////////////////////////////
// 새 부서리스트 생성
var slList2=[];
// 기존 부서리스트 반복문
for(var i:slList) {	
	if(slList2.size() == 0) { //새 부서리스트에 비교할 값이 없으므로 초기값 세팅
		slList2.add(i); //i = slList의 첫번째 인덱스
	} else {// i가 담기면 size는 1이기 때문에 else로 넘어가게 된다.
		var flag=false; //부서가 같지 않다
		for (var j:slList2) { //새부서리스트 반복문
			if(j==i){ // 부서값 비교
//				count = count + 1;
				flag = true; // 부서가 같다
				break; //새 부서리스트 반복문 끝
			}
		}
//		if(count == 0)
		if(!flag) { //부서가 같지 않다 
			slList2.add(i); //새 부서리스트에 추가
		}
	}	
}

//Console.writeLine(slList2);
var exRow; //초기화 해주어야 다음 sheet에서 처음부터 데이터 쌓임.
slList2=["ASU", "CON", "SaT", "TAX"];
for(var i:slList2) {
	//엑셀 생성
	var newExcel = Excel.new(dir2 + i);	//output파일 경로에 생성
	var newSheet = newExcel.sheet(1);
	
	exRow = newSheet.rangeInfo(true)["row"];
	
	for(var j: 0..data_table.size()-1){
		if (data_table[j][1]==i) { //전체 엑셀 부서와 부서리스트 부서 비교
			//제목 입력
			newSheet['A1'] = header;
			
			exRow += 1;
			//'A2'(header = 'A1')를 적게 되면 2배씩 중간에 공백이 생김 'A'라고 적어야 첫줄부터 기입
			newSheet['A' + exRow]=data_table[j];

			//시트 정렬
			newSheet.columnWidth('A:N');
		}
	}
	newExcel.saveAs(dir2 + i);
}

//////////////////// 인덱스으로 비교 //////////////////////////////
//var slList2=[];
//for(var i:0..slList.size()-1) {	
//	if(slList2.size() == 0) {
//		slList2.add(slList[i]);
//	} else {
//		var count=0;
//		for (var j:0..slList2.size()-1) {
//			if(slList2[j]==slList[i]){
//				count = count + 1;
//				break;
//			} else {
//				
//			}
//		}
//		if (count == 0) {
//			slList2.add(slList[i]);
//		}
//	}	
//}
////////////////////BD_news 예쩨///////////////////////////////////
//var previewOut = 0;
//if(previewList.size() == 0){
//	previewList.add(preview);
//} else {
//	for(var i:0 .. previewList.size()-1){
//		if(previewList[i].contains(preview)){
//			Console.writeLine("중복된 기사입니다.");
//			previewOut = previewOut + 1;
//			break;
//		} else {
//			//previewOut = false;	
//		}
//	}
//	
//	if(previewOut == 0){
//		previewList.add(preview);
//	}
//}




