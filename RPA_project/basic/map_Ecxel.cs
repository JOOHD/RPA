//try{
//	excelEx();
//	}catch(e){
//		Console.writeLine(e["stack"]); // 에러내용
//		Console.writeLine(e["line"]); // 에러부분
	//에러났을때 처리방안	
//	}


def excelEx(){
	var dir = 'C:\RPA\Sample1\inputDir\';
	
//	1.엑셀 파일 열기
	var excel = Excel.open(dir + 'input.xlsx');
	var sheet = excel.sheet(1);
	
//	2.엑셀에서 데이터 가져오기 (rangeInfo 사용해보기)
//	var t = sheet["A1:E7"];

//	실무 rangeInfo 사용 방식
//	var range = data_sheet.rangeInfo();
//	var data_table = data_sheet["A2:" + range["column"] + range["row"]];

	var row = sheet.rangeInfo()["row"];
	var column = sheet.rangeInfo()["column"];
	var range = sheet[("A1:" +(column+row)];
	Console.writeLine('sheet["A1:E7"] : ');
	
	
//	for (var row : range) {
//		for (var item : row) {
//			Console.write(item + " ");
//		}
//			Console.writeLine();
//	}
//	3.새로운 엑셀 창에 값 넣기
	var excelNew = Excel.new();
	var sheetNew = excelNew.sheet(1);
	
//	4.새 창에 데이터 복사 (key : value로 작성해보기)
	sheetNew['A1'] = range;
	}
