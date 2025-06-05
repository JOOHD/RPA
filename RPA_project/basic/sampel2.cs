// 1. 엑셀 열기
var excel_Ex = Excel.new(); // 엑셀 파일 열기
var excel_sheet = excel_Ex.sheet(1); // 연 파일의 sheet 1page 지정

// 2. 엑셀에 있는 값 복사.
excel_sheet['A1'] = '예시파일입니다.'; // sheet(1)의 'A1' 구역에 '예시파일입니다.' 작성
excel_sheet['A2'] = [1,2,3];      
excel_sheet['A3:A4'] = [4,5];

var value1 = excel_sheet['A1']; // value 변수에 excel_sheet['A1'] 값 저장
var value2 = excel_sheet['A2'];
var value3 = excel_sheet['A3:A4'];

excel_Ex.saveAs("경로"); // 파일 저장

var excel_Ex2 = Excel.new(); // 두번째 엑셀 파일 열기
var excel2_sheet = excel_Ex2.sheet(1); // 두번째 엑셀 파일 sheet 1page 지정

excel2_sheet["A1"] = value1; // 두번째 엑셀 sheet value1 변수에 적용
excel2_sheet["A2"] = value2;
excel2_sheet["A3:A4"] = value3;

excel_Ex2.saveAs("경로");
