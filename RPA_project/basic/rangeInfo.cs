
var rangeInfo = sheet.rangeInfo();

rangeInfo의 경우는 “row”나 “column”의 값을 넣어줘야합니다.

var rangeInfo = sheet.rangeInfo();
var row = rangeInfo[“row”]; // 현재 값이 존재하는 row의 마지막 숫자Ex) 엑셀의  “A1:E12”인 경우 12를 반환
var column = rangeInfo[“column”]; // 현재 값이 존재하는 Column의 마지막 알파벳 Ex) 엑셀의  “A1:E12”인 경우 E를 반환

// var excelRow = sheet.rangeInfo()["row"];
// var excelColumn = sheet.rangeInfo()["column"];

// var range defSheet.rangeInfo();
// var dataTable = defSheet["A2:"+range["column"]+ range["row"]];

Console.writeLine("max cell = " + rangeInfo['A1'] + rangeInfo['E6']);

위의 답변내용을 근거로 Console을 찍고 싶으면 하기와 같이 작성해야 합니다.
Console.writeLine(“max cell = A1:” + column + row);

//2. 시트에 있는 값 복사
var listData = sheet["A1:E5"];

rangeInfo를 사용해서 범위를 변경해서 적용해보세요
//excel.close();


var sheet2 = excel2.sheet(1);

var rangeInfo = sheet.rangeInfo();
Console.writeLine("max cell = " + rangeInfo['column'] + rangeInfo['row']);

sheet['M9'] = 'New Data';
rangeInfo = sheet.rangeInfo();
Console.writeLine("extended max cell = " + rangeInfo['column'] + rangeInfo['row']);
	
