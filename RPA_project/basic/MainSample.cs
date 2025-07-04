    
    //1. 엑셀 열기
    var excel = Excel.open('C:\Users\hddon\Desktop\RPA\Sample1\inputDir\input.xlsx');
    var sheet = excel.sheet(1);  // 이 예제에서 .sheet("Sheet1"); 로 해도 결과는 동일합니다.
    
	var list = [
	           ["1", "20", "남", "홍길동"],
	           ["", "22", "여", "김숙자"],
	           ["3", "21", "남", "김철수"],
	           ["", "22", "여", "김수미"],
	           ["5", "22", "여", "김미연"],
	           ["6", "26", "여", "김수정"]
	           ];
    
    Console.writeLine(list.toString());

    //3. 새로운 엑셀 열기
    var excelNew = Excel.new(); 
    var sheetNew = excelNew.sheet(1);
    
    	var newExRow = 0;
	
	for(var idx : 0..list.size()-1){
		
		Console.writeLine("idx 값 : " + idx);
		
		if(list[idx][0] == ""){
			Console.writeLine("공란이 맞습니다 : " + list[idx].toString());
			
		}else{
			
			newExRow += 1;
			
			Console.writeLine("공란이 아닙니다 : " + list[idx].toString());
	//    새로운 엑셀 시트에 값 붙여넣기
	sheetNew['A'+newExRow] = list[idx];
			
		}

	}

    //4. 새로운 엑셀 저장
    excelNew.saveAs('C:\Users\hddon\Desktop\RPA\Sample1\outputDir\output.xlsx'); 
    //excelNew.close();
