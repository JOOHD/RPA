


var list = [
           ["1", "20", "남", "홍길동"],
           ["", "22", "여", "김숙자"],
           ["3", "21", "남", "김철수"]
           ];
           
//pause;
	Console.writeLine(list.toString());
	
	//3. 새로운 엑셀 열기
    var excelNew = Excel.new(); 
    var sheetNew = excelNew.sheet(1);
	
	var newExRow = 0;
	
	for(var idx : 0..list.size()-1){
		
		Console.writeLine("idx 값 : " + idx);
		
		if(list[idx][0] == ""){
			Console.writeLine("공란이 맞습니다 : " + list[idx].toString()); // list는 console에서는 toString 찍어주기
			
		}else{
			
			newExRow += 1;
			
			Console.writeLine("공란이 아닙니다 : " + list[idx].toString());
			sheetNew['A'+newExRow] = list[idx];
			
		}
		pause;
	}
	
	
	
