 			
 var t = 1 ;
 
 while(true) {
	if(t == 5) {
		 break;
	}else if(t == 1){
		Console.writeLine("a");
	}else if(t == 2){
		Console.writeLine("b");
	}else if(t == 3){
		Console.writeLine("c");
	}else if(t == 4){
		Console.writeLine("d");	
	}else if(t == 5){
		Console.writeLine("e");	
	}else{		 
	 	Console.writeLine("조건에 해당되지 않는 숫자입니다.");
	}
	System.sleep(3000);
	t += 1;
 }
