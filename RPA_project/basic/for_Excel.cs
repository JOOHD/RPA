// 1.엑셀 열기
var excel = Excel.open('C:\Users\hddon\Desktop\RPA\Sample1\inputDir\input.xlsx');
var sheet = excel.sheet(1); 

pause;
var t = 1;

while(true){
	if(t == 5) {
		break;
	}		
	sheet = excel.sheet(t); 
	excel.addSheet();
	t += 1;	
}
