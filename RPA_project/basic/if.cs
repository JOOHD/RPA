// 1~10까지 6에서 탈출
// idx 변수는 배열 인덱스
// idx2 변수는 list안의 인데스 값
var list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
var idx2 = 0; 
for(var idx : 0..list.size()-1) {
	if(idx2 == 6){
	break;
}
	Console.writeLine(idx2);
	Console.writeLine("list =" + list[idx2]);
	idx2 += 1;
} 

// 1~10까지 6을 빼고 
var t = 0; 
for(var t : 0..list.size()-1) {
	if(list[t] == 6){
     continue;
	}
	Console.writeLine("list =" + list[t]);
	t += 1;
} 
