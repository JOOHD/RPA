
//var list = [6, 7, 8];
//for(var idx : 0..list.size()) {
//  Console.writeLine("list =" + list[idx]);
//   } 

//for(var idx : list) {
//	Console.writeLine("list =" + idx);
//}

// list[0]

var list = ["a", "b", "c"];
for(var idx2 : 0..list.size()-1) {

	// idx2는 list의 인덱스 갯수 0,1,2
	Console.writeLine(idx2);
	
	// list의 인덱스 값을 console로 찍어내는 방법
	Console.writeLine("list =" + list[idx2]);
	idx2 += 1;
  } 




