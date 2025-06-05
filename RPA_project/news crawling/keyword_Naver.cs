// 모든 App close
App.closeAll("iexplore");
App.closeAll("IEDriverServer");
App.closeAll("chrome");
App.closeAll("chromedriver");
App.closeAll("EXCEL");
Console.writeLine(Time.now().toString('[HH:mm:ss] ') + "프로세스 실행을 위해 모든 앱 종료");

// xPath
var xPathTitle ='//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/a';
var xPathPreview = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/div[2]'; 
var xPathDate = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/div/div[2]/span';
var xPathLink = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/a';
// var
var exRow = 0; // 변수 초기화
var newsTitle;
var newsPreview;
var newsDate;
var	newsLink;
var exceptionNum = 1;
var xPath_ilsu = null;
var xPath;
var xPath1;
var xPath2;
var	xPath3;
var xPath4;
var find_xPath1;
var find_xPath2;
var find_xPath3;
var find_xPath4;


//	0.엑셀 열기
var dir = 'C:\RPA\Sample\keyword\';

var excel = Excel.open(dir + 'keyword.xlsx');
var sheet = excel.sheet(1);
  
// 	1.네이버 열기
var naver = Browser.open("https://naver.com");

//  2.검색창에 키워드 검색
xPath_ilsu = '//*[@id="query"]';
naver.find(xPath_ilsu).write("ey한영"); // 검색 창 값 입력
xPath_ilsu = '/html/body/div[2]/div[2]/div[1]/div/div[3]/form/fieldset/button';
naver.find(xPath_ilsu).click(); // 검색 버튼 클릭
		
//naver.find('//*[@id="query"]').write("ey한영"); // 검색 창 값 입력
//naver.find('/html/body/div[2]/div[2]/div[1]/div/div[3]/form/fieldset/button').click(); // 검색 버튼 클릭
	
//	3.뉴스텝에 진입	
naver.find('/html/body/div[3]/div[1]/div/div[2]/div[1]/div/ul/li[2]/a').click(); // 뉴스 클릭

var i = 1;
var pressXpath = "";	
	for(var i : 2..10){
		//	4.뉴스 존재 유무
		if(naver.exists('//*[@id="sp_nws1"]/div[1]/div/a') == true){ // 뉴스 존재시
		}else{
			Console.writeLine("뉴스가 존재 하지 않습니다.");
			// 기사 없을시, 탈출 조건 걸어주기
		}
			var newsCnt = naver.findElements('//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li');
			for(var num : 1..newsCnt.size()) {	
				//title 저장
				xPath1 = xPathTitle.replace("%d", num.toString());   
				find_xPath1 = naver.find(xPath1);
				newsTitle = find_xPath1.read();	
							
				//date 저장		
				xPath2 = xPathDate.replace("%d", num.toString());
				find_xPath2 = naver.find(xPath2);
				newsDate = find_xPath2.read();
					
				//preview 저장
				xPath3 = xPathPreview.replace("%d", num.toString());   					
				find_xPath3 = naver.find(xPath3);
				newsPreview = find_xPath3.read();	
					
				//link 저장
				xPath4 = xPathLink.replace("%d", num.toString());   					
				find_xPath4 = naver.find(xPath4);
				newsLink = find_xPath4.readAttribute("href");
				
				//엑셀 값 저장
				exRow += 1; // 증감
				sheet['A' + exRow] = newsTitle; 
				sheet['B' + exRow] = newsDate;
				sheet['c' + exRow] = newsPreview;
				sheet["D" + exRow] = newsLink;				
				}
			
				//다음페이지 기능 추가
				pressXpath = '//*[@id="main_pack"]/div[2]/div/div/a['+ i +']';
				naver.wait(pressXpath).click();
				}
				sheet.columnWidth('A:B');
				
				//5.뉴스 데이터 엑셀 저장	
				excel.save();
				//excel.close();


//		naver.wait(xPath,1000).write(); //xPath에 값을 입력한다.  
// 		for(){}, if(){}
//		var range = sheet[("A1:" +(column+row)]; 엑셀 범위(가변성)
