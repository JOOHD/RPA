//	모든 App close
App.closeAll("iexplore");
App.closeAll("IEDriverServer");
App.closeAll("chrome");
App.closeAll("chromedriver");
App.closeAll("EXCEL");
Console.writeLine(Time.now().toString('[HH:mm:ss] ') + "프로세스 실행을 위해 모든 앱 종료");

//section xpath
//var noImgTitle ='//*[@id="newsColl"]/div[contains(@class, "wrap_cont")]/ul/li[%d]/div/a';
//var noImgPreview ='//*[@id="newsColl"]/div[contains(@class, "wrap_cont")]/ul/li[%d]/div/p';
//var noImgDate ='//*[@id="newsColl"]/div[contains(@class, "wrap_cont")]/ul/li[%d]/div/span[1]/span[2]';
//var noImgLink ='//*[@id="newsColl"]/div[contains(@class, "wrap_cont")]/ul/li[%d]/div/a';

//이미지 없는xpath
var noImgTitle ='//*[@id="newsColl"]/div[1]/ul/li[%d]/div/a';
var noImgDate ='//*[@id="newsColl"]/div[1]/ul/li[3]/div/span[1]/span[2]';
var noImgPreview ='//*[@id="newsColl"]/div[1]/ul/li[%d]/div/span[1]/span[2]';
var noImgLink ='//*[@id="newsColl"]/div[1]/ul/li[%d]/div/a';

//이미지 있는xpath
var xPathTitle ='//*[@id="newsColl"]/div[1]/ul/li[%d]/div[2]/a'; //이승우
				//*[@id="newsColl"]/div[1]/ul/li[1]/div[2]/a 	백승호
var xPathPreview = '//*[@id="newsColl"]/div[1]/ul/li[%d]/div[2]/p'; 
var xPathDate = '//*[@id="newsColl"]/div[1]/ul/li[%d]/div[2]/span[1]/span[2]';
var xPathLink = '//*[@id="newsColl"]/div[1]/ul/li[%d]/div[2]/a';

//다음 페이지 버튼 xpath
var	pressXpath = '//*[@id="newsColl"]/div[2]/span/span[3]/a';

//*[@id="newsColl"]/div[1]/ul/li[x]/div[2]/a -> 이미지가 있는 경우 제목의 xpath
//*[@id="newsColl"]/div[1]/ul/li[x]/div/a-> 이미지가 없는 경우 제목의 xpath
	
//img 관련 변수
var xPath;
var xPathT;
var xPathD;
var	xPathP;
var xPathL;
var find_xPathT;
var find_xPathD;
var find_xPathP;
var find_xPathL;
var newsTitle;
var newsPreview;
var newsDate;
var	newsLink;

//noImg 관련 변수
var xPathNo;
var xPathNoT;
var xPathNoD;
var xPathNoP;
var xPathNoL;
var	find_xPathNoT;
var	find_xPathNoD;
var	find_xPathNoP;
var	find_xPathNoL;
var newsTitleNo;
var newsDateNo;
var newsPreviewNo;
var newsLinkNo;

//header 설정
var sample = [
			["제목", "날짜", "내용"]	
			 ];
//키워드 리스트 선언			 
var searchList = ["이승우", "백승호", "박지성"];

//엑셀 열기
var dir = 'C:\RPA\과제 자료\Sample\keyword\';
var excel = Excel.open(dir + 'keyword.xlsx');

//다음 열기
var daum = Browser.open("https://daum.net");
	
//초기화 및 변수선언 해주어야 다음 sheet로 넘어갈때 처음부터 데이터 쌓임.
var exRow; 

//sheet 초기화 및 변수 선언
var sheet;

//i 변수 선언 및 초기값 = 0 선언
var i = 0;

//키워드 사이즈 만큼 반복문(배열 인덱스는 0부터), size는 개수라, 1,2,3그래서 -1넣어주는 이유
for(var i : 0..searchList.size()-1){
	
	//엑셀 sheet 이름을 바꾸기 위한 코드이고
	//엑셀 sheet 이름 바꾸기, i는 초기값 0, sheet는 1부터 그래서 i+1
	//rename하고 (sheet숫자, sheet명)
	excel.renameSheet(i+1, searchList[i]);
	
	//sheet를 정의, 첫번째 sheet를 펼치겠다.(엑셀에 만들겠다.(keyword로 작명된))
	//Excel.open으로 엑셀은 열어주었지만, sheet는 열어주지 않은것. sheet열어주는 문장
	//sheet괄호 안에 값을 담기 가능
	sheet = excel.sheet(searchList[i]);  
	
	//rangeInfo['row'] 선언
	//rangeInfo(true)는 공란을 제외한 데이터가 들어간 셀 가져오기
	exRow = sheet.rangeInfo(true)["row"];
	
	//엑셀 header 입력
	sheet['A1'] = sample;
	
	//검색창에 키워드 검색.write(키워드 리스트)
	daum.find('//*[@id="q"]').write(searchList[i]);
	daum.find('//*[@id="daumSearch"]/fieldset/div/div/button[2]').click();

	//뉴스텝에 진입
	if(i == 0){//키워드가 0일때와 다음 키워드가 올때 뉴스탭의 xpath값이 달라서 설정
		daum.find('//*[@id="daumGnb"]/div/ul/li[2]/a/span/span').click(); //뉴스 클릭
	}else{
		daum.find('//*[@id="daumGnb"]/div/ul/li[4]/a/span/span').click(); //뉴스 클릭 
	}
	//기간 설정
	daum.find('//*[@id="newsColl"]/div[1]/div[2]/div/div[1]/a/span[1]').click(); // 기간 클릭
	daum.find('//*[@id="newsColl"]/div[1]/div[2]/div[1]/div[1]/div/ul/li[4]/a').click(); // 최근 7일 클릭
		
	//.관련 인물선택
	if(daum.exists('//*[@id="selectSamename"]')){ //검색어 마다 관련 인물선택이 있을수도, 없을수도 있어서 분기처리.
		daum.find('//*[@id="selectSamename"]').click(); // 항목창 클릭
		daum.find('//*[@id="selectSamename"]/option[3]').click(); // 선택 항목 클릭
	}

	//여기 while문은 다음페이지 버튼을 반복시켜주는 역할 
	while(true) {
		//findElements = xpath의 일치하는 elements를 반환시켜주는(list타입)
		//newsCnt는 1p의 뉴스크롤링 할 구역을 xpath로 잡은 변수, 1페이지 구역에있는 기사들을 담아놓은것
		var newsCnt = daum.findElements('//*[@id="newsColl"]/div[1]/ul/li');	
		
		//newsCnt에 있는 기사들을 1번째 기사부터 size개수 만큼 반복한다. num = 기사1, 기사2, 기사3..기사 하나마다 num에 담는것
		//newsCnt의 1페이지에는 10개의 기사가 있으니, 10개의 num이 있다고 보면될듯.
		for(var num : 1..newsCnt.size()) {
			
			// 더 이상 뉴스기사가 없는 경우 break로 종료
			if(!daum.exists('//*[@id="newsColl"]/div[1]/ul/li')){
				Console.writeLine(Time.now().toString("[HH:mm:ss] ") + "다음 뉴스가 없습니다.");
				break;
			}
					
			//이미지 없는 기사 title
			//xPath = *[@id="newsColl"]/div[1]/ul/li[%d]/div/a			
			//이미지 있는 기사 title
			//xPath = *[@id="newsColl"]/div[1]/ul/li[%d]/div[2]/a
			
			//이미지 xPath 구하기
			//xPath는 페이지에 있는 기사제목들을 위치값만 바뀌는 부분을 %d로 치환, 기사제목 개수 = num개수
			//즉 xPath는 %d로 위치값이 다른 xpath들을 통일, num으로 순번을 정해준 xpathTitle을 xpath변수에 담는다.
			xPath = xPathTitle.replace("%d", num.toString());
			
			//위에 있는 xPath가 없을 경우
			//exists의 기준을 xpathTitle로 잡는 이유는 Img, noImg 구분은 div[]의 대괄호 존재 유무 차이이기 때문이다.
			if(!daum.exists(xPath)){ //div인 경우
				//이미지 없음							    
			    
			    // 제목
			    // 변수  =  타이틀xpath     변하는 수  넘버링
			    xPathNoT = noImgTitle.replace("%d", num.toString());
			    //타이틀 위치값 찾아라 xPathNoT에서
				find_xPathNoT = daum.find(xPathNoT);
				//새 변수 담아라 = 찾은 위치값.가져와라
				//newsTitleNo = daum.find(xPathNoT).read();
				newsTitleNo = find_xPathNoT.read();
				
				// 날짜
				xPathNoD = noImgDate.replace("%d", num.toString());
				find_xPathNoD = daum.find(xPathNoD);
				newsDateNo = find_xPathNoD.read();
				
				// 내용
				xPathNoP = noImgPreview.replace("%d", num.toString());
				find_xPathNoP = daum.find(xPathNoP);
				newsPreviewNo = find_xPathNoP.read();
				
				// 링크
				xPathNoL = noImgLink.replace("%d", num.toString());
				find_xPathNoL = daum.find(xPathNoL);
				newsLinkNo = find_xPathNoL.readAttribute("href");
				
				//이미지 없는 기사 엑셀 값 저장
				exRow += 1; // 증감
				sheet['A' + exRow] = "value" : newsTitleNo, "link" : newsLinkNo;
				sheet['B' + exRow] = newsDateNo;
				sheet['C' + exRow] = newsPreviewNo;

			}else{ //div[]인 경우
				// 이미지 있음			    
				
				//title 저장
				xPathT = xPathTitle.replace("%d", num.toString());   
				find_xPathT = daum.find(xPathT);
				newsTitle = find_xPathT.read();
							
				//date 저장		
				xPathD = xPathDate.replace("%d", num.toString());
				find_xPathD = daum.find(xPathD);
				newsDate = find_xPathD.read();
				
				//preview 저장
				xPathP = xPathPreview.replace("%d", num.toString());   					
				find_xPathP = daum.find(xPathP);
				newsPreview = find_xPathP.read();	
				
				//link 저장
				xPathL = xPathLink.replace("%d", num.toString());   					
				find_xPathL = daum.find(xPathL);
				newsLink = find_xPathL.readAttribute("href");
				
				//이미지 있는 기사 엑셀 값 저장
				exRow += 1; // 행, 한 줄씩 증감 
				sheet['A' + exRow] = "value" : newsTitle, "link" : newsLink; //제목에 링크 씌우기
				sheet['B' + exRow] = newsDate;
				sheet['C' + exRow] = newsPreview;			
			}	
		}//for
		//칼럼 자동 정렬
		sheet.columnWidth('A:C');

		//다음페이지 버튼 없으면 종료
		if(!daum.exists('//*[@id="newsColl"]/div[2]/span/span[3]/a')) {
			Console.writeLine(Time.now().toString("[HH:mm:ss]") + "다음 페이지 버튼이 없으므로 종료");
			break;
		}//if
		daum.wait(pressXpath).click();		
	}//while
	//엑셀 sheet 추가하기
	//i=0 부터니까, sheet는 1부터 페이지니까 +1을 해주는거.	
	//sheet붙은 기능을 구현할때 sheet, addSheet +1해주는거 한번 생각해보기.
	excel.addSheet(searchList[i+1]);
	
	daum.find('//*[@id="q"]').clear(); //검색창 클리어
}//for
//5.뉴스 데이터 엑셀 저장	
//excel.save();
//excel.close();

