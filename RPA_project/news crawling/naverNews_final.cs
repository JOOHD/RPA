//	모든 App close
App.closeAll("iexplore");
App.closeAll("IEDriverServer");
App.closeAll("chrome");
App.closeAll("chromedriver");
App.closeAll("EXCEL");
Console.writeLine(Time.now().toString('[HH:mm:ss] ') + "프로세스 실행을 위해 모든 앱 종료");

// xPath
var xPathTitle = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/a';
var xPathSource = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/div[1]/div[2]/a';
var xPathDate = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/div[1]/div[2]/span';
var xPathLink = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/a';
var xPathPreview = '//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li[%d]/div[1]/div/div[2]/div/a';
var	newsXpathTab = '//*[@id="lnb"]/div[1]/div/ul/li[%d]/a';
var press_xPath = '//*[@id="main_pack"]/div[2]/div/a[2]';

// 변수
var xPath;
var xPath1;
var xPath2;
var xPath3;
var xPath4;
var xPath5;
var newsTitle;
var newsSource;
var newsDate;
var newsLink;
var newsPreview;
var newSheet;

var i = 0;
var exRow;

//input 엑셀 파일 오픈
var dir = 'C:\RPA\naverNews_inquire\input\';

//output 엑셀 파일 오픈
var dir2 = 'C:\RPA\naverNews_inquire\output\';
var excel = Excel.open(dir + '키워드리스트_20220119.xlsx');

//sheet 펼치기
var sheet = excel.sheet(1);

//sheet에 rangeInfo 적용
var range = sheet.rangeInfo(true);

//검색 키워드 리스트, 다차원 배열 이용
var data_table = sheet["B2:" + range["column"] + range["row"]];

//TODO 2022.01.24(월)- 엑셀 전체 테이블 rangeInfo 적용
//data_table[i][0] = keyword
//data_table[i][1] = progress_check
//data_table[i][2] = period

//새로운 output엑셀 파일 오픈
var excelNew = Excel.new(dir2 + '뉴스데이터조회결과_20220124.xlsx');
var sheetNew = excelNew.sheet(1);

//네이버 열기
var naver = Browser.open("https://naver.com");

//header 설정
var header = [
			"NO", "제목", "출처", "등록일", "뉴스 링크", "미리보기"	
			 ];
			 
for(var i : 0..data_table.size()-1){
	
	//검색창에 키워드 검색, 진행 결과 "YES"
	if(data_table[i][1].trim() == "YES"){
		//키워드가 달라질 때마다 xpath 위치 변화	
		if(i == 0){
			naver.find('//*[@id="query"]').write(data_table[i][0].trim());
			naver.find('//*[@id="search_btn"]').click();
		}else{
			naver.find('//*[@id="nx_query"]').write(data_table[i][0].trim()); //검색 창 값 입력
			naver.find('//*[@id="nx_search_form"]/fieldset/button').click();	
		}	
	}else{//진행 결과 "NO"
		Console.writeLine("진행 결과 NO 다음 키워드 검색");
		continue;	
	}
	
	//엑셀 sheet 추가하기
	try{
		excelNew.addSheet(data_table[i+1][0].trim());
	}catch(e){
		Console.writeLine("더 이상 추가할 sheet가 존재하지 않습니다");
		//break 를 걸어놓으면 sheet가 추가 되지 않고, 종료. continue 로 넘겨줘야한다.
		continue;
	}
	//sheet명 변경
	excelNew.renameSheet(i+1, data_table[i][0].trim()); 
	sheetNew = excelNew.sheet(data_table[i][0].trim());

	//sheet(1)에 중복으로 계속 쌓이는 것을 방지
	exRow = sheetNew.rangeInfo(true)["row"];

	//엑셀 header 입력
	sheetNew['A1'] = header;
	
	
	//TODO 2022.01.24(월)-  findElements 사용, 뉴스탭 xpath 변동 해결
	//뉴스 탭 클릭
	//키워드가 달라질 때마다 xpath 위치 변화
	var newsTabList = naver.findElements('//*[@id="lnb"]/div[1]/div/ul/li');
	for(var num : 1..newsTabList.size()){
	    var xPathTab = newsXpathTab.replace("%d", num.toString());
        if(naver.find(xPathTab).read().contains("뉴스") == true){
            naver.find(xPathTab).click();	
        }else{
            continue;
        }             
	}
	
	//정렬 : 최신순 탭 클릭
	if(naver.exists('//*[@id="snb"]/div[1]/div/div[1]/a[2]') == true) //탭 버튼이 활성화 되어있을경우
	{
		//탭 클릭이 안될 경우
		try
		{
			System.sleep(2000);
			naver.find('//*[@id="snb"]/div[1]/div/div[1]/a[2]').jclick(); //옵션 버튼 클릭
		} catch(e) {
			Console.notice(e['line']);
			Console.notice(e['stack']);
			Console.notice(e['message']);
		}
	}
	
	
	
	//옵션 탭 클릭
	if(naver.exists('//*[@id="snb"]/div[1]/div/div[2]/a') == true) //옵션버튼이 활성화 되어있을경우
	{
		//탭 클릭이 안될 경우
		try
		{
			System.sleep(2000);
			naver.find('//*[@id="snb"]/div[1]/div/div[2]/a').jclick(); //옵션 버튼 클릭
		} catch(e) {
			Console.notice(e['line']);
			Console.notice(e['stack']);
			Console.notice(e['message']);
		}
	}
	
	//TODO 2022.01.24(월)-  기간 설정 후, 검색 결과 없으면 sheet에 표시	
	if(data_table[i][2].trim() == "1일"){ 
		naver.find('//*[@id="snb"]/div[2]/ul/li[2]/div/div[1]/a[3]').click(); //기간 1일		
		//뉴스기사가 없는 경우 continue 다음 키워드 넘어가기
		if(!naver.exists('//*[@id="sp_nws1"]')){
			sheetNew['A' + exRow] = "1";
			sheetNew['B' + exRow] = "조회 된 뉴스가 존재 하지 않습니다.";
			//break 를 걸어놓으면 바로 종료. continue 로 넘겨줘야한다.
			continue;
		} 
	}else{ 
		naver.find('//*[@id="snb"]/div[2]/ul/li[2]/div/div[1]/a[8]').click(); //기간 1년
		//뉴스기사가 없는 경우 continue 다음 키워드 넘어가기
		if(!naver.exists('//*[@id="sp_nws1"]')){
			sheetNew['A' + exRow] = "1";
			sheetNew['B' + exRow] = "조회 된 뉴스가 존재 하지 않습니다.";
			continue;
		}
	}	
	while(true){
		//findElements 사용 newsCnt 변수에 담기
		var newsCnt = naver.findElements('//*[@id="main_pack"]/section[contains(@class, "sc_new sp_nnews _prs_nws")]/div/div[contains(@class, "group_news")]/ul/li');
		for(var num : 1..newsCnt.size()){
			
			//xPathTitle을 full xpath로 기입해라.
			xPath = xPathTitle.replace("%d", num.toString());
			if(naver.exists(xPath)){ 
				//제목
				//xPath1 = xPathTitle.replace("%d", num.toString());
				newsTitle = naver.find('//*[@id="sp_nws'+num+'"]/div/div/a').read();
				
				// 웹페이지에서 'RPA' 기사 가져오는 부분
				////*[@id="sp_nws['+num+']"]
//				var naverTitle = naver.find('//*[@id="main_pack"]/div[1]/ul/li['+i+']/dl/dt/a').read(); // 타이틀
//				var naverLink = naver.find('//*[@id="main_pack"]/div[1]/ul/li['+i+']/dl/dt/a').readAttribute('href'); // 링크
//				var naverSource = naver.find('//*[@id="main_pack"]/div[1]/ul/li['+i+']/dl/dd[1]/span[1]').read(); // 출처
//				var naverContent = naver.find('//*[@id="main_pack"]/div[1]/ul/li['+i+']/dl/dd[2]').read(); // 컨텐츠
				
				//언론사
				xPath2 = xPathSource.replace("%d", num.toString());
				newsSource = naver.find(xPath2).read();
				
				//날짜		
				xPath3 = xPathDate.replace("%d", num.toString());
				newsDate = naver.find(xPath3).read();
				
				//link 저장
				xPath4 = xPathLink.replace("%d", num.toString());   					
				newsLink = naver.find(xPath4).readAttribute("href");
					
				//미리보기
				xPath5 = xPathPreview.replace("%d", num.toString());   					
				newsPreview = naver.find(xPath5).read();
					
				//크롤링 기사 엑셀 넣기
				exRow += 1; //행, 증감
				sheetNew['A'+exRow] = exRow-1; //A2부터 번호 1번, 원래는 2번(A1은 header가 있기 떄문에)
				sheetNew['B'+exRow] = newsTitle;
				sheetNew['C'+exRow] = newsSource;
				sheetNew['D'+exRow] = newsDate;
				sheetNew['E'+exRow] = "link" : newsLink; 
				sheetNew['F'+exRow] = newsPreview;
			}//if
		}//for
		//sheet 자동 정렬
		sheetNew.columnWidth('A:F');
		
		//페이지 next 버튼 클릭
		if(naver.find('//*[@id="main_pack"]/div[contains(@class,"api_sc_page_wrap")]/div/a[2]').readAttribute('aria-disabled') == false){ 
			naver.find('//*[@id="main_pack"]/div[contains(@class,"api_sc_page_wrap")]/div/a[2]').jclick();
		}else{
			break;
		}
			
		//다음 페이지 버튼
//		naver.wait(press_xPath).click();
	}//while
	
	//검색창 클리어
	naver.find('//*[@id="nx_query"]').clear();
}//for
//뉴스 데이터 엑셀 저장	
excel.save(dir2 + '뉴스데이터조회결과_20220124.xlsx');
//excel.close();

