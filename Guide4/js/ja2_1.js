//------------------------ 시작 -------------------------------//
//기본변수
var i;//펑션 숫자 변수
var k;
var temp;
var onOff;

//var linkArea ='background-color:cyan;opacity:0.5;filter:alpha(opacity=50);';

//공통함수
function Left(Str, Num){
	if (Num <=0){return "";}
	else if (Num > String(Str).length){return Str;}
	else {return String(Str).substring(0, Num);}		
}
function Right(Str, Num){
	if (Num <=0){return "";}
	else if (Num > String(Str).length){return Str;}
	else {var iLen=String(Str).length;return String(Str).substring(iLen, iLen-Num);}
}

/*HTML 복사*/
function jsCopy(obj) {
 
 var copyObj = obj;
 var copyStr = document.getElementById(copyObj).innerHTML;

 
 window.clipboardData.setData("text", copyStr);
 alert("Html 소스가 복사되었습니다.");
}

//------------------------- 공통 영역 --------------------------------//


//------------------------- 기타 영역 --------------------------------//
// 코딩변경
function reCode(n){
	temp = String(n);
	temp = temp.replace(/\</gi, "&lt;"); //치환	
	temp = temp.replace(/\>/gi, "&gt;"); //치환	
	temp = temp.replace(/\.jpg\"&gt;/gi, '.jpg\"/&gt;'); //치환	
}

//------------------------- 엑셀 영역 --------------------------------//
//엑셀 불러오기
function excel(Location, i){ //엑셀이 입력될 타겟, 기획전타입번호
	var sCode = document.frm.series.value; //document.getElementById('dev_tbl').getElementsByTagName('input')[2*(k-1)+0].value;//시리즈코드 인풋값 반영
	var nCode = document.frm.numb.value; //document.getElementById('dev_tbl').getElementsByTagName('input')[2*(k-1)+1].value;//넷하드 상품번호 인풋값 반영
	var eCode;
	var sum ='';
	
	//---------------------------------상하단 공통영역(개발중)------------------------------------/
	//상품평 이벤트
	var eve=new Array();
	eve[101]=''+
			'<!-- 상품평 이벤트 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/01_ismine/01_common/01_event/01.jpg" alt="상품평 이벤트" />';
	eve[102]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/02_system/01_common/01_event/01.jpg" alt="event" />'; //시스템주방
	eve[103]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/03_ss/01_common/01_event/01.jpg" alt="event" />'; //리바트 홈스타일
	eve[104]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/04_hs/01_common/01_event/01.jpg" alt="event" />'; //이즈마인 홈스타일
	eve[105]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/05_neoce/01_common/01_event/01.jpg" alt="event" />'; //네오스 온라인
	eve[106]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/06_hmondo/01_common/01_event/01.jpg" alt="event" />'; //H몬도
	eve[107]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/07_chairs/01_common/01_event/01.jpg" alt="event" />'; //리바트체어스
	eve[108]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/08_ICT/01_common/01_event/01.jpg" alt="event" />'; //ICT
	eve[109]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/09_mallin/01_common/01_event/01.jpg" alt="event" />'; //몰인몰
	eve[199]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/99_etc/01_common/01_event/01.jpg" alt="event" />'; //etc

	eve[201]=''+
			'<!-- 상품평 이벤트 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/01_livart/01_common/01_event/01.jpg" alt="event" />'; //리바트일때
	eve[202]=''+
			'<!-- 상품평 이벤트 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/02_ritchen/01_common/01_event/01.jpg" alt="event" />'; //리첸

	eve[203]=''+
			'<!-- 상품평 이벤트 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/03_neoce/01_common/01_event/01.jpg" alt="event" />'; //네오스
	eve[204]=''+
			'<!-- 상품평 이벤트 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/04_Lkitchen/01_common/01_event/01.jpg" alt="event" />'; //리바트키친
	eve[205]=''+
			'<!-- 상품평 이벤트 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/05_kids/01_common/01_event/01.jpg" alt="event" />'; //리바트키즈
	eve[206]=''+
			'<!-- event 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/06_hmondo/01_common/01_event/01.jpg" alt="event" />'; //H몬도


	//이즈마인 서약
	var vow=new Array();
	vow[101]='<!-- 이즈마인 서약 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/01_ismine/01_common/02_vow/01.jpg" alt="이즈마인서약" />';
	vow[102]='<!-- vow 영역(공통) -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/02_system/01_common/02_vow/01.jpg" alt="서약" />'; //시스템주방일때
	vow[103]='<!-- 리바트 홈스타일 상단 영역(공통) -->'+ //vow부분에 홈스타일 상단 내용 들어감
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/03_ss/01_common/02_vow/01.jpg" alt="서약" />';
	vow[104]='<!-- 이즈마인 홈스타일 상단 영역(공통) -->'+ //vow부분에 홈스타일 상단 내용 들어감
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/04_hs/01_common/02_vow/01.jpg" alt="이즈마인서약" />';
	vow[105]='<!-- 네오스 상단 영역(공통) -->'+ //vow
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/05_neoce/01_common/02_vow/01.jpg" alt="서약" />'; //네오스 온라인
	vow[106]='<!-- H몬도 상단 영역(공통) -->'+ //vow
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/06_hmondo/01_common/02_vow/01.jpg" alt="서약" />'; //H몬도
	vow[107]='<!-- 리바트체어스 상단 영역(공통) -->'+ //vow
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/07_chairs/01_common/02_vow/01.jpg" alt="서약" />'; //리바트 체어스
	vow[108]='<!-- 리바트 홈스타일 상단 영역(공통) -->'+ //vow부분에 홈스타일 상단 내용 들어감
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/08_ICT/01_common/02_vow/01.jpg" alt="서약" />';
	vow[109]='<!-- 몰인몰 상단 영역(공통) -->'+ //vow부분에 홈스타일 상단 내용 들어감
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/09_mallin/01_common/02_vow/01.jpg" alt="서약" />'; //몰인몰
	vow[199]='<!-- 상단 영역(공통) -->'+ //
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/99_etc/01_common/02_vow/01.jpg" alt="서약" />'; //etc

	vow[201]='<!-- 리바트 환경경영 상단 영역(공통) -->'+ //vow부분에 리바트 환경경영 들어감
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/01_livart/01_common/02_vow/01.jpg" alt="리바트 환경경영" />'; //리바트 환경경영
	vow[202]='<!-- vow영역(공통) -->'+ //vow부분
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/02_ritchen/01_common/02_vow/01.jpg" alt="vow" />'; //vow
	vow[203]='<!-- vow영역(공통) -->'+ //vow부분
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/03_neoce/01_common/02_vow/01.jpg" alt="vow" />'; //vow
	vow[204]='<!-- vow영역(공통) -->'+ //vow부분
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/04_Lkitchen/01_common/02_vow/01.jpg" alt="vow" />'; //vow
	vow[205]='<!-- vow영역(공통) -->'+ //vow부분
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/05_kids/01_common/02_vow/01.jpg" alt="vow" />'; //vow
	vow[206]='<!-- H몬도 상단 영역(공통) -->'+ //vow
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/06_hmondo/01_common/02_vow/01.jpg" alt="" />'; //H몬도

	//배송안내 이즈마인
	var ship=new Array();

	ship[101] ='<!-- 이즈마인 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/01_ismine/01_common/99_ship/01.jpg" alt="리바트 온라인 배송 정보" name="ship" id="ship" />';

	//배송안내 시스템주방
	ship[102] ='<!-- 시스템주방 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/02_system/01_common/99_ship/01.jpg" alt="리바트 시스템주방 온라인 배송 정보" name="ship" id="ship" />';

	//배송안내 홈스타일
	ship[103] ='<!-- 리바트 홈스타일 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/03_ss/01_common/99_ship/01.jpg" alt="리바트 홈스타일 배송 정보" name="ship" id="ship" />';
	ship[104] ='<!-- 이즈마인 홈스타일 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/04_hs/01_common/99_ship/01.jpg" alt="이즈마인 홈스타일 배송 정보" name="ship" id="ship" />';
	ship[105] ='<!-- 온라인 네오스 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/05_neoce/01_common/99_ship/01.jpg" alt="네오스 온라인 배송 정보" name="ship" id="ship" />';
	ship[106] ='<!-- H몬도 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/06_hmondo/01_common/99_ship/01.jpg" alt="H몬도 배송 정보" name="ship" id="ship" />';
	ship[107] ='<!-- 리바트체어스 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/07_chairs/01_common/99_ship/01.jpg" alt="리바트체어스 배송 정보" name="ship" id="ship" />';
	ship[108] ='<!-- 리바트홈 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/08_ICT/01_common/99_ship/01.jpg" alt="리바트홈 배송 정보" name="ship" id="ship" />';
	ship[109] ='<!-- 몰인몰 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/09_mallin/01_common/99_ship/'+sCode+'.jpg" alt="배송 정보" name="ship" id="ship" />'+
			//'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/09_mallin/01_common/99_ship/01.jpg" alt="몰인몰 공통 배송 정보" />'+
			'';
	ship[199] ='<!-- 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/99_etc/01_common/99_ship/01.jpg" alt="배송 정보" name="ship" id="ship" />';

	
	//배송안내 리바트
	ship[201] ='<!-- 리바트 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/01_livart/01_common/99_ship/01.jpg" alt="리바트 온라인 배송 정보" name="ship" id="ship" />';
	ship[202] ='<!-- ship 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/02_ritchen/01_common/99_ship/01.jpg" alt="배송 정보" name="ship" id="ship" />';
	ship[203] ='<!-- ship 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/03_neoce/01_common/99_ship/01.jpg" alt="배송 정보" name="ship" id="ship" />';
	ship[204] ='<!-- ship 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/04_Lkitchen/01_common/99_ship/01.jpg" alt="배송 정보" name="ship" id="ship" />';
	ship[205] ='<!-- ship 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/05_kids/01_common/99_ship/01.jpg" alt="배송 정보" name="ship" id="ship" />';
	ship[206] ='<!-- H몬도 : 양방향 메시징 서비스 영역(공통)  -->'+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/02_offline/06_hmondo/01_common/99_ship/01.jpg" alt="H몬도 배송 정보" name="ship" id="ship" />';

	

	//i값 - 온라인: 101, 오프라인: 201
	switch (i){
		//온라인 브랜드
		case '101': onOff='01_online';outPut='01_ismine';eve=eve[101];vow=vow[101];ship=ship[101];break; //이즈마인
		case '102': onOff='01_online';outPut='02_system';eve=eve[102];vow=vow[102];ship=ship[102];break; //시스템주방
		case '103': onOff='01_online';outPut='03_ss';eve=eve[103];vow=vow[103];ship=ship[103];break; //홈스타일
		case '104': onOff='01_online';outPut='04_hs';eve=eve[104];vow=vow[104];ship=ship[104];break; //하움
		case '105': onOff='01_online';outPut='05_neoce';eve=eve[105];vow=vow[105];ship=ship[105];break; //네오스 온라인
		case '106': onOff='01_online';outPut='06_hmondo';eve=eve[106];vow=vow[106];ship=ship[106];break; //H몬도
		case '107': onOff='01_online';outPut='07_chairs';eve=eve[107];vow=vow[107];ship=ship[107];break; //리바트체어스
		case '108': onOff='01_online';outPut='08_ICT';eve=eve[108];vow=vow[108];ship=ship[108];break; //ICT
		case '109': onOff='01_online';outPut='09_mallin';eve=eve[109];vow=vow[109];ship=ship[109];break; //몰인몰
		case '199': onOff='01_online';outPut='99_etc';eve=eve[199];vow=vow[199];ship=ship[199];break; //etc

		//오프라인 브랜드
		case '201': onOff='02_offline';outPut='01_livart';eve=eve[201];vow=vow[201];ship=ship[201];break; //리바트
		case '202': onOff='02_offline';outPut='02_ritchen';eve=eve[202];vow=vow[202];ship=ship[202];break; //리첸
		case '203': onOff='02_offline';outPut='03_neoce';eve=eve[203];vow=vow[203];ship=ship[203];break; //네오스
		case '204': onOff='02_offline';outPut='04_Lkitchen';eve=eve[204];vow=vow[204];ship=ship[204];break; //키친
		case '205': onOff='02_offline';outPut='05_kids';eve=eve[205];vow=vow[205];ship=ship[205];break; //키즈
		case '206': onOff='02_offline';outPut='06_hmondo';eve=eve[206];vow=vow[206];ship=ship[206];break; //H몬도
		
		default : j=99; alert('다시 설정 해주세요.');break;
	};

	//---------------------------------------------------------------------------------/

	//엑셀용 변수 선언
	var Excel,oRange;

	// 엑셀 객체생성
	Excel = new ActiveXObject("Excel.Application");
	//Excel.Application.Workbooks.Open('\\\\192.167.22.242\\Lgudiesystem\\pd\\'+onOff+'\\'+outPut+'\\'+sCode+'.xls');
	Excel.Application.Workbooks.Open('\\\\172.17.14.220\\Lguidesystem\\pd\\'+onOff+'\\'+outPut+'\\'+sCode+'.xls');
	//Excel.Application.Workbooks.Open('C:/Guide/pd/'+onOff+'/'+outPut+'/'+sCode+'.xls');
	//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/pd/'+onOff+'/'+outPut+'/'+sCode+'.xls'); 

	//---------------------------------공통영역 alt값 (삭제예정)------------------------------------/
	
	//공통영역 번호 노출
	/*Excel.Application.Worksheets("00").Activate;

	//00내 작업영역을 객체에 저장
	oRange = Excel.Application.ActiveSheet.UsedRange;

	//oRange에는 현재 내용이 있는 부분만 선택됨.
	//oRange부분을 선택표시
	oRange.Select();

	//var tem = oRange.Cells(1,2);
	//alert(tem);

	var alt=new Array();//alt값과 코드를 위한 배열선언

	for(var e=7;e<=oRange.Rows.count;e++){ 
		alt[e] = new Array(); //alt[e]값 배열선언
		
		for(var j=1;j<=oRange.Columns.count;j++){ 
			alt[e][j] = oRange.Cells(e,j);
			//alert(alt[e][j]);
		}
	}*/
 
	//---------------------------------고유번호영역------------------------------------/

	//엑셀표시
	Excel.Application.Visible = true;

	//sheet 고유 넷하드 번호 노출
	Excel.Application.Worksheets(nCode).Activate;

	// sheet1내 작업영역을 객체에 저장
	oRange = Excel.Application.ActiveSheet.UsedRange;

	//oRange에는 현재 내용이 있는 부분만 선택됨.
	//oRange부분을 선택표시
	oRange.Select();
 
	//---------------------------------------------------------------------------------/

	//oRange부분을 브라우저에 출력
	var all = document.getElementById('dev_all'); //프리뷰 부분 찾고
	var allC = document.getElementById('dev_allC'); //코딩 부분 찾고
	all.innerHTML = '';	//초기화
	allC.innerHTML = '';	//초기화

	//---------------------------------------------------------------------------------/

	//상품평 이벤트 부분


	//그외 특수한 영역 판단
	var etc=''+
				'<!-- 그 외 특수한 경우 영역 상단 -->';
	var etc2=''+
				'<!-- 그 외 특수한 경우 영역 하단 -->';

	if(sCode.indexOf('sofa')!==-1 && outPut=='01_ismine'){ 
		etc=etc+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/01_ismine/01_common/04_etc/sofa/01.jpg" alt="LIVART SOFA IS...." />';

		if(sCode.indexOf('pu')==-1 && sCode.indexOf('fabric')==-1){ 
			etc2=etc2+
				'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/01_ismine/01_common/05_etc2/leather/01.jpg" alt="천연가죽소파" />';
		} //PU가 아닐때
	} //소파일때
	else if(sCode.indexOf('wood')!==-1 && outPut=='01_ismine'){ 
		etc2=etc2+
			'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/01_ismine/01_common/05_etc2/hardwood/01.jpg" alt="원목식탁" />';
	} //소파 아니면
	else{etc='';}

	// 바로가기 영역
	var go='';
	/*
	var go=new Array();
	if(sCode.indexOf('Ip')!==-1 || i=='199'){go[1]='';go[2]='';go[3]='';go[4]='';go[5]='';go[6]='';go[7]='';go[8]='';}
	else{
		go[1]='	<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/01.jpg" alt="제품 안내 바로가기" />';
		go[2]='	<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/02.jpg" alt="제품 상세페이지가 길어서 불편하셨죠?" />';
		go[3]='<a href="#spec"><img style="display:block;margin:0 auto;border:0;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/03.jpg" alt="제품스펙"/></a>';//스펙
		go[4]='<a href="#size"><img style="display:block;margin:0 auto;border:0;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/04.jpg" alt="상세사이즈"/></a>';//상세사이즈
		go[5]='<a href="#product"><img style="display:block;margin:0 auto;border:0;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/05.jpg" alt="제품기능"/></a>';//제품기능
		go[6]='<a href="#detail"><img style="display:block;margin:0 auto;border:0;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/06.jpg" alt="디자인디테일"/></a>'//디자인디테일
		go[7]='<a href="#module"><img style="display:block;margin:0 auto;border:0;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/07.jpg" alt="모듈"/></a>';//모듈
		go[8]='<a href="#ship"><img style="display:block;margin:0 auto;border:0;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go/08.jpg" alt="배송정보"/></a>';//배송정보
		//alert(sCode+' '+nCode);
	}
	*/
    //---------------------------------------------------------------------------------/

	var ex = new Array(); //ex값 개별 배열선언

	/* 
	ex[1][2] = oRange.Cells(1,2); //제품명
	ex[2][2] = oRange.Cells(2,2); //ERP

	*/
	//alert(ex[2]);
	var idgap;

	//로딩작업
	for(var e=1;e<=oRange.Rows.count;e++){ 
		ex[e] = new Array(); //ex[e]값 배열선언
		for(var j=1;j<=oRange.Columns.count;j++){ 
			ex[e][j] = oRange.Cells(e,j);
		}
	};

	//입력 작업이 끝난 후 코딩 조합
	for(var e=1;e<=oRange.Rows.count;e++){ 
		for(var j=1;j<=oRange.Columns.count;j++){ 
			if(e>=7){
				if(j>=5 && j<=24){
					if(String(ex[e][2])==String(ex[3][2])){
						// alert('ex['+e+']['+j+'] ='+ ex[e][3]);
						switch (String(ex[e][4])){
							//온라인 브랜드
							case '01': eCode='01_group/01_series';idgap=''; break; //시리즈
							case '02': eCode= String(ex[3][2]);idgap=''; break; //스펙
							case '03': eCode='01_group/03_product';idgap='id="product"'; break; //제품기능
							case '04': eCode='01_group/04_detail';idgap='id="detail"'; break; //디자인 디테일
							case '05': eCode='01_group/05_check';idgap='id="check"'; break; //체크사항
							case '06': eCode='01_group/06_special';idgap=''; break; //특별한 이유
							case '07': eCode='01_group/07_module';idgap='id="module"'; break; //모듈
							case '08': eCode='01_group/08_material';idgap=''; break; //소재&amp;컬러
							case '09': eCode='01_group/09_manual';idgap=''; break; //사용설명서
							case '10': eCode='01_group/10_MD';idgap=''; break; //MD 추천
							case '11': eCode='01_group/11_guide';idgap=''; break; //가이드
							case '12': eCode='01_group/12_upgrade';idgap=''; break; //업그레이드
							case '13': eCode='01_group/13_behindstory';idgap=''; break; //비하인드 스토리
							case '14': eCode='01_group/14_size';idgap=''; break; //디테일 사이즈
							case '15': eCode='01_group/15_review';idgap=''; break; //제품리뷰
							case '16': eCode='01_group/16_option';idgap=''; break; //옵션
							case '17': eCode='01_group/17_mix';idgap=''; break; //믹스&매치
							case '18': eCode='01_group/18_intro';idgap=''; break; //소개
							case '19': eCode='01_group/19_all';idgap=''; break; //제품전체
							case '20': eCode='01_group/20_concept';idgap=''; break; //디자인컨셉
							//default : alert('다시 설정 해주세요.');break; 
						};
						if(j==5){
						var ipgo=''; //입고 알림
						ipgo=''+
								//'<!-- 입고 알림 영역(공통) -->'+
								//'<img style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/02_series/'+sCode+'/01_group/00_ipgo/01.jpg" alt="입고 알림" />'+
								'';
						}

						sCode=String(ex[e][1]);
						//id값 할당
						if(String(ex[e][j])!=='01'){
							idgap=''; 
						}
						if(String(ex[e][4])=='02'){
							//idgap='id="spec"';
							if(ex[e][j]=='04'){ // if 값으로 id 값 할당 : spec
								idgap='id="spec"';						}
							else if(ex[e][j]=='11'){ // if 값으로 id 값 할당 : spec
								if(i=='103' || i=='105'){//홈스타일일 경우 size 바로가기 없음
									go[4]='<!--'+go[4]+'-->';
									idgap='';						}								
								else{
									idgap='id="size"';						}								
							}
							else{
								idgap='';
							}
						};
					}
					else if(String(ex[e][1])=='.'){}
					else{
						sCode= String(ex[e][1]);
						switch (String(ex[e][4])){
							//온라인 브랜드
							case '01': eCode='01_group/01_series';idgap=''; break; //시리즈
							case '02': eCode= String(ex[e][2]);idgap=''; break; //스펙
							case '03': eCode='01_group/03_product';idgap='id="product"'; break; //제품기능
							case '04': eCode='01_group/04_detail';idgap='id="detail"'; break; //디자인 디테일
							case '05': eCode='01_group/05_check';idgap='id="check"'; break; //체크사항
							case '06': eCode='01_group/06_special';idgap=''; break; //특별한 이유
							case '07': eCode='01_group/07_module';idgap='id="module"'; break; //모듈
							case '08': eCode='01_group/08_material';idgap=''; break; //소재&amp;컬러
							case '09': eCode='01_group/09_manual';idgap=''; break; //사용설명서
							case '10': eCode='01_group/10_MD';idgap=''; break; //MD 추천
							case '11': eCode='01_group/11_guide';idgap=''; break; //가이드
							case '12': eCode='01_group/12_upgrade';idgap=''; break; //업그레이드
							case '13': eCode='01_group/13_behindstory';idgap=''; break; //비하인드 스토리
							case '14': eCode='01_group/14_size';idgap=''; break; //디테일 사이즈
							case '15': eCode='01_group/15_review';idgap=''; break; //제품리뷰
							case '16': eCode='01_group/16_option';idgap=''; break; //옵션
							case '17': eCode='01_group/17_mix';idgap=''; break; //믹스&매치
							case '18': eCode='01_group/18_intro';idgap=''; break; //소개
							case '19': eCode='01_group/19_all';idgap=''; break; //제품전체
							case '20': eCode='01_group/20_concept';idgap=''; break; //디자인컨셉
							//default : alert('다시 설정 해주세요.');break; 
						};
						idgap='';
					};
					
					var z = String(ex[e][j].value);

					if(ex[e][j]=='.' || ex[e][j].value == undefined || ex[e][j].value == null || ex[e][j].value == false){ex[e][z]='';}
					else if(Right(ex[e][j],2)=='_r'){ //롤오버
						sum = sum+'<img style="display:block;margin:0 auto;" '+idgap+' src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/02_series/'+sCode+'/'+eCode+'/'+Left(ex[e][j],2)+'.jpg" onmouseover="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/02_series/'+sCode+'/'+eCode+'/'+Left(ex[e][j],2)+'_r.jpg\'"  onmouseout="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/02_series/'+sCode+'/'+eCode+'/'+Left(ex[e][j],2)+'.jpg\'" alt="'+ex[e][z]+'" />';	
						//sum = sum+'<img style="display:block;margin:0 auto;" '+idgap+' src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/02_series/'+sCode+'/'+eCode+'/'+Left(ex[e][j],2)+'.jpg" alt="" />';
					}
					else if(Left(ex[e][j],1)=='g'){ //gif
						sum = sum+'<img style="display:block;margin:0 auto;" '+idgap+' src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/02_series/'+sCode+'/'+eCode+'/g'+Right(ex[e][j],2)+'.gif" alt="'+ex[e][z]+'" />';	
					}
					else if(Left(ex[e][j],4)=='http'){ //동영상
						//alert("동영상 추가기능은 차후 개발 예정");
						// 차후 개발 예정 : 
						sum = sum+//'<div style="padding:20px 0;margin: 0px auto; max-width: 720px; background: #fff;"><iframe allowfullscreen="" frameborder="0" height="405" src="'+ex[e][j]+'" title="동영상"></iframe></div>';
											'<div style="max-width:720px;margin:0 auto;">'+
											'<style>.embed-container iframe, .embed-container object, .embed-container embed {position:absolute;top:0;left:0;width:100%;height:100%;}</style>'+
											'<div class="embed-container" style="position:relative;padding-bottom:56.25%;height:0;overflow:hidden;max-width:100%;">'+
											'	<iframe src="'+ex[e][j]+'" title="동영상" width="100%" height="405" frameborder="0" allowfullscreen></iframe></div>'+
											'</div>';
					}
					else{
						if(z=='c01'){z=0; ex[e][z]='';}
						else if(z=='.' || z == undefined || z == null || z == false){ex[e][z]='';}
						else{
							z = 4+Number(z);			
							z = 24+Number(ex[e][z]); //기존 이미지번호에 24 더함
							//alert('['+e+'] ['+z+']'+ex[e][z]); //ex[e][z]
						} 	
						//alert(z); //ex[e][z]  '+ex[ex[e][j]][j]+'
						if(String(ex[e][j]).length==1){
							//alert("이거 한자리수야");
							ex[e][j] = "0"+String(ex[e][j]);
						}
						sum = sum+'<img '+idgap+' style="display:block;margin:0 auto;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/02_series/'+sCode+'/'+eCode+'/'+ex[e][j]+'.jpg" alt="'+ex[e][z]+'" />';
					};

				}
			};
		}

		if(ex[e][5]=='.' || ex[e][5]=='c01'){
			switch (String(ex[e][4])){
				case '03': go[5]='<!--'+go[5]+'-->'; break; //제품기능 바로가기 주석 여부
				case '04': go[6]='<!--'+go[6]+'-->'; break; //디테일 바로가기 주석 여부
				case '07': go[7]='<!--'+go[7]+'-->'; break; //모듈 바로가기 주석 여부
			};
		}
	}

	var goFrom='';
	if(i=='206' || i=='106'){
		goFrom = '';
	}
	else {
		goFrom = ''+
			'<!-- 제품바로가기 공통 영역 -->'+
			'<div class="group">'+
			'<img style="margin: 0px auto; display: block;" alt="제품 안내 바로가기" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/01_ismine/01_common/03_go2/01.jpg"/>'+
			'<table cellpadding=0 cellspacing=0 border=0 style="max-width:720px; margin:0 auto">'+
			'<tr><td width="33.3%"><a href="#spec"><img style="width:100%;margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" alt="제품스펙" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go2/02.jpg"/></a></td>'+
			'<td width="33.3%"><a href="#size"><img style="width:100%;margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" alt="상세사이즈" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go2/03.jpg"/></a></td>'+
			'<td width="33.3%"><a href="#product"><img style="width:100%;margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" alt="제품기능" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go2/04.jpg"/></a></td></tr>'+
			'<tr><td><a href="#detail"><img style="width:100%;margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" alt="디자인디테일" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go2/05.jpg"/></a></td>'+
			'<td><a href="#module"><img style="display:block;width:100%;margin:0 auto;border:0;" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go2/06.jpg" alt="모듈" /></a></td>'+
			'<td><a href="#ship"><img style="width:100%;margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" alt="배송정보" src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/'+onOff+'/'+outPut+'/01_common/03_go2/07.jpg"/></a></td></tr>'+
			'</table>'+
			'';
	}

	go=''+
		goFrom+
		/* 
		go[1]+
		go[2]+
		go[3]+
		go[4]+
		go[5]+
		go[6]+
		go[7]+
		go[8]+
		*/
		'</div>'+
		'';
		
	//alert('ex['+1+']['+2+'] = '+ ex[1][2]);
	all.innerHTML = ''+
					//'<!-- <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'+
					//'<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">'+
					//'<head>'+
					//'<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />'+
					//'<link rel="stylesheet" type="text/css" href="http://mall.hyundailivart.co.kr/css/mobile/livart_m.css"/>'+
					'<!-- <title>'+ex[2][2]+'</title> -->'+
					//'</head>'+
					//'<body>'+
					//'<div class="detail_info_content"> -->'+
										
					//'<!-- // 상품기술서 시작 : 여기 아래로 복사하세요. -->'+
					'<div style="text-align:center;margin:0 auto;">'+
						ipgo+
						vow+
						'<!-- 상품기술서 영역  -->'+
						'<div class="group">'+ 
							
							etc+
							go+
							sum+
							etc2+
						'	<p style="display:block;overflow:hidden;height:0;">제품 안내에 대한 자세한 내용은 고객센터(T.1577-3332)에 문의주세요</p>'+    
						'</div>'+
						eve+
						ship+
					'</div>'+
					//'<!-- 상품기술서 끝 : 여기 위까지 복사하세요. // -->'+

					//'</div>'+
					//'</body>'+
					//'</html>'+
					'';					

	//치환작업
	/*reCode(eve);eve=temp;
	reCode(vow);vow=temp;
	reCode(etc);etc=temp
	reCode(go);go=temp;
	reCode(sum);sum=temp;
	reCode(ship);ship=temp;

	allC.innerHTML = ''+
					 '&lt;div style="width:100%;text-align:center;"&gt;'+
						vow+
						'&lt;!-- 상품기술서 영역  --&gt;'+
						'&lt;div class="group"&gt;'+ 
							etc+
							go+
							sum+
						'&lt;/div&gt;'+
						eve+
						ship+
					'&lt;/div&gt;'+
					'';  */
	



	reCode(all.innerHTML);
	allC.innerHTML= temp;					
					
}

