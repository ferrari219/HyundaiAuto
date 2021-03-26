//------------------------ 시작 -------------------------------//
//기본변수
var i; //펑션 숫자 변수
var k;
var brand = '';
//var loc = 'http://100.10.10.200/00_pd/';
var loc = 'http://100.0.0.1:8887/OAT/00_pd/';
//var loc = 'D:\/Web\/00_pd\/';
var loc2 = 'http://mall.hyundailivart.co.kr/UserFiles/Image/00_pd/';
var temp, tCnt;
var onOff;
var imgAtr = 'display:block; margin:0px auto;';

//var linkArea ='background-color:cyan;opacity:0.5;filter:alpha(opacity=50);';

//공통함수
function Left(Str, Num) {
	if (Num <= 0) {
		return '';
	} else if (Num > String(Str).length) {
		return Str;
	} else {
		return String(Str).substring(0, Num);
	}
}
function Right(Str, Num) {
	if (Num <= 0) {
		return '';
	} else if (Num > String(Str).length) {
		return Str;
	} else {
		var iLen = String(Str).length;
		return String(Str).substring(iLen, iLen - Num);
	}
}
//------------------------- 공통 영역 --------------------------------//

//------------------------- 기타 영역 --------------------------------//

// 코딩변경
function reCode(n) {
	temp = String(n);
	//temp = temp.replace(/\</gi, "&lt;"); //치환
	//temp = temp.replace(/\>/gi, "&gt;"); //치환
	temp = temp.replace(/\(\(/gi, '<strong>'); //치환
	temp = temp.replace(/\)\)/gi, '</strong>'); //치환
	temp = temp.replace(/\[{/gi, "<span class='point3'>"); //치환
	temp = temp.replace(/\}]/gi, '</span>'); //치환
	temp = temp.replace(/\[/gi, "<span class='point1'>"); //치환
	temp = temp.replace(/\]/gi, '</span>'); //치환
	temp = temp.replace(/\{/gi, "<span class='point2'>"); //치환
	temp = temp.replace(/\}/gi, '</span>'); //치환
	temp = temp.replace(/\;/gi, '<br/>'); //치환
	n = temp;
	//temp = temp.replace(/\.jpg\"&gt;/gi, '.jpg\"/&gt;'); //치환
	return n;
}
function reTag(n) {
	temp = String(n);
	temp = temp.replace(/\</gi, '&lt;'); //치환
	temp = temp.replace(/\>/gi, '&gt;'); //치환
	n = temp;
	return n;
}
function jpgPng(n) {
	temp = String(n);
	temp = temp.replace(/jpg/gi, 'png');
	return temp;
}
function reAdrs(n) {
	temp = String(n);
	temp = temp.replace(/\/\/100.10.10.200\/Lguidesystem\/pd\//gi, 'http://mall.hyundailivart.co.kr/UserFiles/Image/00_pd/');
	temp = temp.replace(/png/gi, 'jpg');
	return temp;
}
// 텍스트 변경
function txtCng(k) {
	k = String(k);
	k = k.replace(/\_/gi, '&nbsp;&nbsp;');
	k = k.replace(/\;/gi, '<br/>'); //치환
	return k;
}
//var test = Address('//100.10.10.200/Lguidesystem/00_pd/');
//alert(test)

//------------------------- 엑셀 영역 --------------------------------//
//엑셀 불러오기
function excel(Location, i) {
	//엑셀이 입력될 타겟, 기획전타입번호
	var sCode = document.frm.series.value; //document.getElementById('dev_tbl').getElementsByTagName('input')[2*(k-1)+0].value;//시리즈코드 인풋값 반영
	var nCode = document.frm.numb.value; //document.getElementById('dev_tbl').getElementsByTagName('input')[2*(k-1)+1].value;//넷하드 상품번호 인풋값 반영
	var sum = '',
		sumS = '';

	if (sCode === 'basic') {
		alert('basic은 사용할 수 없습니다. 다른 브랜드명을 사용해주세요.');
	}
	//---------------------------------상하단 공통영역(개발중)------------------------------------/
	function etc1(k) {
		if (Left(k, 1) == '1') {
			return (onOff = 'on');
		} else {
			return (onOff = 'off');
		}
	}
	etc1(i);

	switch (i) {
		//온라인 브랜드
		case '101':
			outPut = '01';
			break; //이즈마인
		case '102':
			outPut = '02';
			break; //시스템주방
		case '103':
			outPut = '03';
			break; //홈스타일
		case '104':
			outPut = '04';
			break; //하움
		case '105':
			outPut = '05';
			break; //네오스 온라인
		case '106':
			outPut = '06';
			break; //H몬도
		case '107':
			outPut = '07';
			break; //리바트체어스
		case '108':
			outPut = '08';
			break; //ICT
		case '109':
			outPut = '09';
			break; //몰인몰
		case '199':
			outPut = '99';
			break; //etc

		//오프라인 브랜드
		case '201':
			outPut = '01';
			break; //리바트
		case '202':
			outPut = '02';
			break; //리첸
		case '203':
			outPut = '03';
			break; //네오스
		case '204':
			outPut = '04';
			break; //키친
		case '205':
			outPut = '05';
			break; //키즈
		case '206':
			outPut = '06';
			break; //하움

		default:
			j = 99;
			alert('다시 설정 해주세요.');
			break;
	}

	//상품평 이벤트
	var eve = new Array();
	eve[110] = '' + '<!-- 상품평 이벤트 영역(공통) -->' + '	<img style="' + imgAtr + '" alt="temp" src="' + loc2 + onOff + '/' + outPut + '/01_cmn/00_top/01.jpg"/>' + '';

	var vow = new Array();
	vow[110] = '' + '<!-- WSI : 브랜드 소개 영역(공통) -->' + '<img name="brand" id="brand" style="margin: 50px auto;display:block;" alt="윌리엄소노마 브랜드" src="' + loc2 + onOff + '/' + outPut + '/01_cmn/99_brand/01.jpg"/>' + '';

	//배송안내 WSI
	var ship = new Array();
	ship[110] = '<!-- WSI : 배송안내 영역(공통)  -->' + '<img name="ship" id="ship" style="margin: 50px auto;display:block;" alt="윌리엄소노마 배송안내" src="' + loc2 + onOff + '/' + outPut + '/01_cmn/99_ship/01.jpg"/>' + '';

	//i값
	i = Number(i); //숫자화
	function etc2(k) {
		return (eve = eve[k]);
		return (vow = vow[k]);
		return (ship = ship[k]);
	}
	etc2(i);

	//---------------------------------------------------------------------------------/

	//엑셀용 변수 선언
	var Excel, oRange;

	// 엑셀 객체생성
	Excel = new ActiveXObject('Excel.Application');
	//Excel.Application.Workbooks.Open('\\\\100.10.10.200\\Lguidesystem\\00_pd\\'+onOff+'\\'+outPut+'\\'+sCode+'.xls');
	Excel.Application.Workbooks.Open('D:/OAT/00_pd/' + onOff + '/' + outPut + '/' + sCode + '.xls');
	//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/00_pd/'+onOff+'/'+outPut+'/'+sCode+'.xls');

	//---------------------------------고유번호영역------------------------------------/

	Excel.Application.Visible = true; //엑셀표시

	Excel.Application.Worksheets('00').Activate; //sheet 고유 넷하드 번호 노출
	oRange0 = Excel.Application.ActiveSheet.UsedRange; // sheet1내 작업영역을 객체에 저장

	Excel.Application.Worksheets(nCode).Activate; //sheet 고유 넷하드 번호 노출
	oRange = Excel.Application.ActiveSheet.UsedRange; // sheet1내 작업영역을 객체에 저장
	//oRange부분을 선택표시
	oRange.Select(); //oRange에는 현재 내용이 있는 부분만 선택됨.

	//---------------------------------------------------------------------------------/

	//oRange부분을 브라우저에 출력
	var all = document.getElementById('dev_allp'); //프리뷰 부분 찾고
	var allC = document.getElementById('dev_allC'); //코딩 부분 찾고
	var allS = document.getElementById('dev_allS'); //슬라이스용 부분 찾고
	all.innerHTML = ''; //초기화
	allC.innerHTML = ''; //초기화
	allS.innerHTML = ''; //초기화

	//---------------------------------------------------------------------------------/

	//그외 특수한 영역 판단

	// 바로가기 영역
	var go = '';

	//---------------------------------------------------------------------------------/

	var ex = new Array(); //ex값 개별 배열선언
	var exS = new Array(); //ex값 00페이지 개별 배열선언
	//var sum0 = ''; //00페이지 합계
	var series = '',
		seriesC = '',
		product = '',
		productC = '',
		detail = '',
		detailC = '',
		check = '',
		checkC = '';
	var special = '',
		specialC = '',
		module = '',
		moduleC = '',
		material = '',
		materialC = '',
		manual = '',
		manualC = '',
		MD = '',
		MDC = '';
	var guide = '',
		guideC = '',
		upgrade = '',
		upgradeC = '',
		behind = '',
		behindC = '',
		sizeT = '',
		sizeTC = '',
		review = '',
		reviewC = '';
	var option = '',
		optionC = '',
		mix = '',
		mixC = '',
		intro = '',
		introC = '',
		allT = '',
		allTC = '',
		concept = '',
		conceptC = '';
	var color = '',
		colorC = '';
	var price = '';

	var idgap, line;
	var sumSethtml = ''; //new Set();

	//function
	var Temp = ''; //임시공간
	var h3 = '',
		dt = '',
		dd = '',
		tag = '',
		sup = '',
		pic = '';
	//타입분류
	function picFn(p, m) {
		//alert(p);
		if (Right(p, 3) === 'mp4') {
			p = String(p);
			p = p.replace(/\mp4/gi, 'gif'); //치환
			alert('s1-3, s1-4의 경우 동영상 첨부가 불가능합니다. gif로 등록해주세요');
		}
		if (p !== '') {
			//alert(p);
			p = jpgPng(p);
			pic =
				pic +
				'		<img src="' +
				loc +
				onOff +
				'/' +
				outPut +
				'/' +
				sCode +
				'/group/' +
				m +
				'/' +
				p +
				'" title="' +
				loc +
				onOff +
				'/' +
				outPut +
				'/' +
				sCode +
				'/group/' +
				m +
				'/' +
				p +
				'" onerror="this.src=\'' +
				loc +
				'/basic/group/03/v01.' +
				Right(p, 3) +
				'\'" alt="">' +
				'';
		}
		return pic;
	}
	function mp4Fn(p) {
		//동영상 함수
		if (p !== '') {
			//alert(p);
			pic =
				pic +
				'<div class="mp4" style="margin:0 auto;">' +
				'	<style>.embed-container iframe, .embed-container object, .embed-container embed {position:absolute;top:0;left:0;width:100%;height:100%;}</style>' +
				'	<div class="embed-container" style="position:relative;padding-bottom:56.25%;height:0;overflow:hidden;max-width:100%;">' +
				'	<iframe src="' +
				p +
				'" title="동영상" width="100%" height="440" frameborder="0" allowfullscreen></iframe></div>' +
				'</div>' +
				'';
		}
		return pic;
	}
	function tType(v, t) {
		//개별 모듈단위로 묶기
		if (t !== '') {
			switch (String(v)) {
				case 'h3':
					h3 = h3 + '		<h3>' + t + '</h3>' + '';
					return h3;
					break;
				case 'dt':
					dt = dt + '			<dt>' + t + '</dt>' + '';
					//alert(dt);
					return dt;
					break;
				case 'dd':
					dd = dd + '			<dd>' + t + '</dd>' + '';
					//alert(dd);
					return dd;
					break;
				case 'sup':
					sup = sup + '		<p class="subTxt">* ' + t + '</p>' + '';
					return sup;
					break;
				case 'tag':
					tag = tag + '		' + t + ' ' + '';
					return tag;
					break;
			}
		}
	}
	function tTypeSum(v, t) {
		//dl만 묶는 형태
		if (t !== '') {
			switch (String(v)) {
				case 'h3':
					h3 = h3 + '		<h3>' + t + '</h3>' + '';
					return h3;
					break;
				case 'dt':
					dl = dl + '		<dl>' + '			<dt>' + t + '</dt>' + '';
					//alert(dt);
					break;
				case 'dd':
					dl = dl + '			<dd>' + t + '</dd>' + '		</dl>' + '';
					//alert(dd);
					break;
				case 'sup':
					sup = sup + '		<p class="subTxt">* ' + t + '</p>' + '';
					return sup;
					break;
			}
			return dl;
		}
	}

	function making(judge, tType1, txt1, tType2, txt2, tType3, txt3, tType4, txt4, tType5, txt5, tType6, txt6, tType7, txt7, bg, jpg, mp4, pic1, pic2, pic3, pic4, mod) {
		h3 = '';
		dt = '';
		dd = '';
		tag = '';
		sup = '';
		pic = '';
		tCnt = 0;
		dl = '';
		//alert(mod);
		//custom 분류
		// if(Left(judge,6) === 'custom'){
		// 	tTypeSum(String(tType1), String(txt1));tTypeSum(String(tType2), String(txt2));
		// 	tTypeSum(String(tType3), String(txt3));tTypeSum(String(tType4), String(txt4));
		// 	tTypeSum(String(tType5), String(txt5));tTypeSum(String(tType6), String(txt6));
		// 	tTypeSum(String(tType7), String(txt7));
		// }
		// else{
		tType(String(tType1), String(txt1));
		tType(String(tType2), String(txt2));
		tType(String(tType3), String(txt3));
		tType(String(tType4), String(txt4));
		tType(String(tType5), String(txt5));
		tType(String(tType6), String(txt6));
		//}

		//	judge,
		// 	tType1,txt1,tType2,txt2,tType3,txt3,tType4,txt4,tType5,txt5,tType6,txt6,tType7,txt7,
		//	bg,jpg,mp4,
		// 	pic1,pic2,pic3,pic4,
		//	 mod

		//동영상이 아니거나 좌우형인 경우
		if (jpg !== 'mp4' || judge === String('s1-3') || judge === String('s1-4')) {
			picFn(String(pic1), mod);
			picFn(String(pic2), mod);
			picFn(String(pic3), mod);
			picFn(String(pic4), mod);
			//이미지 뚜껑
			if (pic !== '') {
				pic = '' + '	<div class="pic">' + pic + '	</div>' + '';
			}
		} else {
			//동영상인 경우
			//alert("mp4");
			mp4Fn(String(mp4));
		}

		//alert(target);
		switch (String(judge)) {
			case 's1-1':
				Temp = '' + '<div class="s1-1 ' + bg + '">' + '	<div class="text">' + h3 + '		<dl>' + dt + dd + '		</dl>' + sup + '	</div>' + pic + '</div>' + '';
				tCnt = 7;
				break;
			case 's1-2':
				Temp = '' + '<div class="s1-2 ' + bg + '">' + pic + '	<div class="text">' + h3 + '		<dl>' + dt + dd + '		</dl>' + sup + '	</div>' + '</div>' + '';
				tCnt = 7;
				break;
			case 's1-3':
				Temp = '' + '<div class="s1-3 ' + bg + '">' + pic + '	<div class="text">' + h3 + '		<dl>' + dt + dd + '		</dl>' + sup + '	</div>' + '</div>' + '';
				tCnt = 7;
				break;
			case 's1-4':
				Temp = '' + '<div class="s1-4 ' + bg + '">' + '	<div class="text">' + h3 + '		<dl>' + dt + dd + '		</dl>' + sup + '	</div>' + pic + '</div>' + '';
				tCnt = 7;
				break;
			case 'custom1-1':
				if (dt == '' && dd == '') {
					dl = '';
				} else {
					dl = '' + '			<dl>' + dt + dd + '			</dl>' + '';
				}
				Temp =
					'' +
					'<div class="custom">' +
					pic +
					'	<div class="text">' +
					//'		<h3>- 엑셀 파일 참고 -</h3>'+
					h3 +
					'		<div class="boxWrapper">' +
					dl +
					'		</div>' +
					sup +
					'	</div>' +
					'</div>' +
					'';
				tCnt = 7;
				break;
			case 'custom1-2':
				Temp =
					'' +
					'<div class="custom">' +
					'	<div class="text">' +
					// '		<h3>엑셀 파일 참고</h3>'+
					h3 +
					'		<div class="boxWrapper">' +
					'			<dl>' +
					dt +
					dd +
					'			</dl>' +
					'		</div>' +
					sup +
					'	</div>' +
					pic +
					'</div>' +
					'';
				tCnt = 7;
				break;
			case 's0':
				Temp = '' + '<section class="price">' + '	<dl>' + dt + dd + '	</dl>' + '</section>' + '';
				tCnt = 7;
				break;
		}
		return Temp, tCnt;
	}

	//Slice용 이미지 출력시에만
	if (Left(nCode, 2) === '00') {
		//00페이지 로딩작업
		for (var e = 41; e <= oRange0.Rows.count; e++) {
			exS[e] = new Array(); //ex[e]값 배열선언
			for (var j = 28; j <= oRange0.Columns.count; j++) {
				exS[e][j] = oRange0.Cells(e, j);
				//alert(ex[e][j])
				if (exS[e][j] == '.' || exS[e][j].value == undefined || exS[e][j].value == null || exS[e][j].value == false) {
					exS[e][j] = '';
				} else {
					exS[e][j] = reCode(exS[e][j]);
				}
			}
		}

		//정보 분류 작업
		//alert(oRange0.Rows.count);
		for (e = 41; e <= oRange0.Rows.count; e++) {
			//alert(e);
			tCnt = 0; //alert(exS[e][30]);
			if (e < oRange0.Rows.count) {
				//커스톰 분류
				//alert(Left(exS[e][28],2));
				// if(Left(exS[e+2][28],6) === 'custom'){
				// 	making(String(exS[e+2][28]),
				// 		exS[e+1][30],exS[e+1][31]+exS[e+1][32],exS[e+2][30],exS[e+2][31]+exS[e+2][32],exS[e+3][30],exS[e+3][31]+exS[e+3][32],exS[e+4][30],exS[e+4][31]+exS[e+4][32],exS[e+5][30],exS[e+5][31]+exS[e+5][32],
				// 		exS[e+6][30],exS[e+6][31]+exS[e+6][32],exS[e+7][30],exS[e+7][31]+exS[e+7][32],
				// 		exS[e+3][28], exS[e+1][28], exS[e+4][28],
				// 		exS[e][33], exS[e][34], exS[e][35], exS[e][36],
				// 		Left(exS[e][28],2)
				// 	);
				// }
				// else{
				making(
					String(exS[e + 2][28]),
					exS[e + 1][30],
					exS[e + 1][31] + exS[e + 1][32],
					exS[e + 2][30],
					exS[e + 2][31] + exS[e + 2][32],
					exS[e + 3][30],
					exS[e + 3][31] + exS[e + 3][32],
					exS[e + 4][30],
					exS[e + 4][31] + exS[e + 4][32],
					exS[e + 5][30],
					exS[e + 5][31] + exS[e + 5][32],
					exS[e + 6][30],
					exS[e + 6][31] + exS[e + 6][32],
					'',
					'',
					exS[e + 3][28],
					exS[e + 1][28],
					exS[e + 4][28],
					exS[e][33],
					exS[e][34],
					exS[e][35],
					exS[e][36],
					Left(exS[e][28], 2)
				);
				//}

				/*making(judge, 
					tType1,txt1,tType2,txt2,tType3,txt3,tType4,txt4,tType5,txt5, 
					tType6,txt6,tType7,txt7,
					bg, jpg, mp4,
					pic1, pic2, pic3, pic4,
					mod
				)
				*/
			}

			//if(Temp==undefined){ Temp = ''; }

			switch (Left(String(exS[e][28]), 2)) {
				case '01':
					series += Temp;
					break;
				case '03':
					product += Temp;
					break;
				case '04':
					detail += Temp;
					break;
				case '05':
					check += Temp;
					break;
				case '06':
					special += Temp;
					break;
				case '07':
					module += Temp;
					break;
				case '08':
					material += Temp;
					break;
				case '09':
					manual += Temp;
					break;
				case '10':
					MD += Temp;
					break;
				case '11':
					guide += Temp;
					break;
				case '12':
					upgrade += Temp;
					break;
				case '13':
					behind += Temp;
					break;
				case '14':
					sizeT += Temp;
					break;
				case '15':
					review += Temp;
					break;
				case '16':
					option += Temp;
					break;
				case '17':
					mix += Temp;
					break;
				case '18':
					intro += Temp;
					break;
				case '19':
					allT += Temp;
					break;
				case '20':
					concept += Temp;
					break;
				// case '21':
				// 	color += Temp;
				// 	break;
				case '99':
					price += Temp;
					break;
			}

			e += tCnt;
		}

		//뚜껑 재조합
		series = '' + '<div class="series">' + series + '</div>' + '';
		product = '' + '<div class="product">' + '	<div class="title"><h2><em>제품기능</em> <span>Product function</span></h2></div>' + product + '</div>' + '';
		detail = '' + '<div class="detail">' + '	<div class="title"><h2><em>디자인 디테일</em> <span>Design detail</span></h2></div>' + detail + '</div>' + '';
		check = '' + '<div class="check">' + '	<div class="title"><h2><em>체크사항</em> <span>Check</span></h2></div>' + check + '</div>' + '';
		special = '' + '<div class="special">' + '	<div class="title"><h2><em>특별한 이유</em> <span>Special reason</span></h2></div>' + special + '</div>' + '';
		module = '' + '<div class="module">' + '	<div class="title"><h2><em>모듈</em> <span>module</span></h2></div>' + module + '</div>' + '';
		material = '' + '<div class="material">' + '	<div class="title"><h2><em>소재&amp;컬러</em> <span>Material&amp;Color</span></h2></div>' + material + '</div>' + '';
		manual = '' + '<div class="manual">' + '	<div class="title"><h2><em>설명서</em> <span>Manual</span></h2></div>' + manual + '</div>' + '';
		MD = '' + '<div class="MD">' + '	<div class="title"><h2><em>MD 추천</em> <span>MD recommendation</span></h2></div>' + MD + '</div>' + '';
		guide = '' + '<div class="guide">' + '	<div class="title"><h2><em>가이드</em> <span>Guide</span></h2></div>' + guide + '</div>' + '';
		upgrade = '' + '<div class="upgrade">' + '	<div class="title"><h2><em>업그레이드</em> <span>Design upgrade</span></h2></div>' + upgrade + '</div>' + '';
		behind = '' + '<div class="behind">' + '	<div class="title"><h2><em>비하인드 스토리</em> <span>Behind story</span></h2></div>' + behind + '</div>' + '';
		sizeT = '' + '<div class="sizeT">' + '	<div class="title"><h2><em>디테일 사이즈</em> <span>Detail size</span></h2></div>' + sizeT + '</div>' + '';
		review = '' + '<div class="review">' + '	<div class="title"><h2><em>제품리뷰</em> <span>Design review</span></h2></div>' + review + '</div>' + '';
		option =
			'' +
			'<div class="option">' +
			'	<div class="title"><h2><em>옵션 디테일컷</em> <span>option</span></h2></div>' +
			'		<img src="' +
			loc +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/16/02.png" title="' +
			loc +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/16/02.png" onerror="this.src=\'' +
			loc +
			'basic/group/16/02.jpg\'" style="' +
			imgAtr +
			'" alt="">' +
			//option+
			'</div>' +
			'';
		mix = '' + '<div class="mix">' + '	<div class="title"><h2><em>믹스</em> <span>mix</span></h2></div>' + mix + '</div>' + '';
		intro = '' + '<div class="intro">' + '	<div class="title"><h2><em>인트로</em> <span>intro</span></h2></div>' + intro + '</div>' + '';
		allT = '' + '<div class="allT">' + '	<div class="title"><h2><em>전체 시리즈</em> <span>All series</span></h2></div>' + allT + '</div>' + '';
		concept = '' + '<div class="concept">' + '	<div class="title"><h2><em>컨셉</em> <span>concept</span></h2></div>' + concept + '</div>' + '';
		// color = ''+
		// 			'<div class="color">'+
		// 			'	<div class="title"><h2><em>옵션 디테일</em> <span>Option detail</span></h2></div>'+
		// 			color+
		// 			'</div>'+
		// 			'';

		//alert(product);
		//출력
		all.innerHTML =
			'' +
			//'<!-- <title>'+ex[13][2]+'</title> -->'+
			'<!-- 상품기술서 시작 입니다. -->' +
			'	<div class="pdtDetail" style="max-width:780px;margin:0 auto;">' +
			price +
			series +
			product +
			detail +
			check +
			special +
			module +
			material +
			manual +
			MD +
			guide +
			upgrade +
			behind +
			sizeT +
			review +
			option +
			mix +
			intro +
			allT +
			concept +
			//color+
			//'	<br/>'+
			//sum0+
			'	</div>' +
			//'<div class="capture"></div>'+
			'<!-- // 상품기술서 끝 입니다. -->' +
			'';
	} else {
		//로딩작업
		for (var e = 13; e <= oRange.Rows.count; e++) {
			ex[e] = new Array(); //ex[e]값 배열선언
			for (var j = 1; j <= oRange.Columns.count; j++) {
				ex[e][j] = oRange.Cells(e, j);
				//alert(ex[e][j])
				if (ex[e][j] == '.' || ex[e][j].value == undefined || ex[e][j].value == null) {
					ex[e][j] = '';
				} else {
					if (Left(ex[e][j], 4) == 'http') {
						ex[e][j] = ex[e][j];
					} else {
						if (ex[e][2] !== '11') {
							ex[e][j] = txtCng(ex[e][j]);
							//alert(ex[e][j]);
						}
					}
				}
			}
			//코딩 모듈영역 구하기
			if (ex[e][1] == '코딩 모듈 영역') {
				//59
				line = e;
			}
		}

		var top = ''; //메인 영역
		var opt = '',
			clropt = ''; //컬러,사이즈 영역 //
		var spec = '',
			specC = '';
		var cfg = '',
			std = '',
			stdCol = '',
			stdRow = '',
			stdCnt = 0,
			clsm = '',
			clsmCol = '',
			clsmColS = '',
			clsmTit = '',
			cateICO = '',
			org = '',
			orgTit = '',
			orgCol = '',
			mfs = '',
			sz = '',
			cert = '',
			ctn = ''; //스펙 영역
		var view = ''; //코딩 뷰 영역
		var specAlt = '';

		//입력 작업이 끝난 후 코딩 조합: 스펙영역
		for (e = 27; e <= line - 1; e++) {
			if (ex[e][5] == '' && ex[e][6] == '') {
			} else {
				switch (String(ex[e][2])) {
					case '01': //타이틀
						top =
							'' +
							'<div class="Box">' +
							'	<div class="headBox">' +
							'		<h1>' +
							ex[e][4] + //NewT
							'			<span>Series</span>' +
							'		</h1>' +
							'		<p>' +
							ex[e][5] +
							'</p>' + //뉴트 2400 거울 옷장세트 B
							'	</div>' +
							'</div>' +
							'';
						specAlt = specAlt + ex[e][5] + '/' + '';
						break;
					case '02':
						// for(j=6;j<=14;j+=2){//옵션값만 확인 //컬러값 확장시 j<=20
						if (ex[e][6] == undefined || ex[e][6] == '') {
						} else {
							//옵션값 부여
							//alert(ex[e][j]);
							if (ex[e][5] == undefined || ex[e][5] == '') {
								clropt =
									clropt +
									// opt = opt +
									'			<dd>' +
									ex[e][5] +
									'</dd>' +
									'';
							} else {
								clropt =
									clropt +
									// opt = opt +
									'		<dd>' +
									'			<em style="background:' +
									ex[e][6] +
									'"></em>' +
									'			<ul><li>' +
									ex[e][5] +
									'</li></ul>' +
									'		</dd>' +
									'';
							}
							specAlt = specAlt + ex[e][6 - 1] + '/' + '';
						}
						// };
						break;
					case '05': //구성 cfg
						cfg = cfg + '<dl class="s1">' + '	<dt>구성</dt>' + '	<dd>' + ex[e][5] + '		<span>' + ex[e][15] + '</span>' + '	</dd>' + '</dl>';
						'';
						break;
					case '06': //규격 std
						switch (
							String(ex[e][3]) //레이아웃 결정
						) {
							case 's1':
								std = std + '<div class="s1">' + '	<dl>' + '		<dt>규격</dt>' + '		<dd>' + ex[e][4] + ' ' + ex[e][5] + ' ' + ex[e][6] + ' ' + ex[e][7] + ' ' + ex[e][8] + ' ' + ex[e][9] + ' ' + ex[e][10] + ' mm</dd>' + '	</dl>' + '</div>' + '';
								specAlt = specAlt + ex[e][4] + ': ' + ex[e][5] + ' ' + ex[e][6] + ' ' + ex[e][7] + ' ' + ex[e][8] + ' ' + ex[e][9] + ' ' + ex[e][10] + ' mm/' + '';
								break;
							case 's2-1':
								for (k = 0; k < 2; k++) {
									//소제목 더하기
									for (l = 0; l < 6; l = l + 2) {
										//옵션명 더하기
										if (ex[e + k][l + 6] !== '') {
											stdCol = stdCol + '		<dd>' + '			<em>' + ex[e + k][l + 5] + '</em>' + '			<span>' + ex[e + k][l + 6] + '</span>' + '		</dd>' + '';
										}
									}
									if (ex[e + k][4] !== '') {
										stdRow = stdRow + '	<dl>' + '		<dt>' + ex[e + k][4] + ' <span>(mm)</span></dt>' + stdCol + '	</dl>' + '';
										stdCol = '';
									}
								}
								std = std + '<div class="s2-1">' + stdRow + '</div>' + '';
								stdRow = '';
								specAlt = specAlt + ex[e][4] + ':' + ex[e][6] + 'x' + ex[e][8] + 'x' + ex[e][10] + 'mm/' + '';
								e = e + 1;
								break;
							case 's3-1':
								for (k = 0; k < 3; k++) {
									//소제목 더하기
									for (l = 0; l < 6; l = l + 2) {
										//옵션명 더하기
										if (ex[e + k][l + 6] !== '') {
											stdCol = stdCol + '		<dd>' + '			<em>' + ex[e + k][l + 5] + '</em>' + '			<span>' + ex[e + k][l + 6] + '</span>' + '		</dd>' + '';
										}
									}
									if (ex[e + k][4] !== '') {
										stdRow = stdRow + '	<dl>' + '		<dt>' + ex[e + k][4] + ' <span>(mm)</span></dt>' + stdCol + '	</dl>' + '';
										stdCol = '';
									}
								}
								std = std + '<div class="s3-1">' + stdRow + '</div>' + '';
								stdRow = '';
								specAlt = specAlt + ex[e][4] + ':' + ex[e][6] + 'x' + ex[e][8] + 'x' + ex[e][10] + 'mm/' + '';
								e = e + 1;
								break;
						}
						break;

					case '07': //마감소재 clsm
						if (String(ex[e][4]) === String(ex[e - 1][4])) {
							//침대 === 침대
							//alert('ex['+e+'] = em 비워');
							clsmTit = '';
							clsmCol = '';
						} else if (String(ex[e][4]) === '') {
							//alert(ex[e][4]);
							clsmTit = '' + '<dd>' + '';
							//alert('ex['+e+'] = ul열어');
							clsmCol = '' + '	<ul>' + clsmCol + '';
						} else {
							//침대!==책상
							//alert('ex['+e+'] = em 더하기');
							switch (String(ex[e][4])) {
								case '침대':
									cateICO = 'edge00';
									break;
								case '책상':
									cateICO = 'edge01';
									break;
								case '책장':
									cateICO = 'edge02';
									break;
								case '드레스룸':
									cateICO = 'edge03';
									break;
								case '옷장':
									cateICO = 'edge04';
									break;
								case '매트리스':
									cateICO = 'edge05';
									break;
								case '붙박이장':
									cateICO = 'edge06';
									break;
								case '서랍장':
									cateICO = 'edge07';
									break;
								case '화장대':
									cateICO = 'edge08';
									break;
								case '소파':
									cateICO = 'edge09';
									break;
								case '거실장':
									cateICO = 'edge10';
									break;
								case '식탁':
									cateICO = 'edge11';
									break;
								case '의자':
									cateICO = 'edge12';
									break;
								case '수납장':
									cateICO = 'edge14';
									break;
								default:
									cateICO = 'edge01 custom';
									break;
							}

							clsmTit = '' + '<dd>' + '	<em class="icon ' + cateICO + '">' + ex[e][4] + '</em>' + '';
							//alert('ex['+e+'] = ul열어');
							clsmCol = '' + '	<ul>' + clsmCol + '';
						}

						//alert('ex['+e+'] = li 사용');
						if (ex[e][5] == '') {
							clsmColS = '';
						} else {
							if (String(ex[e][5]).length >= 4) {
								clsmColS = '			<strong class="let3">' + ex[e][5] + '</strong>';
							} else {
								clsmColS = '			<strong>' + ex[e][5] + '</strong>';
							}
						}

						clsmCol = clsmCol + '		<li>' + clsmColS + '			<span>' + ex[e][6] + ' ' + ex[e][7] + ' ' + ex[e][8] + ' ' + ex[e][9] + ' ' + ex[e][10] + '</span>' + '		</li>' + '';

						if (String(ex[e][4]) !== String(ex[e + 1][4]) || String(ex[e][1]) !== String(ex[e + 1][1])) {
							//alert('ex['+e+'] = ul, dd닫아');
							clsmCol = '' + clsmCol + '	</ul>' + '</dd>' + '';
						}
						//alert(clsmCol);
						clsm = clsm + clsmTit + clsmCol + '';
						clsmCol = '';
						specAlt = specAlt + ex[e][4] + ':' + ex[e][6] + ' ' + ex[e][8] + ' ' + ex[e][10] + '/' + '';
						break;
					case '08': //원산지 org
						org = org + '<p>' + ex[e][5] + ' ' + ex[e][6] + ' ' + ex[e][7] + ' ' + ex[e][8] + ' ' + ex[e][9] + ' ' + ex[e][10] + '</p>' + '';
						specAlt = specAlt + ex[e][5] + ' ' + ex[e][6] + ' ' + ex[e][7] + ' ' + ex[e][8] + ' ' + ex[e][9] + ' ' + ex[e][10] + '/' + '';
						break;
					case '09': //제조원 mfs
						mfs =
							mfs +
							//'제조원<br/>'+
							'<p>' +
							ex[e][5] +
							' ' +
							ex[e][6] +
							' ' +
							ex[e][7] +
							' ' +
							ex[e][8] +
							' ' +
							ex[e][9] +
							' ' +
							ex[e][10] +
							'</p>' +
							'';
						specAlt = specAlt + ex[e][5] + ' ' + ex[e][6] + ' ' + ex[e][7] + ' ' + ex[e][8] + ' ' + ex[e][9] + ' ' + ex[e][10] + '/' + '';
						break;
					case '11': //사이즈 sz
						sz =
							sz +
							//'사이즈<br/>'+
							'<li>- ' +
							ex[e][5] +
							' ' +
							ex[e][6] +
							' ' +
							ex[e][7] +
							' ' +
							ex[e][8] +
							' ' +
							ex[e][9] +
							' ' +
							ex[e][10] +
							'</li>' +
							'';
						break;
					case '12': //인증 cert
						cert = cert + '<dl class="s1">' + '	<dt>인증</dt>' + '	<dd>' + ex[e][5] + '</dd>' + '</dl>';
						'';
						break;
					case '13': //주의 ctn
						var ctnTemp = '';
						if (ex[e][5] == '' || ex[e][5] == '-') {
							ctnTemp = '';
						} else if (ex[e][5] == '선반하중') {
							ctnTemp = '' + '<img src="' + loc2 + onOff + '/' + outPut + '/01_cmn/cau01.jpg" alt="선반하중 주의사항" style="' + imgAtr + '"/>' + '';
							// alert(ctnTemp);
						} else if (ex[e][5] == '벽고정') {
							ctnTemp = '' + '<img src="' + loc2 + onOff + '/' + outPut + '/01_cmn/cau02.jpg" alt="벽고정필수 주의사항" style="' + imgAtr + '"/>' + '';
							// alert(ctnTemp);
						} else {
						}

						ctn = ctn + ctnTemp + '';
						break;
				}
			}
		}
		//입력 작업이 끝난 후 코딩 조합

		var sup = '';

		if (ex[31][11] == '' || ex[31][11] == '-') {
		} else {
			sup =
				sup +
				'	<p class="sup"><span>' +
				txtCng(ex[31][11]) +
				'</span><i class="bg"></i></p>' + //연출용 사진으로 소품은 // ex[e][21]
				'';
		}

		opt =
			'' +
			'<div class="opt">' +
			'	<h3>옵션<span>' +
			ex[e][4] +
			'</span></h3>' +
			'	<p><img src="' +
			loc +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/' +
			nCode +
			'/v01.png" title="' +
			loc +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/' +
			nCode +
			'/v01.png" onerror="this.src=\'' +
			loc +
			'/basic/01/v01.png\'" alt="" /></p>' +
			sup +
			'	<div class="sltOpt">' +
			'		<ul>' +
			//opt+
			'		</ul>' +
			'	</div>' +
			'</div>' +
			'';

		top = '' + '<section class="top">' + top + opt + '</section>' + '';

		//마감소재
		clsm = '' + '<dl class="s1-1">' + '	<dt>마감소재</dt>' + clsm + '</dl>' + '';
		clropt = '' + '<dl class="s1-1 clrOpt">' + '	<dt>옵션</dt>' + clropt + '</dl>' + '';
		org = '' + '<dl>' + '	<dt>원산지</dt>' + '	<dd>' + org + '	</dd>' + '</dl>' + '';
		mfs = '' + '<dl>' + '	<dt>제조원</dt>' + '	<dd>' + mfs + '	</dd>' + '</dl>' + '';
		if (sz == '') {
			sz = '';
		} else {
			sz =
				'' +
				'<div class="size">' +
				'	<dl>' +
				'		<dt>사이즈</dt>' +
				'		<dd>' +
				'			<img src="' +
				loc +
				onOff +
				'/' +
				outPut +
				'/' +
				sCode +
				'/' +
				nCode +
				'/v02.png" title="' +
				loc +
				onOff +
				'/' +
				outPut +
				'/' +
				sCode +
				'/' +
				nCode +
				'/v02.png" onerror="this.src=\'' +
				loc +
				'/basic/01/v02.png\'" alt="">' +
				'			<ul>' +
				sz +
				'			</ul>' +
				'		</dd>' +
				'	</dl>' +
				'</div>' +
				'';
		}

		spec = '' + cfg + std + clsm + clropt + '<div class="s2">' + org + mfs + '</div>' + cert + sz + ctn + '';
		// spec = reCode(spec);
		spec = '' + '<div class="specBox">' + '	<h2><span>스펙</span></h2>' + '	<section class="spec">' + '		<h3>제품스펙</h3>' + '		<div class="spcDetail">' + spec + '		</div>' + '	</section>' + '</div>' + '';

		//모듈합계
		function cModule(e, k) {
			//alert(e);
			Temp = '';
			var hvr = '';
			for (y = 4; y <= 23; y++) {
				if (ex[e][y] !== '') {
					//한자리수 처리
					if (ex[e][y].length === 1) {
						ex[e][y] = '0' + ex[e][y];
					} else if (ex[e][y].length === 2 && Left(ex[e][y], 1) === 'g') {
						ex[e][y] = 'g0' + Right(ex[e][y], 1);
					} else if (ex[e][y].length === 2 && Left(ex[e][y], 1) === 'r') {
						ex[e][y] = 'r0' + Right(ex[e][y], 1);
					}

					//확장자 처리
					if (Left(ex[e][y], 1) === 'g') {
						ex[e][y] = '' + '			<img src="' + loc2 + onOff + '/' + outPut + '/' + sCode + '/group/' + k + '/' + Right(ex[e][y], 2) + '.gif' + '" ' + hvr + ' style="' + imgAtr + '" alt="' + ex[e + 2][y] + '"/>' + '';
						hvr = '';
					} else if (Left(ex[e][y], 1) === 'r') {
						hvr = ' onmouseover="' + loc2 + onOff + '/' + outPut + '/' + sCode + '/group/' + k + '/' + Right(ex[e][y], 2) + '_r.jpg"' + ' onmouseout="' + loc2 + onOff + '/' + outPut + '/' + sCode + '/group/' + k + '/' + Right(ex[e][y], 2) + '.jpg"';
						ex[e][y] = '' + '			<img src="' + loc2 + onOff + '/' + outPut + '/' + sCode + '/group/' + k + '/' + Right(ex[e][y], 2) + '.jpg' + '" ' + hvr + ' style="' + imgAtr + '" alt="' + ex[e + 2][y] + '"/>' + '';
					} else if (Left(ex[e][y], 4) === 'http') {
						//alert(ex[e][y]);
						hvr = '';
						ex[e][y] =
							'' +
							//'			<img src="'+loc2+onOff+'/'+outPut+'/'+sCode+'/group/'+k+'/'+
							// 				ex[e][y]+
							// 				'.jpg" '+
							// ' 			style="'+imgAtr+'" alt="동영상"/>'+
							'<div class="mp4" style="max-width:780px;margin:0 auto;">' +
							'	<style>.embed-container iframe, .embed-container object, .embed-container embed {position:absolute;top:0;left:0;width:100%;height:100%;}</style>' +
							'	<div class="embed-container" style="position:relative;padding-bottom:56.25%;height:0;overflow:hidden;max-width:100%;">' +
							'	<iframe src="' +
							ex[e][y] +
							'" title="동영상" width="100%" height="440" frameborder="0" allowfullscreen></iframe></div>' +
							'</div>' +
							'';
						//alert(ex[e][y]);
					} else {
						ex[e][y] = '' + '			<img src="' + loc2 + onOff + '/' + outPut + '/' + sCode + '/group/' + k + '/' + ex[e][y] + '.jpg' + '" ' + hvr + ' style="' + imgAtr + '" alt="' + ex[e + 2][y] + '"/>' + '';
					}

					Temp = Temp + ex[e][y] + '';
					//alert(Temp);
				}
			}
			return Temp;
		}

		//코딩모듈영역
		for (e = line + 1; e <= oRange.Rows.count; e++) {
			//alert(ex[e][2]);
			switch (String(ex[e][2])) {
				case '01':
					cModule(e, String(ex[e][2]));
					seriesC =
						seriesC +
						'<div id="series">' +
						//'series<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('series');
					sumSethtml = sumSethtml + seriesC;
					break;
				case '02':
					//alert(ctn);
					specC =
						'' +
						//'specC'+
						'<div id="spec"><!-- <a href="#option"> --><img src="' +
						loc2 +
						onOff +
						'/' +
						outPut +
						'/' +
						sCode +
						'/' +
						nCode +
						'/01.jpg" alt="' +
						specAlt +
						'" style="' +
						imgAtr +
						'" border=0/><!-- </a> --></div>' +
						'<div id="size"><img src="' +
						loc2 +
						onOff +
						'/' +
						outPut +
						'/' +
						sCode +
						'/' +
						nCode +
						'/02.jpg" alt="" style="' +
						imgAtr +
						'"/></div>' +
						ctn +
						'';
					//sumSet.add('spec');
					sumSethtml = sumSethtml + specC;
					break;
				case '03':
					cModule(e, String(ex[e][2]));
					productC =
						productC +
						'<div id="product">' +
						//'product<br/>'+
						//'	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('product');
					sumSethtml = sumSethtml + productC;
					break;
				case '04':
					cModule(e, String(ex[e][2]));
					detailC =
						detailC +
						'<div id="detail">' +
						//'detail<br/>'+
						//'	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('detail');
					sumSethtml = sumSethtml + detailC;
					break;
				case '05':
					cModule(e, String(ex[e][2]));
					checkC =
						checkC +
						'<div id="check">' +
						//'check<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('check');
					sumSethtml = sumSethtml + checkC;
					break;
				case '06':
					cModule(e, String(ex[e][2]));
					specialC =
						specialC +
						'<div id="special">' +
						//'special<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('special');
					sumSethtml = sumSethtml + specialC;
					break;
				case '07':
					cModule(e, String(ex[e][2]));
					moduleC =
						moduleC +
						'<div id="module">' +
						//'module<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('module');
					sumSethtml = sumSethtml + moduleC;
					break;
				case '08':
					cModule(e, String(ex[e][2]));
					materialC =
						materialC +
						'<div id="material">' +
						//'material<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('material');
					sumSethtml = sumSethtml + materialC;
					break;
				case '09':
					cModule(e, String(ex[e][2]));
					manualC =
						manualC +
						'<div id="manual">' +
						//'manual<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('manual');
					sumSethtml = sumSethtml + manualC;
					break;
				case '10':
					cModule(e, String(ex[e][2]));
					MDC =
						MDC +
						'<div id="MD">' +
						//'MD<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('MD');
					sumSethtml = sumSethtml + MDC;
					break;
				case '11':
					cModule(e, String(ex[e][2]));
					guideC =
						guideC +
						'<div id="guide">' +
						//'guide<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('guide');
					sumSethtml = sumSethtml + guideC;
					break;
				case '12':
					cModule(e, String(ex[e][2]));
					upgradeC =
						upgradeC +
						'<div id="upgrade">' +
						//'upgrade<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('upgrade');
					sumSethtml = sumSethtml + upgradeC;
					break;
				case '13':
					cModule(e, String(ex[e][2]));
					behindC =
						behindC +
						'<div id="behind">' +
						//'behind<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					// sumSet.add('behind');
					sumSethtml = sumSethtml + behindC;
					break;
				case '14':
					cModule(e, String(ex[e][2]));
					sizeTC =
						sizeTC +
						'<div id="sizeT">' +
						//'sizeT<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('sizeT');
					sumSethtml = sumSethtml + sizeTC;
					break;

				case '17':
					cModule(e, String(ex[e][2]));
					mixC =
						mixC +
						'<div id="mix">' +
						//'mix<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('mix');
					sumSethtml = sumSethtml + mixC;
					break;
				case '18':
					cModule(e, String(ex[e][2]));
					introC =
						introC +
						'<div id="intro">' +
						//'intro<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('intro');
					sumSethtml = sumSethtml + introC;
					break;
				case '19':
					cModule(e, String(ex[e][2]));
					allTC =
						allTC +
						'<div id="allT">' +
						//'allT<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('allT');
					sumSethtml = sumSethtml + allTC;
					break;
				case '20':
					cModule(e, String(ex[e][2]));
					conceptC =
						conceptC +
						'<div id="concept">' +
						//'concept<br/>'+
						// '	<div class="preview">'+
						Temp +
						// '	</div>'+
						// '	<div class="capture"></div>'+
						'</div>' +
						'';
					//sumSet.add('concept');
					sumSethtml = sumSethtml + conceptC;
					break;
				// case '21':
				// 	cModule(e, String(ex[e][2]));
				// 	colorC = colorC+
				// 		'<div id="color">'+
				// 		//'color<br/>'+
				// 		// '	<div class="preview">'+
				// 			Temp+
				// 		// '	</div>'+
				// 		// '	<div class="capture"></div>'+
				// 		'</div>'+
				// 		'';
				// 		//sumSet.add('color');
				// 	sumSethtml = sumSethtml + colorC;
				// 	break;
			}
		}

		reviewC =
			reviewC +
			'<div id="review">' +
			'	<img src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/15/01.jpg" alt="리뷰" style="' +
			imgAtr +
			'"/>' +
			'	<img src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/15/02.jpg" alt="" style="' +
			imgAtr +
			'"/>' +
			'</div>' +
			'';

		optionC =
			optionC +
			'<div id="option">' +
			'	<img src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/16/01.jpg" alt="옵션" style="' +
			imgAtr +
			'"/>' +
			'	<img src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/16/02.jpg" alt="" style="' +
			imgAtr +
			'"/>' +
			'</div>' +
			'';

		//1회만 노출
		//var sumSethtml='';
		// sumSet.forEach(function(e){
		// 	switch(e){
		// 		case('series'):
		// 			sumSethtml += seriesC;
		// 			break;
		// 		case('spec'):
		// 			sumSethtml += specC;
		// 			break;
		// 		case('product'):
		// 			sumSethtml += productC;
		// 			break;
		// 		case('detail'):
		// 			sumSethtml += detailC;
		// 			break;
		// 		case('check'):
		// 			sumSethtml += checkC;
		// 			break;

		// 		case('special'):
		// 			sumSethtml += specialC;
		// 			break;
		// 		case('module'):
		// 			sumSethtml += moduleC;
		// 			break;
		// 		case('material'):
		// 			sumSethtml += materialC;
		// 			break;
		// 		case('manual'):
		// 			sumSethtml += manualC;
		// 			break;
		// 		case('MD'):
		// 			sumSethtml += MDC;
		// 			break;
		// 		case('upgrade'):
		// 			sumSethtml += upgradeC;
		// 			break;
		// 		case('behind'):
		// 			sumSethtml += behindC;
		// 			break;
		// 		case('sizeT'):
		// 			sumSethtml += sizeTC;
		// 			break;
		// 		case('review'):
		// 			sumSethtml += reviewC;
		// 			break;
		// 		case('option'):
		// 			sumSethtml += optionC;
		// 			break;

		// 		case('mix'):
		// 			sumSethtml += mixC;
		// 			break;
		// 		case('intro'):
		// 			sumSethtml += introC;
		// 			break;
		// 		case('allT'):
		// 			sumSethtml += allTC;
		// 			break;
		// 		case('concept'):
		// 			sumSethtml += conceptC;
		// 			break;
		// 		case('color'):
		// 			sumSethtml += colorC;
		// 			break;
		// 	}
		// });
		//치환
		//view = reCode(sumSethtml);

		go =
			'' +
			'<!-- 바로가기 -->' +
			'	<div class="group">' +
			'		<img style="' +
			imgAtr +
			'" alt="바로가기" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/go1.jpg"/>' +
			'		<table cellpadding=0 cellspacing=0 border=0 style="max-width:720px; margin:0 auto">' +
			'			<tr><td width="33.3%"><a href="#spec"><img style="width:100%;border:0;' +
			imgAtr +
			'" alt="스펙" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/go2.jpg" border=0/></a></td>' +
			'				<td width="33.3%"><a href="#size"><img style="width:100%;border:0;' +
			imgAtr +
			'" alt="사이즈" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/go3.jpg" border=0/></a></td>' +
			'				<td width="33.3%"><a href="#product"><img style="width:100%;border:0;' +
			imgAtr +
			'" alt="기능" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/go4.jpg" border=0/></a></td></tr>' +
			'			<tr><td><a href="#detail"><img style="width:100%;border:0;' +
			imgAtr +
			'" alt="디테일" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/go5.jpg" border=0/></a></td>' +
			'				<td><a href="#module"><img style="width:100%;border:0;' +
			imgAtr +
			'" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/go6.jpg" alt="모듈" border=0/></a></td>' +
			'				<td><a href="#ship"><img style="width:100%;border:0;' +
			imgAtr +
			'" alt="배송정보" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/go7.jpg" border=0/></a></td></tr>' +
			'		</table>' +
			'	</div>' +
			'';

		//alert(brand);
		all.innerHTML =
			'' +
			'<!-- <title>' +
			ex[13][2] +
			'</title> -->' +
			'<!-- 상품기술서 시작 -->' +
			// '<div class="preview">'+
			'	<div class="pdtDetail" style="max-width:780px;margin:0 auto;">' +
			//sum+
			top +
			spec +
			//view+
			//ship+
			'	</div>' +
			// '</div>'+
			//'<div class="capture"></div>'+
			'	<!-- // 상품기술서 끝 -->' +
			'';

		sum =
			'' +
			'	<!-- <title>' +
			ex[27][4] +
			'</title> -->' +
			'	<!-- 상품기술서 시작 -->' +
			'	<div class="pdtDetail" style="margin:0 auto;">' +
			'		<div id="ipgo"><img src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/00/01.jpg" alt="입고지연" style="' +
			imgAtr +
			'"/></div>' +
			go +
			'	<!-- Top -->' +
			'	<img style="margin:1px auto;display:block;" alt="브랜드1" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/bnd01.jpg"/>' +
			'	<img style="margin:1px auto;display:block;" alt="시리즈1" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/01/c01.jpg"/>' +
			sumSethtml +
			optionC +
			reviewC +
			'	<img style="margin:1px auto;display:block;" alt="시리즈2" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/' +
			sCode +
			'/group/01/c02.jpg"/>' +
			'	<img style="margin:1px auto;display:block;" alt="브랜드2" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/bnd02.jpg"/>' +
			'	<!-- bottom -->' +
			'	<img name="ship" id="ship" style="margin:10px auto;display:block;" alt="배송" src="' +
			loc2 +
			onOff +
			'/' +
			outPut +
			'/01_cmn/shp.jpg"/>' +
			'	<p style="display:block;overflow:hidden;height:0;">궁금한 사항은 고객센터(T.1577-3332) 문의</p>' +
			'</div>' +
			'<!-- // 상품기술서 끝 -->' +
			'';

		alert('(코딩글자수: ' + sum.length + '/12,000)');
		allC.innerHTML = reTag(sum);
		// allC.innerHTML = sum;
	}
}
