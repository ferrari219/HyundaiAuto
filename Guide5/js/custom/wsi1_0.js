//------------------------ 시작 -------------------------------//
//기본변수
var i; //펑션 숫자 변수
var k;
var brand = '';
var loc = 'http://100.10.10.200/pd/';
var loc2 = 'http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/';
var temp;
var onOff;

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
	temp = temp.replace(/\</gi, '&lt;'); //치환
	temp = temp.replace(/\>/gi, '&gt;'); //치환
	temp = temp.replace(/\.jpg\"&gt;/gi, '.jpg"/&gt;'); //치환
	return temp;
}
function reAdrs(n) {
	temp = String(n);
	temp = temp.replace(/src="http:\/\/100.10.10.200\/pd\//gi, 'src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/');
	return temp;
}

// 텍스트 변경
function txtCng(k) {
	k = String(k);
	k = k.replace(/\;/gi, '<br/>'); //치환
	k = k.replace(/\_/gi, '&nbsp;&nbsp;');
	return k;
}

//------------------------- 엑셀 영역 --------------------------------//
//엑셀 불러오기
function excel(Location, i, m) {
	//엑셀이 입력될 타겟, 기획전타입번호
	var sCode = document.frm.series.value.trim(); //document.getElementById('dev_tbl').getElementsByTagName('input')[2*(k-1)+0].value;//시리즈코드 인풋값 반영
	var nCode = '01'; //document.frm.numb.value.trim(); //document.getElementById('dev_tbl').getElementsByTagName('input')[2*(k-1)+1].value;//넷하드 상품번호 인풋값 반영
	var eCode;
	var sum = '';
	var sumS = '';
	var main = '';
	var mainC = '';
	var mainS = '';
	var mainTip = '';
	var mainExp = '';
	var info = '';
	var infoC = '';
	var color = '';
	var colorC = '';
	var colorOpt = '';
	var features = '';
	var featuresC = '';
	var featuresS = '';
	var fPicTr = '';
	var fPicTrC = '';
	var fPicTrS = '';
	var fPic = '';
	var fPicC = '';
	var fPicS = '';
	var fCnt = 1;
	var fCntS = 1;
	var use = '';
	var useC = '';
	var useS = '';
	var label = '';
	var labelC = '';

	//---------------------------------상하단 공통영역(개발중)------------------------------------/
	bnd = m; //Left(nCode, 1);
	//alert(bnd);
	switch (bnd) {
		//WSI 세부 브랜드
		case '1':
			pnt = '01_ws';
			break; //WILLIAMS SONOMA
		case '8':
			pnt = '02_we';
			break; //WESTELM
		case '2':
			pnt = '03_pb';
			break; //POTTERYBARN
		case '9':
			pnt = '04_pk';
			break; //POTTERYBARN KIDS
		//default : j=99; alert('다시 설정 해주세요.');break;
	}

	switch (i) {
		//WSI 브랜드
		case '110':
			onOff = '01_online';
			outPut = '10_wsi';
			break; //WSI
		//default : j=99; alert('다시 설정 해주세요.');break;
	}

	//상품평 이벤트
	var top = new Array();
	top[110] = '' + '<!-- 상품평 이벤트 영역(공통) -->' + '	<img style="width:100%;margin: 0px auto; display: block;" alt="temp" src="' + loc2 + onOff + '/' + outPut + '/' + pnt + '/01_common/00_top/01.jpg"/>' + '';

	var brand = new Array();
	brand[110] = '' + '<!-- WSI : 브랜드 소개 영역(공통) -->' + '<img name="brand" id="brand" style="margin: 50px auto; display: block;" alt="윌리엄소노마 브랜드" src="' + loc2 + onOff + '/' + outPut + '/' + pnt + '/01_common/99_brand/01.jpg"/>' + '';

	//배송안내 WSI
	var ship = new Array();
	ship[110] = '<!-- WSI : 배송안내 영역(공통)  -->' + '<img name="ship" id="ship" style="margin: 50px auto; display: block;" alt="윌리엄소노마 배송안내" src="' + loc2 + onOff + '/' + outPut + '/' + pnt + '/01_common/99_ship/01.jpg"/>' + '';

	//i값 - WSI: 110
	switch (i) {
		//WSI 브랜드
		case '110':
			top = top[110];
			brand = brand[110];
			ship = ship[110];
			break; //WSI
		//default : j=99; alert('다시 설정 해주세요.');break;
	}

	//---------------------------------------------------------------------------------/

	//엑셀용 변수 선언
	var Excel, oRange;

	// 엑셀 객체생성
	Excel = new ActiveXObject('Excel.Application');
	Excel.Application.Workbooks.Open('\\\\100.10.10.200\\Lguidesystem\\pd\\' + onOff + '\\' + outPut + '\\' + pnt + '\\' + sCode + '.xls');
	//Excel.Application.Workbooks.Open('C:/Guide/pd/'+onOff+'/'+outPut+'/'+sCode+'.xls');
	//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/pd/'+onOff+'/'+outPut+'/'+sCode+'.xls');

	//---------------------------------고유번호영역------------------------------------/dinnerware

	Excel.Application.Visible = true; //엑셀표시

	Excel.Application.Worksheets('00').Activate; //sheet 고유 넷하드 번호 노출
	oRange0 = Excel.Application.ActiveSheet.UsedRange; // sheet1내 작업영역을 객체에 저장

	Excel.Application.Worksheets(nCode).Activate; //sheet 고유 넷하드 번호 노출
	oRange = Excel.Application.ActiveSheet.UsedRange; // sheet1내 작업영역을 객체에 저장
	//oRange부분을 선택표시
	oRange.Select(); //oRange에는 현재 내용이 있는 부분만 선택됨.

	//---------------------------------------------------------------------------------/

	//oRange부분을 브라우저에 출력
	var all = document.getElementById('dev_all'); //프리뷰 부분 찾고
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
	var exS = new Array(); //ex값 개별 배열선언
	var idgap;
	var line;

	//로딩작업
	for (var e = 20; e <= oRange.Rows.count; e++) {
		ex[e] = new Array(); //ex[e]값 배열선언
		for (var j = 1; j <= oRange.Columns.count; j++) {
			ex[e][j] = oRange.Cells(e, j);
			//alert("ex["+e+"]["+j+"]="+ex[e][j]);
			if (ex[e][j] == '.' || ex[e][j].value == undefined || ex[e][j].value == null) {
				ex[e][j] = '';
			} else {
				// sum = sum + 'ex['+e+']['+j+'] = ' + ex[e][j] + '  ';
				// if(j==oRange.Columns.count){
				// 	sum = sum + '<br/>';
				// }
			}
		}
		//코딩 모듈영역 구하기
		if (ex[e][1] == '섹션명') {
			//39
			line = e;
			//alert(line);
		}
	}
	//입력 작업이 끝난 후 코딩 조합
	// ==== info =====
	for (var e = 30; e <= 37; e++) {
		//필요한 부분부터 재로딩
		if (ex[e][1] !== '') {
			if (ex[e][j] !== '') {
				info = info + '				<tr>' + '					<th>' + ex[e][1] + '</th>' + '					<td>' + txtCng(ex[e][3]) + '</td>' + '				</tr>' + '';
			}
			//뚜껑
		}
		colorOpt = '';
	}

	// ==== color & label =====
	for (var e = 22; e <= 27; e++) {
		//필요한 부분부터 재로딩

		if (ex[e][3] !== '') {
			//alert(ex[e][5]);
			for (var j = 6; j <= 25; j++) {
				if (ex[e][j] !== '' && ex[e][j] !== undefined) {
					//alert(ex[e][j]);
					colorOpt = colorOpt + '<span>' + ex[e][j] + '</span>' + '';
				}
			}
			if (ex[20][3] == '2' || ex[20][3] == '3') {
				//alert(ex[e][7]);
				if (ex[e][11] !== '') {
					alert('SKU레이아웃이 2나 3인 경우, 옵션 노출 갯수가 제한이 됩니다. 레이아웃 1을 사용해주세요.');
				}
			}

			//뚜껑
			colorOpt = '' + '						<div class="box">' + colorOpt + '						</div>' + '';
			//alert("ex["+e+"][5]="+ex[e][5]);
			if (e >= 23 && String(ex[e - 1][5]).indexOf('-') >= 0 && String(ex[e][5]).indexOf('-') >= 0) {
				//alert(String(ex[e-1][5]).indexOf('-'));
				//alert(String(ex[e][5]).indexOf('-'));
			} else {
				color =
					color +
					'				<li>' +
					'					<div>' +
					'						<img style="width:100%;margin: 0px auto;border:1px solid #aaa; display: block;" alt="' +
					ex[e][5] +
					'" title="' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/' +
					nCode +
					'/' +
					ex[e][5] +
					'.jpg" src="' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/' +
					nCode +
					'/' +
					ex[e][5] +
					'.jpg" onerror="this.src=\'' +
					loc +
					onOff +
					'/' +
					outPut +
					'/01_ws/02_series/basic/10000000000/0000001.jpg\'"/>' +
					'						<p class="ttl">' +
					txtCng(ex[e][3]) +
					'</p>' +
					colorOpt +
					'					</div>' +
					'				</li>' +
					'';
			} //if
		}
		colorOpt = '';
		if (ex[e][3] !== '') {
			label =
				label +
				'			<li><img style="margin: 0px auto; display: block;border:1px solid #aaa" alt="' +
				ex[e][5] +
				'_label" title="' +
				loc +
				onOff +
				'/' +
				outPut +
				'/' +
				pnt +
				'/02_series/' +
				sCode +
				'/' +
				nCode +
				'/' +
				ex[e][5] +
				'_label.jpg" src="' +
				loc +
				onOff +
				'/' +
				outPut +
				'/' +
				pnt +
				'/02_series/' +
				sCode +
				'/' +
				nCode +
				'/' +
				ex[e][5] +
				'_label.jpg" onerror="this.src=\'' +
				loc +
				onOff +
				'/' +
				outPut +
				'/01_ws/02_series/basic/10000000000/0000001_label.jpg\'"/></li>' +
				'';
		}
	}
	//뚜껑

	for (var e = line + 1; e <= oRange.Rows.count; e++) {
		//필요한 부분부터 재로딩
		//alert(ex[e][2]=='01');
		//공통영역 관리
		switch (String(ex[e][2])) {
			case '01': //01_main
				if (ex[e][5] !== '') {
					mainTip = '		<p class="tip">' + txtCng(ex[e][5]) + '</p>';
				}
				if (ex[e][6] !== '') {
					mainExp = '		<p class="exp">' + txtCng(ex[e][6]) + '</p>';
				}
				main =
					main +
					'		<img style="width:100%;margin: 0px auto; display: block;" alt="시리즈" title="' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/01_main/' +
					ex[e][3] +
					'.jpg" src="' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/01_main/' +
					ex[e][3] +
					'.jpg" onerror="this.src=\'' +
					loc +
					onOff +
					'/' +
					outPut +
					'/01_ws/02_series/basic/01_group/01_main/v01.jpg\'"/>' +
					mainTip +
					mainExp +
					'';
				mainC =
					mainC +
					'<img style="margin: 0px auto; display: block;" alt="메인" src="' +
					loc2 +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/01_main/' +
					Right(ex[e][3], 2) +
					'.jpg" title="' +
					loc2 +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/01_main/01.jpg" onerror="this.src=\'' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/01_main/' +
					Right(ex[e][3], 2) +
					'.jpg\'"/>' +
					'';
				break;
			case '04': //04_features
				if (Left(ex[e][3], 1) == 'v') {
					fPicTr =
						fPicTr +
						'					<td width=49%>' +
						'						<img style="width:100%;margin: 0px auto; display: block;" alt="" title="' +
						loc +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						ex[e][3] +
						'.jpg" src="' +
						loc +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						ex[e][3] +
						'.jpg" onerror="this.src=\'' +
						loc +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						ex[e][3] +
						'.jpg\'"/>' +
						'					</td>' +
						'';
					//alert(fCnt);
					if (fCnt % 2 == 0) {
						fPicTr = '' + '				<tr>' + fPicTr + '				</tr>' + '				<tr>' + '					<td colspan=3 height=10></td>' + '				</tr>' + '';
						fPic = fPic + fPicTr;
						fPicTr = '';
					} else {
						fPicTr = '' + fPicTr + '					<td width=2%></td>' + '';
					}
					//alert(fPicTrC);
					fCnt++;
				} else {
					features = features + '			<li>' + txtCng(ex[e][6]) + '</li>' + '';
					//alert(features);
					featuresC =
						featuresC +
						'		<img style="margin: 0px auto; display: block;" alt="" src="' +
						loc2 +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						ex[e][3] +
						'.jpg" title="' +
						loc2 +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						ex[e][3] +
						'.jpg" onerror="this.src=\'' +
						loc +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						ex[e][3] +
						'.jpg\'"/>' +
						'';
				}
				break;
			case '05': //05_use
				use = use + '			<li>' + txtCng(ex[e][6]) + '</li>' + '';
				useC =
					useC +
					'	<img style="margin: 0px auto; display: block;" alt="" src="' +
					loc2 +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/05_use/' +
					ex[e][3] +
					'.jpg" title="' +
					loc2 +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/05_use/' +
					ex[e][3] +
					'.jpg" onerror="this.src=\'' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/05_use/' +
					ex[e][3] +
					'.jpg\'"/>' +
					'';
				break;
			default:
				break;
		}
	} //for

	//뚜껑
	main = '' + '	<div class="main">' + main + '	</div>' + '';
	mainC =
		mainC +
		'<img style="margin: 0px auto; display: block;" alt="Temp" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/01_group/01_main/c01.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/01_ws/02_series/basic/01_group/01_main/c01.jpg\'"/>' +
		'';
	//홀수값 산정
	if (fCnt % 2 == 0) {
		fPicTr = '' + '				<tr>' + fPicTr + '					<td width=49%></td>' + '				</tr>' + '				<tr>' + '					<td colspan=3 height=10></td>' + '				</tr>' + '';
	}

	fPic = '' + '		<table style="width:100%;max-width:760px;">' + '			<tbody>' + fPic + fPicTr + '			</tbody>' + '		</table>' + '';
	fPicC = reAdrs(fPic);

	info = '' + '	<div class="info">' + '		<h2>기본정보</h2>' + '		<table border=0 cellspacing=0>' + '			<tbody>' + info + '			</tbody>' + '		</table>' + '';
	infoC =
		'' +
		'<img style="margin: 0px auto; display: block;" alt="기본정보" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/01.jpg" title="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/01.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/01.jpg\'"/>' +
		'';
	features = '' + '	<div class="features">' + '		<h2>주요특징</h2>' + fPic + '		<ul class="dot">' + features + '		</ul>' + '	</div>' + '';
	featuresC =
		'' +
		'	<div style="padding-top:20px;">' +
		featuresC +
		'	<img style="margin: 0px auto; display: block;" alt="공통" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/01_group/04_features/c01.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/01_ws/02_series/basic/01_group/04_features/c01.jpg\'"/>' +
		'	<img style="margin: 0px auto; display: block;" alt="Temp" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/c04.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/01_ws/02_series/basic/10000000000/c04.jpg\'" />' +
		'	</div>' +
		'';
	use = '' + '	<div class="use">' + '		<h2>사용&amp;관리</h2>' + '		<ul class="dot">' + use + '		</ul>' + '	</div>' + '';
	useC =
		'' +
		'<div>' +
		'	<img style="margin: 0px auto; display: block;" alt="사용&amp;관리" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/01_common/99_etc/05.gif"/>' +
		useC +
		'	<img style="margin: 0px auto; display: block;" alt="공통" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/01_group/05_use/c01.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/01_ws/02_series/basic/01_group/05_use/c01.jpg\'"/>' +
		'	<img style="margin: 0px auto; display: block;" alt="Temp" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/c05.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/01_ws/02_series/basic/10000000000/c05.jpg\'"/>' +
		'</div>' +
		'';
	if (color !== '') {
		color = '' + '		<div class="color">' + '			<h2>상품정보</h2>' + '			<p class="tip">* 모니터 해상도에 따라 색상차이가 있을 수 있습니다.</p>' + '			<ul class="' + ex[20][4] + '">' + color + '			</ul>' + '		</div>' + '	</div>' + '';
	}
	colorC =
		'' +
		'<img style="margin: 0px auto; display: block;" alt="옵션" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/02.jpg" title="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/02.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/02.jpg\'"/>' +
		'';
	if (label !== '') {
		label = '' + '	<div class="label">' + '		<ul class="' + ex[20][4] + '">' + label + '		</ul>' + '	</div>' + '';
	}
	labelC =
		'' +
		'<img style="margin: 0px auto; display: block;" alt="label" src="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/03.jpg" title="' +
		loc2 +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/03.jpg" onerror="this.src=\'' +
		loc +
		onOff +
		'/' +
		outPut +
		'/' +
		pnt +
		'/02_series/' +
		sCode +
		'/' +
		nCode +
		'/03.jpg\'"/>' +
		'';
	// go=''+
	// 	'';

	all.innerHTML =
		'' +
		'<!-- <title>' +
		nCode +
		'</title> -->' +
		'<!-- 상품기술서 시작 입니다. -->' +
		'<div class="wsi_detail" style="max-width:760px;margin:0 auto;">' +
		top +
		main +
		info +
		color +
		features +
		use +
		label +
		brand +
		ship +
		'</div>' +
		'<!-- // 상품기술서 끝 입니다. -->' +
		'';

	sum =
		'' +
		//'<!-- <title>'+nCode+'</title> -->'+
		//'<!-- 상품기술서 시작 입니다. -->'+
		//'<div class="wsi_detail" style="max-width:760px;margin:0 auto;background:#fff">'+
		//mainC+
		//infoC+
		//colorC+
		//'<div>'+
		//'	<img style="margin: 0px auto; display: block;" alt="주요특징" src="'+loc2+onOff+'/'+outPut+'/'+pnt+'/01_common/99_etc/04.gif"/>'+
		//fPicC+
		//featuresC+
		//'</div>'+
		//useC+
		//brand+
		//labelC+
		//ship+
		//'</div>'+
		//'<!-- // 상품기술서 끝 입니다. -->'+
		'';
	allC.innerHTML = '';
	//allC.innerHTML = reCode(sum);
	//allC.innerHTML = sum;

	//00페이지 로딩작업
	for (var e = 21; e <= oRange0.Rows.count; e++) {
		exS[e] = new Array(); //ex[e]값 배열선언
		for (var j = 1; j <= oRange0.Columns.count; j++) {
			exS[e][j] = oRange0.Cells(e, j);
			//alert(ex[e][j])
			if (exS[e][j] == '.' || exS[e][j].value == undefined || exS[e][j].value == null || exS[e][j].value == false) {
				exS[e][j] = '';
			} else {
				sumS = sumS + 'exS[' + e + '][' + j + '] = ' + exS[e][j] + '  ';
				if (j == oRange0.Columns.count) {
					sumS = sumS + '<br/>';
				}
			}
		}
	}

	//입력 작업이 끝난 후 코딩 조합
	var fPicArrS = new Array();

	for (e = 21; e <= oRange0.Rows.count; e++) {
		switch (String(exS[e][2])) {
			case '01': //01_main
				if (exS[e][7] !== '') {
					mainTipS = '		<p class="tip">' + exS[e][7] + '</p>';
				}
				if (exS[e][9] !== '') {
					mainExpS = '		<p class="exp">' + exS[e][9] + '</p>';
				}
				mainS =
					mainS +
					'		<img style="width:100%;margin: 0px auto; display: block;" alt="' +
					exS[e][3] +
					'" title="' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/01_main/' +
					exS[e][3] +
					'.jpg" src="' +
					loc +
					onOff +
					'/' +
					outPut +
					'/' +
					pnt +
					'/02_series/' +
					sCode +
					'/01_group/01_main/' +
					exS[e][3] +
					'.jpg" onerror="this.src=\'' +
					loc +
					onOff +
					'/' +
					outPut +
					'/01_ws/02_series/basic/01_group/01_main/v01.jpg\'"/>' +
					mainTipS +
					mainExpS +
					'';
				break;
			case '04': //04_features
				if (Left(exS[e][3], 1) == 'v') {
					fPicTrS =
						fPicTrS +
						'					<td width=49%>' +
						'						<img style="width:100%;margin: 0px auto; display: block;" alt="' +
						exS[e][3] +
						'" title="' +
						loc +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						exS[e][3] +
						'.jpg" src="' +
						loc +
						onOff +
						'/' +
						outPut +
						'/' +
						pnt +
						'/02_series/' +
						sCode +
						'/01_group/04_features/' +
						exS[e][3] +
						'.jpg" onerror="this.src=\'' +
						loc +
						onOff +
						'/' +
						outPut +
						'/01_ws/02_series/basic/01_group/04_features/' +
						exS[e][3] +
						'.jpg\'"/>' +
						'					</td>' +
						'';
					if (fCntS % 2 == 0) {
						fPicTrS = '' + '				<tr>' + fPicTrS + '					<td width=2%></td>' + '				</tr>' + '';

						fPicS = fPicS + fPicTrS;
						fPicTrS = '';
					} else {
						fPicTrS = '' + fPicTrS + '';
						// fPicTrC = fPicTrC+
						// 		Right(ex[e][3],2)+
						// 		'';
					}
					//alert(fPicTrC);
					fCntS++;
				} else {
					featuresS =
						featuresS +
						//'			<li>'+exS[e][9]+'</li>'+
						'';
				}
				break;
			case '05': //05_use
				useS =
					useS +
					//'			<li>'+exS[e][9]+'</li>'+
					'';
				break;
		}
	}

	//fPicS

	//뚜껑
	mainS = '' + '	<div class="main">' + mainS + '	</div>' + '';
	if (fPicS !== '') {
		fPicS = '' + '		<table style="width:100%;max-width:760px;">' + '			<tbody>' + fPicS + '			</tbody>' + '		</table>' + '';
	}

	featuresS = '' + '	<div class="features">' + fPicS + '		<ul class="dot">' + featuresS + '		</ul>' + '	</div>' + '';
	useS = '' + '	<div class="use">' + '		<ul class="dot">' + useS + '		</ul>' + '	</div>' + '';

	allS.innerHTML =
		'' +
		'<div class="wsi_detail" style="max-width:760px;margin:0 auto;">' +
		//'	<h4>01_main.PSD</h4>'+
		//mainS+
		//'	<h4>04_features.PSD</h4>'+
		//featuresS+
		//'	<h4>05_use.PSD</h4>'+
		//useS+
		//sumS+
		'</div>';
	('');

	//allC.innerHTML=reCode(allC.innerHTML);
}
