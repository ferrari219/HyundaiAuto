//------------------------ 시작 -------------------------------//
//기본변수
var i;//펑션 숫자 변수
var k;
var brand='';
var loc = 'http://172.17.14.220/04_ban/',
	loc2 = 'http://mall.hyundailivart.co.kr/UserFiles/Image/04_ban/';
var temp, tCnt, Temp, TempCol, TempRow;
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
//------------------------- 공통 영역 --------------------------------//


//------------------------- 기타 영역 --------------------------------//

// 코딩변경
function reCode(n){
	temp = String(n);
	//temp = temp.replace(/\</gi, "&lt;"); //치환	
	//temp = temp.replace(/\>/gi, "&gt;"); //치환
	temp = temp.replace(/\;/gi, " <br/>"); //치환
	// temp = temp.replace(/\[/gi, "<span class='point1'>"); //치환	
	// temp = temp.replace(/\]/gi, "</span>"); //치환	
	temp = temp.replace(/\[{/gi, "<span class='point3'>"); //치환	
	temp = temp.replace(/\}]/gi, "</span>"); //치환	
	temp = temp.replace(/\[/gi, "<span class='point1'>"); //치환	
	temp = temp.replace(/\]/gi, "</span>"); //치환	
	temp = temp.replace(/\{/gi, "<span class='point2'>"); //치환	
	temp = temp.replace(/\}/gi, "</span>"); //치환	
	n = temp;
	//temp = temp.replace(/\.jpg\"&gt;/gi, '.jpg\"/&gt;'); //치환	
	return n;
}
function reTag(n){
	temp = String(n);
	temp = temp.replace(/\</gi, "&lt;"); //치환	
	temp = temp.replace(/\>/gi, "&gt;"); //치환
	n = temp;
	return n;
}
function reAdrs(n){
	temp = String(n);
	temp = temp.replace(/\/\/172.17.14.220\/Lguidesystem\/pd\//gi,"http://mall.hyundailivart.co.kr/UserFiles/Image/00_pd/")
	return temp;
}
// 텍스트 변경
function txtCng(k){
	k = String(k);
	k = k.replace(/\_/gi, '&nbsp;&nbsp;'); 
	k = k.replace(/\;/gi, '<br/>'); //치환	
	return k;
}
//var test = Address('//172.17.14.220/Lguidesystem/00_pd/'); 
//alert(test)

//------------------------- 엑셀 영역 --------------------------------//
//엑셀 불러오기
function excel(Location, i){ //엑셀이 입력될 타겟, 기획전타입번호
	var sCode = document.frm.series.value; //document.getElementById('dev_tbl').getElementsByTagName('input')[2*(k-1)+0].value;//시리즈코드 인풋값 반영
	var sum = '';
	
	var imgAtr = 'display:block; margin:0px auto;';
	
	if(sCode === "basic"){
		alert("basic은 사용할 수 없습니다. 다른 브랜드명을 사용해주세요.");
	}
	//---------------------------------상하단 공통영역(개발중)------------------------------------/
	function etc1(k){
		if (Left(k,1)=='1'){
			return onOff='on';
		}
		else{
			return onOff='off';
		};
	}
	etc1(i);
	
	switch (i){
		//온라인 브랜드
		case '101': outPut='01';break; //이즈마인
		
		default : j=99; alert('다시 설정 해주세요.');break;
	};

	//---------------------------------------------------------------------------------/

	//엑셀용 변수 선언
	var Excel,oRange;

    // 엑셀 객체생성
	Excel = new ActiveXObject("Excel.Application");
	Excel.Application.Workbooks.Open('\\\\172.17.14.220\\Lguidesystem\\04_ban\\'+sCode+'.xls');
	//Excel.Application.Workbooks.Open('D:/Guide/00_pd/'+onOff+'/'+outPut+'/'+sCode+'.xls');
	// Excel.Application.Workbooks.Open('D:/'+sCode+'.xls');
	//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/00_pd/'+onOff+'/'+outPut+'/'+sCode+'.xls'); 

	//---------------------------------고유번호영역------------------------------------/

	Excel.Application.Visible = true;	//엑셀표시

	Excel.Application.Worksheets('00').Activate;	//sheet 고유 넷하드 번호 노출
	oRange = Excel.Application.ActiveSheet.UsedRange;	// sheet1내 작업영역을 객체에 저장
	//oRange부분을 선택표시
	oRange.Select();	//oRange에는 현재 내용이 있는 부분만 선택됨.
	
	//---------------------------------------------------------------------------------/

	//oRange부분을 브라우저에 출력
	var all = document.getElementById('dev_all'); //프리뷰 부분 찾고
	all.innerHTML = '';	//초기화
	
    //---------------------------------------------------------------------------------/

	var ex = new Array(); //ex값 개별 배열선언
	
	//Slice용 이미지 출력시에만
		//로딩작업
		for(var e=25;e<=oRange.Rows.count;e++){ 
			ex[e] = new Array(); //ex[e]값 배열선언
			for(var j=18;j<=oRange.Columns.count;j++){ 
				ex[e][j] = oRange.Cells(e,j);
				//alert(ex[e][j])
				if(ex[e][j]=='.' || ex[e][j].value == undefined || ex[e][j].value == null || ex[e][j].value == false){ex[e][j]=''}
				else{
					ex[e][j] = reCode(ex[e][j]);
				}
			}
		};
		
		//정보 분류 작업
		var banMaker = ''; //배너 만들어 담을 그릇
		var lft, tp, position, lOpt, tOpt, wdh, zidx;
		var list = [
			'type', //0
			'bg', //1
			'LT', //2
			'CT', //3
			'RT', //4
			'LM', //5
			'CM', //6
			'RM', //7
			'LB', //8
			'CB', //9
			'RB', //10
		];		  
		for(i=2;i<=10;i++){
			list[i]= [
				'LR', //0
				'TB', //1
				'wide', //2
				'depth', //3
				'h3', //4
				'dt', //5
				'dd', //6
				'roundBtn', //7
				'squareBtn', //8
				'alert', //9
				'obj', //10
			];
			//['','','','','','',''];
		};
				
		//alert(list[2][0]);
		// var Applylist = {
		// 	A1: { w: 1140, h: 110 },
		// 	A2: { w: 140, h: 90 },
		// 	A3_C1: { w: 1140, h: 478 },
		// 	A4_C2: { w: 640, h: 460 },
		// 	A5: { w: 1140, h: 250 },
		// 	A6_C4: { w: 640, h: 190 },
		// 	A7_C5: { w: 1140, h: 150 },
		// 	A8: { w: 390, h: 600 },
		// 	B3: { w: 222, h: 222 },
		// 	//B4: { w: 0, h: 0 },
		// 	C3: { w: 910, h: 354 },
		// 	C6: { w: 344, h: 243 },
		// 	//C7: { w: 0, h: 0 },
		// }
		
		function lftTp(l,t,lv,tv){
			switch (l) {
				case 1:
					lft = "L"; 
					lOpt = 'left:'+lv+'%;';
					break;
					case 3:
					lft = "C";
					lOpt = 'left:'+lv+'%;';
					break;
					case 5:
					lft = "R";
					lOpt = 'right:'+lv+'%;';
					break;
				default:
					break;
			}
			switch (t) {
				case 0:
					tp = "T";
					tOpt = 'top:'+tv+'%;';
					break;
					case 8:
					tp = "M";
					tOpt = 'top:'+tv+'%;';
					break;
					case 16:
					tp = "B";
					tOpt = 'bottom:'+tv+'%;';
					break;
				default:
					break;
			}
			position = lft+tp;
			return position, lOpt, tOpt;
		}

		function handleMaker (t,e,k){
			// if(ex[e+7][k+2]==''){ex[e+7][k+2]==''}
			// else{ex[e+7][k+2]=='h3';}
			Temp = '';temp = '';

			for (n=0;n<=16;n+=8) { //8개 단위로 증가
				for (m=1;m<=5;m+=2) {
					for (i=4;i<=11;i++) {
						//alert(e+8+i);
						if(ex[e+i+n][k+m+1]==''){}
						else{
							//Temp += ex[e+i+n][k+m]+' : '+ex[e+i+n][k+m+1]+'<br/>';
							switch (ex[e+i+n][k+m]) {
								case '좌/우여백(%)':
									lOpt = ex[e+i+n][k+m+1];
									break;
								case '상/하여백(%)':
									tOpt = ex[e+i+n][k+m+1];
									break;
								case '너비':
									//wdh = 
									break;
								case '깊이':									
									zidx = 'z-index:'+ex[e+i+n][k+m+1]+';';
									//alert('ex['+Number(e+i+n)+']['+Number(k+m+1)+']='+ex[e+i+n][k+m+1]);
									break;
								case '타이틀(소)':
									//list[2][4] = '<h3>'+ex[e+i+n][k+m+1]+'</h3>';
									temp += '<h3>'+ex[e+i+n][k+m+1]+'</h3>';
									break;
								case '타이틀(대)':
									if(Left(ex[e+i+n][k+m+1],4)=='http'){ temp += '<dl><dt><img src="'+ex[e+i+n][k+m+1].trim()+'" alt="오브젝트" /></dt></dl>';}
									else{ temp += '<dl><dt>'+ex[e+i+n][k+m+1]+' </dt></dl>'; }
									break;
								case '설명':
									temp += '<dl><dd><span>'+ex[e+i+n][k+m+1]+' </dd></span></dl>';
									break;
								case '쿠폰':
									temp += '<dl class="coupon"><dd><img src="'+ex[e+i+n][k+m+1]+'" alt="쿠폰" /></dd></dl>';
									break;
								case '오브젝트':
									temp += '<p class="obj"><img src="'+ex[e+i+n][k+m+1]+'" alt="오브젝트" /></p>';
									break;
								default:
									break;
							}
						}
					}
					lftTp(m, n, lOpt, tOpt);
					//alert(position);

					Temp = Temp+
							'<div class="'+position+'" style="'+lOpt+tOpt+zidx+'">'+
							temp+
							'</div>'
							// '<br/>'+
							'';
					temp = '';
					lOpt=''; tOpt='';
				}
			};
			
			if(Left(ex[e+2][k],4)=='http'){
				// alert("http");
				list[1] = 'background:url('+ex[e+2][k]+') no-repeat center center';
			}
			else{
				list[1] = 'background:'+ex[e+2][k];
			}
			//list[1] =  //BG

			Temp = ''+
					'<div class="sec '+t+'">'+
					'	<h2>'+t+'배너</h2>'+
					'	<div class="frame" style="'+list[1]+'">'+
					Temp+
					'	</div>'+
					'</div>'+
					'';

			// Temp = ''+
				// '<div class="sec '+t+'">'+
				// '	<h3>'+t+'배너</h3>'+
				// '	<div class="frame" style="background:#b8af9e">'+
				// //url(http://172.17.14.220/04_ban/test/bg1.jpg) no-repeat center center;
				// '		<div class="CT" style="left:50%;top:0%;">'+
				// //			list.LT.Content4+
				// '			<dl>'+
				// '				<dt>딱 3일만 주는 혜택</dt>'+
				// '				<dd><span>설명글이 주저리주러리</span></dd>'+
				// '				<dd class="coupon"><img src="http://172.17.14.220/04_ban/test/coupon2.png" alt="쿠폰" /></dd>'+
				// '			</dl>'+
				// '			<p class="round btn">버튼</p>'+
				// '		</div>'+
				// '		<div class="RB" style="right:10%;bottom:0%;">'+
				// '			<p class="obj"><img src="http://172.17.14.220/04_ban/test/obj2.png" alt="오브젝트" /></p>'+
				// '		</div>'+
				// '	</div>'+
				// '</div>'+

				// '적용유형:'+t+'<br/>'+
				// 'BG: <img src="'+ex[e+2][k]+'"/><br/>'+
				// 'LT: 1.'+ex[e+4][k+1]+' '+ex[e+4][k+2]+' '+
				// 	'2.'+ex[e+5][k+1]+' '+ex[e+5][k+2]+' '+
				// 	'3.'+ex[e+6][k+1]+' '+ex[e+6][k+2]+' '+
				// 	'4.'+ex[e+7][k+1]+' '+ex[e+7][k+2]+' '+
				// 	'5.'+ex[e+8][k+1]+' '+ex[e+8][k+2]+' '+
				// 	'6.'+ex[e+9][k+1]+' '+ex[e+9][k+2]+' '+
				// 	'7.'+ex[e+10][k+1]+' '+ex[e+10][k+2]+' '+
				// 	'8.'+ex[e+11][k+1]+' '+ex[e+11][k+2]+'<br/>'+
				// 'CT: '+ex[e+4][k+3]+' '+ex[e+4][k+4]+' '+
				// 	ex[e+5][k+3]+' '+ex[e+5][k+4]+' '+
				// 	ex[e+6][k+3]+' '+ex[e+6][k+4]+' '+
				// 	ex[e+7][k+3]+' '+ex[e+7][k+4]+' '+
				// 	ex[e+8][k+3]+' '+ex[e+8][k+4]+' '+
				// 	ex[e+9][k+3]+' '+ex[e+9][k+4]+' '+
				// 	ex[e+10][k+3]+' '+ex[e+10][k+4]+' '+
				// 	ex[e+11][k+3]+' '+ex[e+11][k+4]+'<br/>'+
				// 'RT: '+ex[e+4][k+5]+' '+ex[e+4][k+6]+' '+
				// 	ex[e+5][k+5]+' '+ex[e+5][k+6]+' '+
				// 	ex[e+6][k+5]+' '+ex[e+6][k+6]+' '+
				// 	ex[e+7][k+5]+' '+ex[e+7][k+6]+' '+
				// 	ex[e+8][k+5]+' '+ex[e+8][k+6]+' '+
				// 	ex[e+9][k+5]+' '+ex[e+9][k+6]+' '+
				// 	ex[e+10][k+5]+' '+ex[e+10][k+6]+' '+
				// 	ex[e+11][k+5]+' '+ex[e+11][k+6]+'<br/>'+

				// 'LM: '+ex[e+12][k+1]+' '+ex[e+12][k+2]+' '+
				// 	ex[e+13][k+1]+' '+ex[e+13][k+2]+' '+
				// 	ex[e+14][k+1]+' '+ex[e+14][k+2]+' '+
				// 	ex[e+15][k+1]+' '+ex[e+15][k+2]+' '+
				// 	ex[e+16][k+1]+' '+ex[e+16][k+2]+' '+
				// 	ex[e+17][k+1]+' '+ex[e+17][k+2]+' '+
				// 	ex[e+18][k+1]+' '+ex[e+18][k+2]+' '+
				// 	ex[e+19][k+1]+' '+ex[e+19][k+2]+'<br/>'+
				// 'CM: '+ex[e+12][k+3]+' '+ex[e+12][k+4]+' '+
				// 	ex[e+13][k+3]+' '+ex[e+13][k+4]+' '+
				// 	ex[e+14][k+3]+' '+ex[e+14][k+4]+' '+
				// 	ex[e+15][k+3]+' '+ex[e+15][k+4]+' '+
				// 	ex[e+16][k+3]+' '+ex[e+16][k+4]+' '+
				// 	ex[e+17][k+3]+' '+ex[e+17][k+4]+' '+
				// 	ex[e+18][k+3]+' '+ex[e+18][k+4]+' '+
				// 	ex[e+19][k+3]+' '+ex[e+19][k+4]+'<br/>'+
				// 'RM: '+ex[e+12][k+5]+' '+ex[e+12][k+6]+' '+
				// 	ex[e+13][k+5]+' '+ex[e+13][k+6]+' '+
				// 	ex[e+14][k+5]+' '+ex[e+14][k+6]+' '+
				// 	ex[e+15][k+5]+' '+ex[e+15][k+6]+' '+
				// 	ex[e+16][k+5]+' '+ex[e+16][k+6]+' '+
				// 	ex[e+17][k+5]+' '+ex[e+17][k+6]+' '+
				// 	ex[e+18][k+5]+' '+ex[e+18][k+6]+' '+
				// 	ex[e+19][k+5]+' '+ex[e+19][k+6]+'<br/>'+

				// 'LB: '+ex[e+20][k+1]+' '+ex[e+20][k+2]+' '+
				// 	ex[e+21][k+1]+' '+ex[e+21][k+2]+' '+
				// 	ex[e+22][k+1]+' '+ex[e+22][k+2]+' '+
				// 	ex[e+23][k+1]+' '+ex[e+23][k+2]+' '+
				// 	ex[e+24][k+1]+' '+ex[e+24][k+2]+' '+
				// 	ex[e+25][k+1]+' '+ex[e+25][k+2]+' '+
				// 	ex[e+26][k+1]+' '+ex[e+26][k+2]+' '+
				// 	ex[e+27][k+1]+' '+ex[e+27][k+2]+'<br/>'+
				// 'CB: '+ex[e+20][k+3]+' '+ex[e+20][k+4]+' '+
				// 	ex[e+21][k+3]+' '+ex[e+21][k+4]+' '+
				// 	ex[e+22][k+3]+' '+ex[e+22][k+4]+' '+
				// 	ex[e+23][k+3]+' '+ex[e+23][k+4]+' '+
				// 	ex[e+24][k+3]+' '+ex[e+24][k+4]+' '+
				// 	ex[e+25][k+3]+' '+ex[e+25][k+4]+' '+
				// 	ex[e+26][k+3]+' '+ex[e+26][k+4]+' '+
				// 	ex[e+27][k+3]+' '+ex[e+27][k+4]+'<br/>'+
				// 'RB: '+ex[e+20][k+5]+' '+ex[e+20][k+6]+' '+
				// 	ex[e+21][k+5]+' '+ex[e+21][k+6]+' '+
				// 	ex[e+22][k+5]+' '+ex[e+22][k+6]+' '+
				// 	ex[e+23][k+5]+' '+ex[e+23][k+6]+' '+
				// 	ex[e+24][k+5]+' '+ex[e+24][k+6]+' '+
				// 	ex[e+25][k+5]+' '+ex[e+25][k+6]+' '+
				// 	ex[e+26][k+5]+' '+ex[e+26][k+6]+' '+
				// 	ex[e+27][k+5]+' '+ex[e+27][k+6]+'<br/>'+'<br/>'+
				// 	'';
			return Temp;
		}

		for(e=25;e<=oRange.Rows.count;e++){
			for (j=18;j<=oRange.Columns.count;j++) {
				if(ex[e][j]!==''){
					//alert('ex['+e+']['+j+']] = ' + ex[e][j]);
					//alert(e);
					//var k = 18;
					switch (ex[e][j]) {
						case 'A1':
							handleMaker('A1',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'A2':
							handleMaker('A2',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'A3_C1':
							handleMaker('A3_C1',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'A4_C2':
							handleMaker('A4_C2',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'A5':
							handleMaker('A5',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'A6_C4':
							handleMaker('A6_C4',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'A7_C5':
							handleMaker('A7_C5',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'A8':
							handleMaker('A8',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'B3':
							handleMaker('B3',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'B4':
							handleMaker('B4',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'C3':
							handleMaker('C3',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'C6':
							handleMaker('C6',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						case 'C7':
							handleMaker('C7',e,18);
							banMaker = banMaker+
									   		Temp+
									   		'';		
							break;
						default:
							break;
					}
				};
			
				sum = sum +
					'ex['+e+']['+j+']=' + ex[e][j]+ ' / '+
					'';				
			};
			e=e+28;
		}	
		
		//출력
		all.innerHTML = ''+
						//'<!-- <title>'+ex[13][2]+'</title> -->'+
						'<!-- 상품기술서 시작 입니다. -->'+
						'	<div class="pdtDetail" style="">'+
								banMaker+
						//		sum+
						//'		test'+
						'	</div>'+
						'<!-- // 상품기술서 끝 입니다. -->'+
						'';
	
	
}
