//------------------------ 시작 -------------------------------//
//기본변수
var i;//펑션 숫자 변수
var temp;var txt;
var wSize;
var mall;//몰변수
var scp;//CSS & font값 변수

//공통함수
function round(num,ja){
	ja=Math.pow(10,ja);
	return Math.round(num*ja)/ja;
}
function ceil(num,ja){
	ja=Math.pow(10,ja);
	return Math.ceil(num*ja)/ja;
}
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

function Mid(Str, Num1, Num2){
	//alert(Num1);
	return String(Str).substr(Num1, Num2);
}

function addComma(n) {
    if(isNaN(n)){return 0;}
     var reg = /(^[+-]?\d+)(\d{3})/;   
     n += '';
     while (reg.test(n))
       n = n.replace(reg, '$1' + ',' + '$2');
     return n;

}

function cntCng(k){
	if(k<10){k = '0'+k;}
	else{k = k;}
	return k;
}

String.prototype.trim = function(){
    return this.replace(/\s/g, "");
}


//오늘날짜 자동입력
window.onload=function(){
	now=new Date();
	var year = Right(now.getYear(),2);
	var mon = now.getMonth()+1;
	var day = now.getDate();
	
	//alert(monC);
	if(String(mon).length==1){
		mon="0"+mon;	}
	if(String(day).length==1){
		day="0"+day;	}

	//alert(year+mon+day);
	//return year+mon+day;
	document.frm.today.value = '161231_340_basic';
	//document.frm.today.value = document.frm.today.value;
	//document.frm.today.value = year+mon+day+'_딜영문명';
	dayCode=document.frm.today.value;
}

/*HTML 복사*/
function jsCopy(obj) {
 
 var copyObj = obj;
 var copyStr = document.getElementById(copyObj).innerHTML;
 
 window.clipboardData.setData("text", copyStr);
 alert("Html 소스가 복사되었습니다.");
}

// 이미지 변환 
function picResize(secPic){
	/* alert(secPic);*/
	temp = String(secPic);
	temp = temp.replace("img_95", "img_615"); //치환	
	temp = temp.replace("img_140", "img_615");
	temp = temp.replace("img_150", "img_615");
	temp = temp.replace("img_170", "img_615");
	temp = temp.replace("img_220", "img_615");
	temp = temp.replace("_detail1_5", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_5", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_5", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_5", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_5", "_detail5_ORIGIN"); 

	temp = temp.replace("_detail1_4", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_4", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_4", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_4", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_4", "_detail5_ORIGIN"); 

	temp = temp.replace("_detail1_3", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_3", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_3", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_3", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_3", "_detail5_ORIGIN");
	
	temp = temp.replace("_detail1_2", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_2", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_2", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_2", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_2", "_detail5_ORIGIN");

	temp = temp.replace("_detail1_1", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_1", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_1", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_1", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_1", "_detail5_ORIGIN");

	temp = temp.replace("THUMB_1", "ORIGIN"); 
	temp = temp.replace("THUMB_2", "ORIGIN"); 
	temp = temp.replace("THUMB_3", "ORIGIN"); 
	temp = temp.replace("THUMB_4", "ORIGIN"); 
	temp = temp.replace("THUMB_5", "ORIGIN");
	
	return temp;
}

// 텍스트 변경
function txtCng(k){
	/* alert(secPic);*/
	cng = String(k);
	//alert(cng);
	cng = cng.replace(/;/gi, "<br/>");cng = cng.replace(/_/gi, "&nbsp;&nbsp;&nbsp;&nbsp;");
	cng = cng.replace("[", "<span>");cng = cng.replace("]", "</span>");
	//cng = ""+cng;
	
	return cng;	
}

//------------------------- 공통 영역 --------------------------------//
//오픈 날짜 반영
function todayCode(){
	dayCode = document.frm.today.value;//오픈날짜 인풋값 반영
	//alert(dayCode+'로 찾아볼게요!');
}

//------------------------- 기타 영역 --------------------------------//


//엑셀 불러오기
function excel(Location, i){ //엑셀이 입력될 타겟, 기획전타입번호

	todayCode();

	//alert(i);
	cTemp='border-left:1px dotted #ccc;border-right:1px dotted #ccc;border-bottom:1px dotted #ccc;';
	switch (i){
		//입점
		case '0': mall='cjmall';outPut='_cj_deal';wSize='758';break;
		case '1': mall='hmall';outPut='_h_deal';wSize='778';break; 
		case '2': mall='gsshop';outPut='_gs_deal';wSize='778';break;
		case '3': mall='lottei';outPut='_lottei_deal';wSize='778';break;
		case '4': mall='naver';outPut='_naver_deal';wSize='758';break; 
		case '5': mall='11st';outPut='_e_deal';wSize='858';break; //11번가
		case '6': mall='auction';outPut='_a_deal';wSize='858';break; //옥션
		case '7': mall='gmarket';outPut='_g_deal';wSize='858';break; //G마켓
		case '8': mall='g9';outPut='_g9_deal';wSize='778';break; 
		case '9': mall='tmon';outPut='_tmon_deal';wSize='768';break; 
		case '99': mall='';outPut='';wSize='758';break; //etc
		default : alert('다시 설정 해주세요.');break;
		
	};
	//alert(wSize);

	//엑셀용 변수 선언
	var Excel,oRange;

	// 엑셀 객체생성
	Excel = new ActiveXObject("Excel.Application");
	Excel.Application.Workbooks.Open('\\\\172.17.14.220\\Lguidesystem\\ip\\'+dayCode+outPut+'.xls');
	//Excel.Application.Workbooks.Open('C:/Guide/ip/'+dayCode+outPut+'.xls');
	//Excel.Application.Workbooks.Open('http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/excel/'+dayCode+outPut+'.xls'); 
	//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/ip/'+dayCode+outPut+'.xls'); 

	//엑셀표시
	Excel.Application.Visible = true;
	//sheet1 선택
	Excel.Application.Worksheets("sheet1").Activate;
	// sheet1내 작업영역을 객체에 저장
	oRange = Excel.Application.ActiveSheet.UsedRange;//oRange에는 현재 내용이 있는 부분만 선택됨.
	//oRange부분을 선택표시
	oRange.Select();
 

	//oRange부분을 브라우저에 출력
	var Type; 
	var all = document.getElementById('dev_all'); //찾고
	all.innerHTML = '';	//초기화

	var sum='';//합계값
	var oBox='';var oBoxC='';//옵션영역
	var oSec='';var oSecC='';//옵션구분영역
	var tBox='';var tBoxC='';//타이틀영역
	var tSec='';var tSecC='';//타이틀구분영역
	var cnt=0;var cntT;//카운트 변수
	var sNum; // 번호
	var ttl; //타이틀
	var pic; //사진
	var dc; //할인
	var pc; //가격
	var pct1;var pct2;var pct3; //가격명
	var pcn1;var pcn2;var pcn3; //가격
	var pc1;var pc2;var pc3; //가격 합체값
	var ee;//설명값
	var tt;//타이틀값
	var ss;//옵션값
	var stNum; //타이틀번호
	var tTtl; //타이틀
	var tpic; //사진
	var tdc; //할인
	var tpc; //가격
	var tpct1;var tpct2;var tpct3; //가격명
	var tpcn1;var tpcn2;var tpcn3; //가격
	var tpc1;var tpc2;var tpc3; //가격 합체값
	var part1; var part2; var part3;
	var tpart1; var tpart2; var tpart3;
	var emt=''; //빈영역 값

	var ex= new Array();

	//값 받기
	for(var e=12;e<=oRange.Rows.count;e++){ 
		ex[e] = new Array(); //excel값 배열 선언
		for(var j=1;j<=oRange.Columns.count;j++){ 
			ex[e][j] = oRange.Cells(e,j);

			 /*unde제거 및 기본값 설정*/
			if(ex[e][j]=='.' || ex[e][j].value == undefined || ex[e][j].value == null || ex[e][j].value == false){
				switch (j){ 
					case 2: ex[e][j]='d0';break; //디자인
					case 4: ex[e][j]='s2';break; //상품 나열 방식
					case 20: ex[e][j]='s2';break; //가격표시
					case 36: ex[e][j]='h0';break; //이미지 확대
					case 37: ex[e][j]='h2';break; //옵션 높이
					case 38: ex[e][j]='h1';break; //타이틀 높이					
					default : ex[e][j]='';break;
				}; /*switch (j) */
			} /*unde제거 및 기본값 설정*/
			//sum=sum+ex[e][j]+' ';

			if(e>=13){
				/* 값 할당 */
				if(j==59){
					
					if(ex[e][11]==''){ee='';} //설명글 재구성
					else if(ex[e][11]=='@'){ee='				<p class="e emt"></p>';}
					else{ee='				<p class="e">'+txtCng(ex[e][11])+'</p>';};
					if(ex[e][15]==''){tt='';} //타이틀 재구성
					else{tt='				<p class="t">'+txtCng(ex[e][15])+'</p>';};
					if(ex[e][16]==''){ss='';} //옵션 재구성
					else{ss='				<p class="s">'+txtCng(ex[e][16])+'</p>';};

					if(ex[e][21]==''){pct1='';} //가격명1 재구성
					else{pct1='				<span>'+ex[e][21]+'</span>';};
					if(ex[e][22]==''){pcn1='';} //가격1 재구성
					else{pcn1='				<em>'+addComma(ex[e][22])+'<span>원</span></em>';};
					
					if(ex[e][23]==''){pct2='';} //가격명2 재구성
					else{pct2='				<span>'+ex[e][23]+'</span>';};
					if(ex[e][24]==''){pcn2='';} //가격2 재구성
					else{pcn2='				<em>'+addComma(ex[e][24])+'<span>원</span></em>';};

					if(ex[e][25]==''){pct3='';} //가격명3 재구성
					else{pct3='				<span>'+ex[e][25]+'</span>';};
					if(ex[e][26]==''){pcn3='';} //가격3 재구성
					else{pcn3='				<em>'+addComma(ex[e][26])+'<span>원</span></em>';};

					ex[e][32]=round(ex[e][32],2); //
					
					if(pct1=='' && pcn1==''){pc1='';}
					else if(String(ex[e][2])=='d0' || String(ex[e][2])=='tmon0'){//기본형인 경우 reverse
						pc1=''+
								'				<div class="pc1">'+
								pcn1+pct1+
								'				</div>';	
					}
					else{
						pc1=''+
								'				<div class="pc1">'+
								pct1+pcn1+
								'				</div>';	
					}
					if(pct2=='' && pcn2==''){pc2='';}
					else if(String(ex[e][2])=='d0' || String(ex[e][2])=='tmon0'){//기본형인 경우 reverse
						pc2=''+
								'				<div class="pc2">'+
								pcn2+pct2+
								'				</div>';	
					}
					else{
						pc2=''+
								'				<div class="pc2">'+
								pct2+pcn2+
								'				</div>';	
					}
					if(pct3=='' && pcn3==''){pc3='';}
					else if(String(ex[e][2])=='d0' || String(ex[e][2])=='tmon0'){//기본형인 경우 reverse
						pc3=''+
								'				<div class="pc3">'+
								pcn3+pct3+
								'				</div>';	
					}
					else{
						pc3=''+
								'				<div class="pc3">'+
								pct3+pcn3+
								'				</div>';	
					}
					if(ex[e][27]==''){etc='';}
					else{
						etc=''+
								'				<div class="etc">'+
								ex[e][27]+
								'				</div>';
					}
					
					sNum=''+
							'<!-- 숫자영역 sNum -->'+
							'<div class="sNum '+ex[e][37]+'">'+
							'	<table class="typeA">'+
							'		<tr>'+
							'			<td>'+
							'				<span>'+ex[e][9]+'</span>'+
							'				<strong>'+ex[e][10]+'</strong>'+
							'			</td>'+
							'		</tr>'+
							'	</table>'+
							'</div>'+
							'';
					ttl=''+
							'<!-- 타이틀영역 -->'+
							'<div class="ttl '+ex[e][37]+'">'+
							'	<table class="typeA">'+
							'		<tr>'+
							'			<td>'+
											ee+
											tt+
											ss+
							'			</td>'+
							'		</tr>'+
							'	</table>'+
							'</div>'+
							'';
					if(ex[e][31]=='' && ex[e][32]=='' && ex[e][33]==''){dc='';}
					else{
						dc=''+
								'<!-- 할인영역 -->'+
								'<div class="dc">'+
								'	<table class="typeA">'+
								'		<tr>'+
								'			<td>'+
								'				<em>'+ex[e][31]+'</em>'+ //아이콘명
								'				<strong>'+ex[e][32]+'</strong>'+ //수치
								'				<span>'+ex[e][33]+'</span>'+ //단위
								'			</td>'+
								'		</tr>'+
								'	</table>'+
								'</div>'+
								'';
					};
					pic=''+
							'<!-- 사진영역 -->'+
							'<div class="pic '+ex[e][36]+'">'+
							'	<p><img src="'+picResize(ex[e][34])+'" alt="" /></p>'+					
							'</div>'+
							'';
					pc=''+
							'<!-- 가격영역 -->'+
							'<div class="prc">'+
							'	<table class="typeA">'+
							'		<tr>'+
							'			<td>'+
							'				<!-- 가격영역 -->'+
							pc1+
							pc2+
							pc3+
							etc+
							'			</td>'+
							'		</tr>'+
							'	</table>'+
							'</div>'+
							'';
					
					/* 타이틀 값 영역 할당 */
					if(ex[e][9]=='@' || ex[e][10]=='@'){
						sNum=''+
							'<!-- 숫자영역 sNum -->'+
							'<div class="sNum emt '+ex[e][37]+'">'+
							'</div>'+
							'';	
						stNum=''+
									'<div class="sNum emt  '+ex[e][38]+'">'+
									'</div>'+
									'';	
					}
					else if(ex[e][9]=='' && ex[e][10]==''){
						sNum=''+
							'<!-- 숫자영역 sNum -->'+
							'<div class="sNum emt '+ex[e][37]+'">'+
							'</div>'+
							'';	
						stNum=''+
									//'<div class="sNum emt  '+ex[e][38]+'">'+
									//'</div>'+
									'';	
					}
					else{
						stNum=''+
								'<!-- 넘버 영역 -->'+
								'<div class="sNum '+ex[e][38]+'">'+
								'	<table class="typeA">'+
								'		<tr>'+
								'			<td>'+
								'				<em>'+ex[e][9]+'</em><span>'+ex[e][10]+'</span>'+
								'			</td>'+
								'		</tr>'+
								'	</table>'+
								'</div>'+
								'';
					};
					tTtl=''+
							'<div class="ttl '+ex[e][37]+'">'+
							'	<table class="typeA">'+
							'		<tr>'+
							'			<td>'+
											ee+
											tt+
											ss+
							'			</td>'+
							'		</tr>'+
							'	</table>'+
							'</div>'+
							'';
					if(ex[e][22]=='' && ex[e][24]=='' && ex[e][26]==''){
						pc='';
						tpc='';
					}
					else{
						tpc=''+
								'<div class="prc '+ex[e][20]+' '+ex[e][38]+'">'+
								'	<table class="typeA">'+
								'		<tr>'+
								'			<td>'+
								'				<!-- 가격영역 -->'+
										pc1+
										pc2+
										pc3+
										etc+
								'			</td>'+
								'		</tr>'+
								'	</table>'+
								'</div>'+
								'';
					};
				};//값 할당
				/* 조립 */
				if(j==60){
					/* Css & Font */
					if(e==13){			
						var css; //CSS 정의 변수
						if(ex[e][2]=='l1' || ex[e][2]=='l2'){
							css='l0';
						}
						else{css=ex[e][2];}
					}; // e값이 12일때 한번만 실행
					scp='<link rel="stylesheet" type="text/css" href="css/'+css+'.css"/> '; //추가 스크립트값

					//alert(ex[e][2]);
					switch (String(ex[e][2])){
						case 'gs0':
								scp=scp+
										'<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //추가 스크립트값
								part1=''+
										'<div class="part1">'+
										sNum+
										ttl+
										'</div>'+
										'';
								part2=''+
										'<div class="part2">'+
										pic+
										dc+
										'</div>'+
										'';
								part3=''+
										'<div class="part3">'+
										pc+
										'</div>'+
										'';
								tpart1=''+
										'<div class="part1">'+
										stNum+
										'</div>'+
										'';
								tpart2=''+
										'<div class="part2">'+
										tTtl+
										'</div>'+
										'';
								tpart3=''+
										'';
							break;
						case 'gs1':
								scp=scp+
										'<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //추가 스크립트값
								part1=''+
										'<div class="part1">'+
										sNum+
										ttl+
										'</div>'+
										'';
								part2=''+
										'<div class="part2">'+
										pic+
										dc+
										'</div>'+
										'';
								part3=''+
										'<div class="part3">'+
										pc+
										'</div>'+
										'';
								tpart1=''+
										'<div class="part1">'+
										stNum+
										'</div>'+
										'';
								tpart2=''+
										'<div class="part2">'+
										tTtl+
										'</div>'+
										'';
								tpart3=''+
										'';
							break;
						case 'l0':
								scp=scp+
										'<link href="https://cdn.rawgit.com/theeluwin/NotoSansKR-Hestia/master/stylesheets/NotoSansKR-Hestia.css" rel="stylesheet" type="text/css"/>'+
										''; //추가 스크립트값
								part1=''+
										'<div class="part1">'+
										sNum+
										pic+
										'</div>'+
										'';
								part2=''+
										'<div class="part2">'+
										ttl+
										dc+
										'</div>'+
										'';
								part3=''+
										'<div class="part3">'+
										pc+
										'</div>'+
										'';
								tpart1=''+
										'<div class="part1">'+
										stNum+
										'</div>'+
										'';
								tpart2=''+
										'<div class="part2">'+
										tTtl+
										'</div>'+
										'';
								tpart3=''+
										'<div class="part3">'+
										tpc+
										'</div>'+
										'';
							break;
						case 'a0':
								scp=scp+
										'<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //추가 스크립트값
								part1=''+
										'<div class="part1">'+
										sNum+
										ttl+
										'</div>'+
										'';
								part2=''+
										'<div class="part2">'+
										pc+
										dc+
										'</div>'+
										'';
								part3=''+
										'<div class="part3">'+
										pic+
										'</div>'+
										'';
								tpart1=''+
										'<div class="part1">'+
										stNum+
										'</div>'+
										'';
								tpart2=''+
										'<div class="part2">'+
										tTtl+
										'</div>'+
										'';
								tpart3=''+
										'<div class="part3">'+
										tpc+
										dc+
										'</div>'+
										'';
							break;
						case 'tmon0' : 
								scp=scp+
										'<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //추가 스크립트값
								part1=''+
										'<div class="part1">'+
										sNum+
										ttl+
										'</div>'+
										'';
								part2=''+
										'<div class="part2">'+
										pic+
										'</div>'+
										'';
								part3=''+
										'<div class="part3">'+
										dc+
										pc+
										'</div>'+
										'';
								tpart1=''+
										'<div class="part1">'+
										stNum+
										'</div>'+
										'';
								tpart2=''+
										'<div class="part2">'+
										tTtl+
										'</div>'+
										'';
								tpart3=''+
										'<div class="part3">'+
										tpc+
										'</div>'+
										'';
							break;			
						default : 
								scp=scp+
										'<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>'+
										'<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //추가 스크립트값
								part1=''+
										'<div class="part1">'+
										sNum+
										ttl+
										'</div>'+
										'';
								part2=''+
										'<div class="part2">'+
										pic+
										'</div>'+
										'';
								part3=''+
										'<div class="part3">'+
										dc+
										pc+
										'</div>'+
										'';
								tpart1=''+
										'<div class="part1">'+
										stNum+
										'</div>'+
										'';
								tpart2=''+
										'<div class="part2">'+
										tTtl+
										'</div>'+
										'';
								tpart3=''+
										'<div class="part3">'+
										tpc+
										'</div>'+
										'';
							break;			
					}; 
					oSec=oSec+
							'<div class="oSec">'+
							'	<a href="'+ex[e][35]+'" target="_blank">'+
									part1+part2+part3+
							'	</a>'+
							'</div>'+
							'';

					tSec=''+
							'<!-- 타이틀 영역 -->'+
							'<div class="tSec">'+
									tpart1+tpart2+tpart3+
							'</div>'+
							//'<p style="clear:both;margin:0;padding:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/03_ipjum/'+ex[e][50]+'.jpg" style="display:block;margin:0 auto;" border="0" alt=""/></p>'+
							'';
					
					if(String(e-12).length==1){cntT= '0'+String(e-12)}
					else{cntT= String(e-12);};
					tSecC=''+
								'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/guide/t'+cntT+'.jpg" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0"></p>'+
								'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/b'+cntT+'.jpg" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0"></p>'+
								'';

					if(ex[13][4]=='s2' && e%2==0){ //짝수일때만 실행
						oBox=oBox+
									'<!-- 옵션리스트 영역 -->'+
									'<div class="oBox '+ex[13][4]+'">'+
									oSec+
									'</div>'+
									'';
						oSec=''; //구분값 변수 초기화
						cnt=cnt+1; //카운트
						//자릿수 변형
						if(String(cnt).length==1){cntC='0'+String(cnt);}
						else{cntC=cnt;}
						oSecC=oSecC+
									'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/guide/o'+cntC+'.jpg" usemap="#lvtMap'+cntC+'" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0">'+
									'	<map name="lvtMap'+cntC+'" id="lvtMap'+cntC+'">'+
									//'	  <area shape="rect" coords="0,0,387,648" href='+ex[e-1][35]+' title="'+ex[e-1][15]+'" target="_blank" />'+  
									//'	  <area shape="rect" coords="392,4,777,645" href='+ex[e][35]+' title="'+ex[e][15]+'" target="_blank" />'+
									'	</map>'+
									'</p>'+
									'';
					} //if 짝수값일때만 닫는다
					
					else if(e==oRange.Rows.count && e%2==1){ //홀수값처리
						oSec=''+
							'<div class="oSec">'+
							'	<a href="'+ex[e][35]+'" target="_blank">'+
									part1+part2+part3+
							'	</a>'+
							'</div>'+
							'';
						emt='<div class="oSec empty">'+
								'	<table class="typeA">'+
								'		<tr>'+
								'			<td>'+
								'				<img src="http://mall.hyundailivart.co.kr/image/mobile/renewal/logo.png" alt="현대리바트"/>'+ 		
								'			</td>'+
								'		</tr>'+
								'	</table>'+
								//'	<p>비었다아아</p>'+
								'</div>';
						emt=''+
									'<!-- 옵션리스트 영역 -->'+
									'<div class="oBox '+ex[13][4]+'">'+
									oSec+
									emt+
									'</div>'+
									'';
						oSec='';
					};


					if(ex[13][4]=='s2' && e==oRange.Rows.count && e%2==1){
						cntC=cnt+1;
						if(String(cntC).length==1){cntC='0'+String(cntC);}
						emtC='<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/guide/o'+cntC+'.jpg" usemap="#lvtMap'+cntC+'" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0">'+
									'	<map name="lvtMap'+cntC+'" id="lvtMap'+cntC+'">'+							
									//'	  <area shape="rect" coords="0,0,387,648"  href='+ex[e-1][35]+' title="'+ex[e-1][15]+'" target="_blank" />'+ //링크값
									'	</map>'+
									'</p>'+
									'';
					}
					else{emtC='';}; //if emtC 값 생성

					tBox=tBox+tSec;

					oBoxC=oSecC+emtC;
					tBoxC=tBoxC+tSecC;

				}; //j40
			}; //if e>=13

		} //for j

	} //for e
	
	
	//tBox='';
	/* 값 재구성 */
	/*for(var e=10;e<=oRange.Rows.count;e++){ 
		for(var j=1;j<=oRange.Columns.count;j++){ 
			
		};//for j
	};*/ //for e

	/*최종 결과값 산출*/
	sum =   ''+
			scp+ //Css, Font 호출값
			'<div class="oWrap" style="width:'+wSize+'px;">'+
			oBox+
			emt+
			'</div>'+
			'<!-- 타이틀 영역 -->'+
			'<div class="tBox" style="width:'+wSize+'px;">'+
			tBox+
			'</div>'+

			'<div id="dev_allC" style="display:none;">'+ //  
			'	<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/a01.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+ 
			oBoxC+
			tBoxC+
			'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/01_ismine/01_common/99_ship/01.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+ // 배송정보
			'</div>'+
			//'test'+
			'';			
	//alert(sum);	
	/*최종 결과값 산출*/

	all.innerHTML = sum;
	
}


