//------------------------ 시작 -------------------------------//
//기본변수
var i;//펑션 숫자 변수
var temp;var txt;
var wSize;
var mall;//몰변수
var scp;//CSS & font값 변수
var loc='\\\\172.17.14.220\\Lguidesystem\\ja\\'; // 'C:/Guide/ja/'; //

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
	document.frm.today.value = '171231_pickBasicNew';
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
function spanCng(k){
	/* alert(secPic);*/
	cng2 = String(k);
	//alert(cng);
	cng2 = cng2.replace("(", "<span>(");
	cng2 = cng2.replace(")", ")</span>");
	//cng2 = ""+cng2;
	
	return cng2;	
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
		//자사
		case '0': outPut='_week';break;
		case '1': outPut='_jung';break; 
		case '2': outPut='_bed';break;
		case '3': outPut='_room';break;
		case '4': outPut='_kitchen';break; 
		case '5': outPut='_system';break; 
		case '6': outPut='_book';break; 
		case '7': outPut='_office';break; 
		case '8': outPut='_kids';break; 
		case '9': outPut='_mallin';break;
		case '10': outPut='_home';break;
		case '99': outPut='';break; //etc
		default : j=99; alert('다시 설정 해주세요.');break;		
	};
	
	//엑셀용 변수 선언
	var Excel,oRange;

	// 엑셀 객체생성
	Excel = new ActiveXObject("Excel.Application");
	Excel.Application.Workbooks.Open(loc+dayCode+outPut+'.xls');
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
	var osum='';var tsum='';//oSec, tSec 합계값
	var oBox='';var oBoxC='';//옵션영역
	var oSec='';var oSecC='';//옵션구분영역
	var tBox='';var tBoxC='';//타이틀영역
	var tSec='';var tSecC='';//타이틀구분영역
	var cnt=0;var cntT;//카운트 변수
	var cntA=0;var cntB=0;
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
	for(var e=13;e<=oRange.Rows.count;e++){ 
		ex[e] = new Array(); //excel값 배열 선언
		for(var j=1;j<=45;j++){ 
			ex[e][j] = oRange.Cells(e,j);
			
			 /*unde제거 및 기본값 설정*/
			if(ex[e][j]=='.' || ex[e][j].value == undefined || ex[e][j].value == null || ex[e][j].value == false || ex[e][j]=='x'){
				switch (j){ 
					case 3: ex[e][j]='d0';break; //디자인
					case 10: ex[e][j]='jpg';break; //확장자
					case 29: ex[e][j]='10';break; //쿠폰율
					case 36: ex[e][j]='<dt class="pic"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg" alt=""/></dt>';break; //이미지url
					case 37: ex[e][j]='z1';break; //이미지 확대
					case 38: ex[e][j]='h2';break; //이미지 위치
					default : ex[e][j]='';break;
				}; /*switch (j) */
			}
			else{
				switch (j){ 
					//case 11: ex[e][j]=''+ex[e][j]+'';break; //
					//case 12: ex[e][j]='<dd class="ttl"><span>['+ex[e][12]+']</span>'+ex[e][11]+'</dd>';break; //
					case 4: ex[e][j]='<span>'+ex[e][j]+'</span>';break; //
					case 5: ex[e][j]='<strong>'+ex[e][j]+'</strong>';break; //
					case 13: ex[e][j]='	<dd class="txt">'+ex[e][j]+'</dd>';break; //설명문구
					case 14: ex[e][j]=addComma(Number(ex[e][j]));break; //정상가
					case 15: ex[e][j]=addComma(Number(ex[e][j]));break; //할인가
					case 16: ex[e][j]=addComma(Number(ex[e][j]));break; //최종가
					case 22: //출시여부
						if(ex[e][j]=="soldout"){ex[e][j]='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
						else if(ex[e][j]=="upcoming"){ex[e][j]='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
						//ex[e][j]='['+ex[e][j]+']';
						break; 
					case 24:
						if(ex[e][23]=="ico00"){
							ex[e][24]=''+
								'<table class="'+ex[e][23]+'">'+
								'	<tbody>'+
								'	  <tr>'+
								'		<td><strong>'+ex[e][24]+'</strong> <span>%</span></td>'+
								'	  </tr>'+
								'	</tbody>'+
								'</table>'+
								'';
						}
						else if(ex[e][23]=="ico01"){
							ex[e][24]=''+
								'		<table class="'+ex[e][23]+'">'+
								'				<tr>'+
								'					<td><span>무료<br/>배송</span></td>'+
								'				</tr>'+
								'		</table>'+
								'';
						}
						else if(ex[e][23]=="ico02"){
							ex[e][24]=''+
								'		<table class="'+ex[e][23]+'">'+
								'				<tr>'+
								'					<td><span>신상품</span></td>'+
								'				</tr>'+
								'		</table>'+
								'';
						}
						else{
							ex[e][24] = '';
						};break;
					case 27:
						if(ex[e][26]=="ico00"){
							ex[e][27]=''+
								'<table class="'+ex[e][26]+'">'+
								'	<tbody>'+
								'	  <tr>'+
								'		<td><strong>'+ex[e][24]+'</strong> <span>%</span></td>'+
								'	  </tr>'+
								'	</tbody>'+
								'</table>'+
								'';
						}
						else if(ex[e][26]=="ico01"){
							ex[e][27]=''+
								'		<table class="'+ex[e][26]+'">'+
								'				<tr>'+
								'					<td><span>무료<br/>배송</span></td>'+
								'				</tr>'+
								'		</table>'+
								'';
						}
						else if(ex[e][26]=="ico02"){
							ex[e][2]=''+
								'		<table class="'+ex[e][26]+'">'+
								'				<tr>'+
								'					<td><span>신상품</span></td>'+
								'				</tr>'+
								'		</table>'+
								'';
						}; break;
					case 35: ex[e][j]='<div class="lk"><a href="'+ex[e][j]+'"></a></div>';break; //url 열려있는 태그
					case 36: ex[e][j]='<img src="'+ex[e][j]+'">';break; //이미지url
					case 37: ex[e][j]=' '+ex[e][j];break; //확대
					case 38: ex[e][j]=' '+ex[e][j];break; //s1위치
					case 45: ex[e][j]=' id="'+ex[e][j]+'" ';break; //id값
					
				}; /*switch (j) */
			} /*unde제거 및 기본값 설정*/
			//sum=sum+ex[e][j]+' ';

			/* 값 할당 */
			if(j==45){
				//sum = sum + ex[e][j] + '<br/><br/>';
				if(ex[e][2]=='deal1'){//딜1
					//alert("딜1");
					ex[e][5] = '											<td>'+ex[e][4]+ex[e][5]+'</td>';
					ex[e][11] = '												<p class="t">'+txtCng(spanCng(ex[e][11]))+'</p>';
					ex[e][14] = '												<div class="pc1"> <span>정상가</span> <em>'+ex[e][14]+'<span>원</span></em> </div>';
					ex[e][15] = '												<div class="pc2"> <span>최대가</span> <em>'+ex[e][15]+'<span>원</span></em> </div>';
					ex[e][16] = '												<div class="pc3"> <span>최대혜택</span> <em>'+ex[e][16]+'<span>원</span></em> </div>';
					//alert('ex['+e+']'+'['+j+'] = '+ex[e][j]);
					ex[e][27] = ex[e][24] + ex[e][27];
					ex[e][36] = '				<p>'+ex[e][36]+'</p>';
					
					osum = osum + 
						'<div class="oBox">'+
							'<div class="oSec">'+
							'		<div class="part1">'+
							'			<table class="typeA">'+
							'				<tbody>'+
							'					<tr>'+
							'						<td>'+
							'							<!-- 숫자영역 sNum -->'+
							'							<div class="sNum h1">'+
							'								<table class="typeA">'+
							'									<tbody>'+
							'										<tr>'+
							ex[e][5]+
							'										</tr>'+
							'									</tbody>'+
							'								</table>'+
							'							</div>'+
							'							<!-- 타이틀영역 -->'+
							'							<div class="ttl h1">'+
							'								<table class="typeA">'+
							'									<tbody>'+
							'										<tr>'+
							'											<td>'+
							ex[e][11]+
							'											</td>'+
							'										</tr>'+
							'									</tbody>'+
							'								</table>'+
							'							</div>'+
							'							<!-- 가격영역 -->'+
							'							<div class="prc">'+
							'								<table class="typeA">'+
							'									<tbody>'+
							'										<tr>'+
							'											<td>'+
							'												<!-- 가격영역 -->'+
							ex[e][14]+
							ex[e][15]+
							ex[e][16]+
							'											</td>'+
							'										</tr>'+
							'									</tbody>'+
							'								</table>'+
							'							</div>'+
							'						</td>'+
							'					</tr>'+
							'				</tbody>'+
							'			</table>'+
							'		</div>'+
							'		<div class="part2">'+
							'			<!-- 사진영역 -->'+
							'			<div class="pic '+ex[e][38]+' '+ex[e][37]+'">'+
							ex[e][36]+
							'			</div>'+
							'			<!-- 할인영역 -->'+
							'			<div class="dc">'+
							ex[e][27]+
							'			</div>'+
							'		</div>'+
							'		<div class="part3"></div>'+
								ex[e][35]+
							'</div>'+
						'</div>'+
						'';
					tsum = tsum +
						'';
				}
				
				else if(ex[e][2]=='box1'){//이미지1
					//ex[e][5] = '											<td>'+ex[e][4]+ex[e][5]+'</td>';
					loop = '<!-- // box1영역 시작 부분 입니다.  -->'+
						'<div class="'+ex[e][3]+'">'+
						'		<p class="'+ex[e][2]+'"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+ex[e][9]+'.'+ex[e][10]+'"></p>'+
						ex[e][35]+
						'</div>'+
						'';
					sum = sum + loop;
					loop='';
				}
				else if(ex[e][2]=='box2'){//이미지2
					k=2;
					for (l=0;l<k;l++){
						loop = loop+
							'		<li><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+oRange.Cells(e+l,9)+'.'+oRange.Cells(e+l,10)+'"></li>'+
							'';
						//alert(ex[e]);
						//e = Number(e + 1);
						//alert(ex[e]);
					}
					loop = '<!-- // box'+k+'영역 시작 부분 입니다.  -->'+
						'<div class="'+ex[e][3]+'">'+
						'	<ul class="box'+k+'">'+
						loop+
						'	</ul>'+
						'</div>'+
						'';

					sum = 	sum + loop;
				}
			}//값 할당
			else{
				//sum = sum + ' ' + ex[e][j];
			}
			
			
		} //for j

	} //for e
	
	
	/*최종 결과값 산출*/
	//sum =   ''+
	//		'';			

	//alert(sum);	
	/*최종 결과값 산출*/

	all.innerHTML = sum+
					'<div class="oWrap">'+
						osum+	
					'</div>'+
					'';
					
	
}

function nowLoc(i){
	if(i==1){
		loc = 'C:/Guide/ja';
	}
	else if(i==2){
		//loc = 'D:/Guide/ja/';
		loc = '\\\\172.17.14.220\\Lguidesystem\\ja\\';
	}

	return(loc);
}

//alert(clock());