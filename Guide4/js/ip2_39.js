//------------------------ 시작 -------------------------------//
//기본변수
var i;//펑션 숫자 변수
var a=1;//펑션 숫자 변수
var trFs; //체크	
var temp;var txt;
var cng;var tcng;
var rowNum; //전체갯수/2
var coding; //코딩변수
var tempIco;
var tempPriOn;
var wSize;
var cTemp;
var cjTF;
var mall;//몰변수
var count;//카운트변수

//var linkArea ='background-color:cyan;opacity:0.5;filter:alpha(opacity=50);';

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

function PriOnMake(m,n){
	switch (String(m)){
		case 'O': temp=String(n);break;
		case 'X': temp='<img border="0" src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/etc/btn/btn.gif" alt="바로가기"/>';break;
		case '.': alert('가격표시는 O나 X로만 표시해주세요.');break;
	};
};

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
	//document.frm.today.value = '160119_table_pc'; //'151231_test341';
	//document.frm.today.value = document.frm.today.value;
	document.frm.today.value = year+mon+day+'_카테고리영문';
	dayCode=document.frm.today.value;
}

/*HTML 복사*/
function jsCopy(obj) {
 
 var copyObj = obj;
 var copyStr = document.getElementById(copyObj).innerHTML;
 
 window.clipboardData.setData("text", copyStr);
 alert("Html 소스가 복사되었습니다.");
}

//------------------------- 공통 영역 --------------------------------//
//오픈 날짜 반영
function todayCode(){
	dayCode = document.frm.today.value;//오픈날짜 인풋값 반영
	//alert(dayCode+'로 찾아볼게요!');
}

//------------------------- 기타 영역 --------------------------------//
//아이콘 변경
function icoMake(m){
	switch (String(m)){
		//입점
		case 'ico00': temp='<table width="65" border="0" cellpadding="0" cellspacing="0">'+
							'	<tr>'+
							'		<td height="60" align="center" valign="middle" bgcolor="#dd080c"><strong style="font:bold 28px Arial, Helvetica, sans-serif;color:#fff;">'+tempIco+'</strong><span style="font:bold 12px Arial, Helvetica, sans-serif;color:#fff;">%</span></td>'+
							'	</tr>'+
							'</table>';break;
		case 'x': temp='';break;
		default : alert('다시 설정 해주세요.');break;
		
	};
};

//가격보기
function PriOnMake(m,n){
	switch (String(m)){
		case 'O': temp=String(n);break;
		case 'X': temp='<img border="0" src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/etc/btn/btn.gif" alt="바로가기"/>';break;
		case '.': alert('가격표시는 O나 X로만 표시해주세요.');break;
	};
};

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

function txtCng(k){
	/* alert(secPic);*/
	cng = String(k);
	//alert(cng);
	cng = cng.replace(/;/gi, "<br/>");cng = cng.replace(/_/gi, "&nbsp;&nbsp;&nbsp;&nbsp;");
	cng = cng.replace("[", "<span>");cng = cng.replace("]", "</span>");
	//cng = ""+cng;
	
	return cng;	
}
function emCng(k){
	/* alert(secPic);*/
	em = String(k);
	//alert(em);
	em = em.replace("{", "<em>");em = em.replace("}", "</em>");
	
	return em;	
}

//------------------------- 엑셀 영역 --------------------------------//
//DB 불러오기
function excelAll(Location, i){

	var ExcelAll,oRangeAll;

	// 엑셀 객체생성
	ExcelAll = new ActiveXObject("Excel.Application");
	//ExcelAll.Application.Workbooks.Open('C:/Guide/db/db_150209.xls');
	//ExcelAll.Application.Workbooks.Open('\\\\172.17.14.220\\Lguidesystem\\db\\db_ip_150401.xlsx');
	//ExcelAll.Application.Workbooks.Open('http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/Excel/'+dayCode+outPut+'.xls'); 
	//ExcelAll.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/db/'+dayCode+outPut+'.xls'); 

	//엑셀표시
	ExcelAll.Application.Visible = true;

	//sheet1 선택
	ExcelAll.Application.Worksheets("오픈마켓").Activate;

	// sheet1내 작업영역을 객체에 저장
	oRangeAll = ExcelAll.Application.ActiveSheet.UsedRange;

	//oRangeAll에는 현재 내용이 있는 부분만 선택됨.
	//oRangeAll부분을 선택표시
	oRangeAll.Select();
}


/* function eSub(i){
	//enter Key
	if(event.keyCode==13){
		event.cancelBubble=false;
		//alert("되냐");
		excel('dev_all', '2'); //실제 submit function
	}
} */

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
		case '4': mall='naver';outPut='_naver_deal';wSize='858';break; 
		case '5': mall='11st';outPut='_e_deal';wSize='858';break; //11번가
		case '6': mall='auction';outPut='_a_deal';wSize='858';break; //옥션
		case '7': mall='gmarket';outPut='_g_deal';wSize='858';break; //G마켓
		case '8': mall='g9';outPut='_g9_deal';wSize='778';break; 
		case '9': mall='tmon';outPut='_tmon_deal';wSize='768';break; 
		case '99': mall='';outPut='';wSize='758';break; //etc
		default : alert('다시 설정 해주세요.');break;
		
	};
	//alert(wSize);

	/* 
	if(dayCode==''){
		alert('오픈날짜를 입력하지 않으면 제가 찾을 수 없어요ㅠ_ㅠ');
	}
	else{ */
		//엑셀용 변수 선언
		var Excel,oRange;

		count=0;//카운트변수
		hcount=0;//하프카운트변수

		// 엑셀 객체생성
		Excel = new ActiveXObject("Excel.Application");
		//Excel.Application.Workbooks.Open('\\\\192.167.22.242\\Lgudiesystem\\ip\\'+dayCode+outPut+'.xls');
		Excel.Application.Workbooks.Open('\\\\172.17.14.220\\Lguidesystem\\ip\\'+dayCode+outPut+'.xls');
		//Excel.Application.Workbooks.Open('C:/Guide/ip/'+dayCode+outPut+'.xls');
		//Excel.Application.Workbooks.Open('http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/excel/'+dayCode+outPut+'.xls'); 
		//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/ip/'+dayCode+outPut+'.xls'); 

		//엑셀표시
		Excel.Application.Visible = true;

		//sheet1 선택
		Excel.Application.Worksheets("sheet1").Activate;

		// sheet1내 작업영역을 객체에 저장
		oRange = Excel.Application.ActiveSheet.UsedRange;

		//oRange에는 현재 내용이 있는 부분만 선택됨.
		//oRange부분을 선택표시
		oRange.Select();
	 

		//oRange부분을 브라우저에 출력
		var Type; 
		var all = document.getElementById('dev_all'); //찾고
		all.innerHTML = '';	//초기화

		var wSizer = new Array();
		wSizer[2] = (Number(wSize)-30-20)/2;
		wSizer[3] = (Number(wSize)-30-20*2)/3;
		wSizer[4] = (Number(wSize)-30-18*3)/4;

		/* coding='<p style="padding:0;margin:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/cjmall/150709_deal_sofatable/guide/01.jpg" usemap="#lvMap01" style="display:block;margin:0 auto;" border="0"/>'+
				'	<map name="lvMap01" id="lvMap01">'+
				'	  <area shape="rect" coords="3,6,376,636" href="#" target="_blank" />'+
				'	  <area shape="rect" coords="381,8,757,636" href="#" target="_blank" />'+
				'	</map>'+
				'</p>'; */

		var sum='';//합계값
		var optionSum ='';//옵션 합계값 
		var optionSumC ='';//옵션 코딩 합계값 
		var ttlSum ='';//타이틀 합계값
		var ttlSumC ='';//타이틀 코딩합계값
		var num='';
		var empty ='';//홀수 짝수 제어
		var emptyC ='';//홀수 짝수 코딩
		var ex= new Array();

		for(var e=10;e<=oRange.Rows.count;e++){ 
			ex[e] = new Array(); //alt[e]값 배열선언
		
			ico=''; //아이콘 합계 변수
			for(var j=1;j<=oRange.Columns.count;j++){ 
				ex[e][j] = oRange.Cells(e,j)
				/* ex[e][j] = ex[e][j].replace(/(\,*)/g, ""); */

				//alert(String(ex[e][j]).length);
				//alert(ex[e][j]);

				if(ex[e][j]=='.' || ex[e][j].value == undefined || ex[e][j].value == null || ex[e][j].value == false){
					//ex[e][j]='';
					switch (String(j)){
						//입점
						case '1': ex[e][1]='사은품없음';break; //사은품
						case '2': ex[e][2]='d0';break; //디자인
						case '3': ex[e][3]='선택';break; //번호
						case '6': ex[e][6]=ex[e][5];break; //타이틀
						case '21': ex[e][21]='in';break; //표시형식
						case '22': ex[e][22]='h3';break; //옵션높이 
						case '23': ex[e][23]='h1';break; //타이틀높이 
						case '37': ex[e][37]='s2';break; //타이틀가격분할 
						default : ex[e][j]='';break;
						
					};
				}
				else{
					num = '				<td>'+emCng(txtCng(ex[e][3]))+'</td>';
					
					if(ex[e][4]=='.' || ex[e][j].value == undefined){	txt1 = '';	} //설명문구
					else{	txt1 = ex[e][4];	}
					if(ex[e][6]=='.'){	txt2 = ex[e][5];	} //타이틀
					else{	txt2 = '<dt>'+ex[e][6]+'</dt>';	}
					if(ex[e][7]=='.'){	txt3 = '';	} //옵션
					else{	txt3 = ex[e][7];	}
					
					if(ex[e][8]=='.'){	prcN1 = '';	} //가격명1
					else{	prcN1 = '								<dt>'+ex[e][8]+'</dt>';	} 
					if(ex[e][9]=='.'){	prcC1 = '';	} //가격1
					else if(String(ex[e][9]).match("~"))
					{	
						tmp = String(ex[e][9]).split("~");
						prcC1 = '								<dd>'+addComma(tmp[0])+'<span>원</span>~'+addComma(tmp[1])+'<span>원</span></dd>';	
					}
					else{	prcC1 = '								<dd><strike>'+addComma(round(ex[e][9],0))+'</strike><span>원</span></dd>';	}
					if(ex[e][10]=='.'){	prcN2 = '';	} //가격명2
					else{	prcN2 = '								<dt>'+ex[e][10]+'</dt>';	}
					if(ex[e][11]=='.'){	prcC2 = '';	} //가격2
					else if(String(ex[e][11]).match("~"))
					{	
						tmp = String(ex[e][11]).split("~");
						prcC2 = '								<dd>'+addComma(tmp[0])+'<span>원</span>~'+addComma(tmp[1])+'<span>원</span></dd>';	
					}
					else{	prcC2 = '								<dd>'+addComma(round(ex[e][11],0))+'<span>원</span></dd>';	}
					if(ex[e][12]=='.'){	prcN3 = '';	} //가격명3
					else{	prcN3 = '								<dt><span>'+ex[e][12]+'</span></dt>';	}
					if(ex[e][13]=='.'){	prcC3 = '';	} //가격3 
					else if(String(ex[e][13]).match("~"))
					{	
						tmp = String(ex[e][13]).split("~");
						prcC3 = '								<dd>'+addComma(tmp[0])+'<span>원</span>~'+addComma(tmp[1])+'<span>원</span></dd>';	
					}
					else{	prcC3 = '								<dd>'+addComma(round(ex[e][13],0))+'<span>원</span></dd>';	}
					if(ex[e][14]=='.'){	etc = '';	} //그외 
					else{	etc = '								<p class="etc">'+ex[e][14]+'</p>';	}
					if(ex[e][15]==''){ dc = '';} // 할인율
					else{	dc = '			<p class="dc"><span>'+round(Number(ex[e][15])*100,0)+'</span>%</p>';	}
					picResize(ex[e][19]);
					ex[e][19]=temp; //사진
				};
				

			} //for j
			
			if(e>=12 && Left(ex[e][2],1)!=='t0'){
				//번호체크
				if(ex[e][3]==''){ num='';	}
				else{num = '				<td>'+emCng(txtCng(ex[e][3]))+'</td>';}

				//설명문구,옵션 체크
				if(ex[e][4]==''){ txt1='';	}
				else{txt1 = '				<p>'+ex[e][4]+'</p>';}
				if(ex[e][7]==''){ txt3='';	}
				else{txt3 = '				<dd>'+ex[e][7]+'</dd>';}

				//가격체크
				if(ex[e][8]=='' && ex[e][9]==''){ prc1='';	}
				else{prc1='<dl class="price1">'+txtCng(prcN1)+prcC1+'</dl>';}
				if(ex[e][10]=='' && ex[e][11]==''){ prc2='';	}
				else{prc2='<dl class="price2">'+txtCng(prcN2)+prcC2+'</dl>';}
				if(ex[e][12]=='' && ex[e][13]==''){ prc3='';	}
				else{prc3='<dl class="price3">'+txtCng(prcN3)+prcC3+'</dl>';}

				//링크 체크
				if(ex[e][20]==''){ ex[e][20]='';	}
				else{ex[e][20] = 'href="'+ex[e][20]+'"';
					//alert(ex[e][20]);
				}

				//사은품체크
				if(ex[e][28]==''){}
				else{ico=ico+'<p class="p1"><img src="/Guide3/img/gift/'+ex[e][28]+'.png" alt=""/></p>'; }//alert(ico);
				if(ex[e][29]==''){}
				else{ico=ico+'<p class="p15"><img src="/Guide3/img/gift/'+ex[e][29]+'.png" alt=""/></p>'; }
				if(ex[e][30]==''){}
				else{ico=ico+'<p class="p25"><img src="/Guide3/img/gift/'+ex[e][30]+'.png" alt=""/></p>'; }
				if(ex[e][31]==''){}
				else{ico=ico+'<p class="p3"><img src="/Guide3/img/gift/'+ex[e][31]+'.png" alt=""/></p>'; }
				if(ex[e][32]==''){}
				else{ico=ico+'<p class="p7"><img src="/Guide3/img/gift/'+ex[e][32]+'.png" alt=""/></p>'; }
				if(ex[e][33]==''){}
				else{ico=ico+'<p class="p75"><img src="/Guide3/img/gift/'+ex[e][33]+'.png" alt=""/></p>'; }
				if(ex[e][34]==''){}
				else{ico=ico+'<p class="p8"><img src="/Guide3/img/gift/'+ex[e][34]+'.png" alt=""/></p>'; }
				if(ex[e][35]==''){}
				else{ico=ico+'<p class="p85"><img src="/Guide3/img/gift/'+ex[e][35]+'.png" alt=""/></p>'; }
				if(ex[e][36]==''){}
				else{ico=ico+'<p class="p9"><img src="/Guide3/img/gift/'+ex[e][36]+'.png" alt=""/></p>'; }


				if(Left(ex[e][2],2)!=='l1'){
					count=count+1; //카운트
					optionSum = optionSum +'<div class="optionSub '+ex[e][22]+'">'+
									'<a '+ex[e][20]+' target="_blank">'+
									'	<div class="slt '+ex[e][22]+'">'+
									'		<table class="typeA">'+
									'			<tr>'+
									num+
									'			</tr>'+
									'		</table>'+
									'	</div>'+
									'	<div class="cnt">'+
									
									'		<div class="ttl '+ex[e][22]+'">'+
									'			<table class="typeA">'+
									'				<tr>'+
									'					<td>'+
									txtCng(txt1)+
									'						<dl>'+
									txtCng(txt2)+
									txtCng(txt3)+
									'						</dl>'+						
									'					</td>'+
									'				</tr>'+
									'			</table>'+
									'		</div>'+
									'		<div class="pic '+ex[e][21]+'">'+
									'			<div><span><img src="'+ex[e][19]+'" alt=""/></span></div>'+
									dc+
									ico+ //340 사은품
									'		</div>'+
									'		<div class="price">'+
									'			<table class="typeA">'+
									'				<tr>'+
									'					<td>'+
									prc1+
									prc2+
									prc3+
									etc+
									'					</td>'+
									'				</tr>'+
									'			</table>'+
									'		</div>'+
									'	</div>'+
									'</a>'+
									'</div>';
				
					//빈 코딩 생성
					if(e>=12 && e%2==1){
						hcount=hcount+1;
						
						optionSumC = optionSumC +
										'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/guide/o'+cntCng(hcount)+'.jpg" usemap="#lvtMap'+cntCng(hcount)+'" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0">'+
										'	<map name="lvtMap'+cntCng(hcount)+'" id="lvtMap'+cntCng(hcount)+'">'+

										//'	  <area shape="rect" coords="3,4,387,648" href="" target="_blank" />'+ 
										//'	  <area shape="rect" coords="392,4,777,645" href="" target="_blank" />'+
										'	  <area shape="rect" coords="3,4,387,648" '+ex[e-1][20]+' title="'+ex[e-1][5]+'" target="_blank" />'+  //+oRange.Cells(11+(count*2-1),20)+
										'	  <area shape="rect" coords="392,4,777,645" '+ex[e][20]+' title="'+ex[e][5]+'" target="_blank" />'+
										'	</map>'+
										'</p>'+
										'';
					}
				}; //if

				

				//if(ex[e][1]=='t0'){
				ttlSum=ttlSum+
						'<div class="ttlSub '+ex[e][2]+'">'+
						'	<div class="slt '+ex[e][23]+'">'+
						'		<table class="typeA">'+
						'			<tr>'+
						num+
						'			</tr>'+
						'		</table>'+
						'	</div>'+
						'	<div class="cnt '+ex[e][23]+'">'+
						'		<div class="ttl">'+
						'			<table class="typeA">'+
						'				<tr>'+
						'					<td>'+
						txtCng(txt1)+
						'						<dl>'+
						txtCng(txt2)+
						txtCng(txt3)+
						'						</dl>'+
						'					</td>'+
						'				</tr>'+
						'			</table>'+
						'		</div>'+
						'		<div class="price '+ex[e][37]+'" style="display:none">'+
						'			<table class="typeA">'+
						'				<tr>'+
						'					<td>'+
						dc+
						prc1+
						prc2+
						prc3+
						'					</td>'+
						'				</tr>'+
						'			</table>'+
						'		</div>'+
						'	</div>'+
						'</div>'+
						'';
				ttlSumC=ttlSumC+
						//'<p style="padding:0;margin:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/MALL/YYMMDD_CATE_deal/guide/t00.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+
						//'<p style="padding:0;margin:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/MALL/YYMMDD_CATE_deal/00.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+
						
						//'<p style="padding:0;margin:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/guide/t'+cntCng(count)+'.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+
						//'<p style="padding:0;margin:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/b'+cntCng(count)+'.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+
						
						'<p style="padding:0;margin:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/guide/t'+cntCng(count)+'.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+ //340  title="'+ex[e][5]+' 타이틀" 
						'<p style="padding:0;margin:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/b'+cntCng(count)+'.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+ //340   title="'+ex[e][5]+' 상세기술서"
						'';
							
				//};
			}//if e값체크 12부터
		}

		if(count%2==1){ //카운트 최종값이 홀수라면 
			//alert("홀수구먼");
			empty='<div class="optionSub">'+
					'	<p class="empty"></p>'+
					'</div>';
			emptyC='<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/guide/o'+cntCng(round(count/2,0))+'.jpg" usemap="#lvtMap'+cntCng(round(count/2,0))+'" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0">'+
							'	<map name="lvtMap'+cntCng(round(count/2,0))+'" id="lvtMap'+cntCng(round(count/2,0))+'">'+							
							//'	  <area shape="rect" coords="3,4,387,648" href="" target="_blank" />'+
							'	  <area shape="rect" coords="3,4,387,648" '+ex[e-1][20]+' title="'+ex[e-1][5]+'" target="_blank" />'+ //340 링크값
							'	</map>'+
							'</p>'+
							'';
		}
		else{
			empty='';
			emptyC='';
		}

		//alert(wSize);
		sum =   ''+
				'<div class="option '+ex[12][2]+'" style="width:'+wSize+'px;">'+
					optionSum+
					empty+
				'</div>'+
			    '<div class="ttlBox '+ex[12][2]+'" style="width:'+wSize+'px;">'+
					ttlSum+
				'</div>'+
				'<div id="dev_allC" style="display:none;">'+
					'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/'+mall+'/'+dayCode+'_deal/a01.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+ //342 상단정보
					optionSumC+
					emptyC+
					ttlSumC+
					'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/01_ismine/01_common/99_ship/01.jpg" style="display:block;margin:0 auto;" border="0"/></p>'+ //342 배송정보
				'</div>'+
				'';			

		all.innerHTML = sum;
	/* } */
}

