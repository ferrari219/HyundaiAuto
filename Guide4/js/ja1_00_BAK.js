//------------------------ 시작 -------------------------------//
//기본변수
var i;//펑션 숫자 변수
var trFs; //체크	
var dayCode='151231';
var temp;
var loc='D:/Guide/'; //'\\\\172.17.14.220\\Lguidesystem\\ja\\'; // 'C:/Guide/ja/'; //
var icon='';
var icon1='';var icon2='';var icon3='';var icon4='';var icon5='';var icon6='';var icon7='';var icon8='';//아이콘
var id1='';var id2='';var id3='';var id4='';var id5='';var id6='';var id7='';var id8='';
var allCode='';

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
function addComma(n) {
    if(isNaN(n)){return 0;}
     var reg = /(^[+-]?\d+)(\d{3})/;   
     n += '';
     while (reg.test(n))
       n = n.replace(reg, '$1' + ',' + '$2');
     return n;

}
//날짜
window.onload=function(){
	now=new Date();
	var year = Right(now.getYear(),2);
	var mon = now.getMonth()+1;
	var day = now.getDate();

	if(String(mon).length==1){
		mon="0"+mon;	}
	if(String(day).length==1){
		day="0"+day;	}

	//alert(year+mon+day);
	//return year+mon+day;
	document.frm.today.value = '171231_miniBasicNew'; //
	//document.frm.today.value = year+mon+day+'_기획전명'; 
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
	temp = temp.replace("_detail6_5", "_detail6_ORIGIN"); 

	temp = temp.replace("_detail1_4", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_4", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_4", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_4", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_4", "_detail5_ORIGIN"); 
	temp = temp.replace("_detail6_4", "_detail6_ORIGIN"); 

	temp = temp.replace("_detail1_3", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_3", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_3", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_3", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_3", "_detail5_ORIGIN");
	temp = temp.replace("_detail6_3", "_detail6_ORIGIN");
	
	temp = temp.replace("_detail1_2", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_2", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_2", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_2", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_2", "_detail5_ORIGIN");
	temp = temp.replace("_detail6_2", "_detail6_ORIGIN");

	temp = temp.replace("_detail1_1", "_detail1_ORIGIN");
	temp = temp.replace("_detail2_1", "_detail2_ORIGIN");
	temp = temp.replace("_detail3_1", "_detail3_ORIGIN");
	temp = temp.replace("_detail4_1", "_detail4_ORIGIN");
	temp = temp.replace("_detail5_1", "_detail5_ORIGIN");
	temp = temp.replace("_detail6_2", "_detail6_ORIGIN");

	temp = temp.replace("THUMB_1", "ORIGIN"); 
	temp = temp.replace("THUMB_2", "ORIGIN"); 
	temp = temp.replace("THUMB_3", "ORIGIN"); 
	temp = temp.replace("THUMB_4", "ORIGIN"); 
	temp = temp.replace("THUMB_5", "ORIGIN"); 
	temp = temp.replace("THUMB_6", "ORIGIN");

	/*for(a=1;a<=5;a++){
		for(b=1;b<=6;b++){
			temp = temp.replace("_detail'+b+'_'+a+'", "_detail'+b+'_ORIGIN");
		}

		temp = temp.replace("THUMB_'+a+'", "ORIGIN");
	} */

}


//------------------------- 엑셀 영역 --------------------------------//
//엑셀 불러오기
function excel(Location, i){ //엑셀이 입력될 타겟, 기획전타입번호
	//nowLoc(2);
	todayCode();

	//alert(i);
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
		
		//입점
	};
	//alert(outPut);

	/* 
	if(dayCode==''){
		alert('오픈날짜를 입력하지 않으면 제가 찾을 수 없어요ㅠ_ㅠ');
	}
	else{ */
		//엑셀용 변수 선언
		var Excel,oRange;

		// 엑셀 객체생성
		Excel = new ActiveXObject("Excel.Application");
		//Excel.Application.Workbooks.Open('C:/Guide/ja/'+dayCode+outPut+'.xls');
		//Excel.Application.Workbooks.Open('\\\\192.167.22.242\\Lgudiesystem\\ja\\'+dayCode+outPut+'.xls');
		//Excel.Application.Workbooks.Open('http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/excel/'+dayCode+outPut+'.xls'); 
		//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/ja/'+dayCode+outPut+'.xls');
		//Excel.Application.Workbooks.Open(loc+dayCode+outPut+'.xls');
		Excel.Application.Workbooks.Open(loc+outPut+'.xls');
		//alert(loc+dayCode+outPut);

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
		all.innerHTML = ''; //'<!-- <p><a href=""><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'_week/01.jpg" alt=""/></a></p> -->';	//초기화

		for(var k=12;k<=oRange.Rows.count;k++){ 
			Type = oRange.Cells(k,2); //타입선택값
			
			//값이 있을때만 실행
			dType=oRange.Cells(k,3);//디자인타입	
			
			//alert(secNme);
			//alert(Type);
			
			//값이 있을때만 실행
			if(Type=='t0'){//타이틀
				secNme=oRange.Cells(k,11);//타이틀명
				//alert(secNme);
				allCode = allCode+
										'<div class="'+dType+'">'+
										'	<p class="t0">'+secNme+'</p>'+
										'</div>'+
										'';
			}


			else if(Type=='r2'){//롤링배너
				//alert("r5");
				pic1 = oRange.Cells(k,9); k = k+1;
				pic2 = oRange.Cells(k,9); k = k+1;
				pic3 = oRange.Cells(k,9); 
			
				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									
									'<div class="fotorama" data-autoplay="true" data-nav="thumbs" data-thumbwidth="226" data-thumbheight="124">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic1+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic2+'.jpg">'+
									'</div>'+ 
									'';			
			}

			else if(Type=='r3'){//롤링배너
				//alert("r5");
				pic1 = oRange.Cells(k,9); k = k+1;
				pic2 = oRange.Cells(k,9); k = k+1;
				pic3 = oRange.Cells(k,9); 
			
				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="fotorama" data-autoplay="true" data-nav="thumbs" data-thumbwidth="226" data-thumbheight="124">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic1+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic2+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic3+'.jpg">'+
									'</div>'+ 
									'';			
			}

			else if(Type=='r4'){//롤링배너
				//alert("r5");
				pic1 = oRange.Cells(k,9); k = k+1;
				pic2 = oRange.Cells(k,9); k = k+1;
				pic3 = oRange.Cells(k,9); k = k+1;
				pic4 = oRange.Cells(k,9); 
			
				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="fotorama" data-autoplay="true" data-nav="thumbs" data-thumbwidth="226" data-thumbheight="124">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic1+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic2+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic3+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic4+'.jpg">'+
									'</div>'+ 
									'';			
			}
			
			else if(Type=='r5'){//롤링배너
				//alert("r5");
				pic1 = oRange.Cells(k,9); k = k+1;
				pic2 = oRange.Cells(k,9); k = k+1;
				pic3 = oRange.Cells(k,9); k = k+1;
				pic4 = oRange.Cells(k,9); k = k+1;
				pic5 = oRange.Cells(k,9);
			
				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="fotorama" data-autoplay="true" data-nav="thumbs" data-thumbwidth="226" data-thumbheight="124">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic1+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic2+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic3+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic4+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic5+'.jpg">'+ 
									'</div>'+
									'';			
			}

			else if(Type=='r6'){//롤링배너
				//alert("r5");
				pic1 = oRange.Cells(k,9); k = k+1;
				pic2 = oRange.Cells(k,9); k = k+1;
				pic3 = oRange.Cells(k,9); k = k+1;
				pic4 = oRange.Cells(k,9); k = k+1;
				pic5 = oRange.Cells(k,9); k = k+1;
				pic6 = oRange.Cells(k,9);
			
				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="fotorama" data-autoplay="true" data-nav="thumbs" data-thumbwidth="226" data-thumbheight="124">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic1+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic2+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic3+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic4+'.jpg">'+
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic5+'.jpg">'+ 
									'  <img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+pic6+'.jpg">'+ 
									'</div>'+
									'';			
			}



			else if(Type=='m4'){//미니기획전
				//alert("m4");
				//var logo;var bnd;
				mbg=oRange.Cells(k,11);//BG
				if(mbg=='.' || mbg.value == undefined || mbg.value == null || mbg.value == false){	mbg='';	}
				else{mbg='background:url('+mbg+')';}
				k = k+1;

				logo=oRange.Cells(k,11);//로고값
				if(logo=='.' || logo.value == undefined || logo.value == null || logo.value == false){	logo='';	}
				else{logo='						<p class="m0"><img src="'+logo+'" alt="로고영역" /></p>';}
				k = k+1;

				bnd=oRange.Cells(k,11);//브랜드 설명
				if(bnd=='.' || bnd.value == undefined || bnd.value == null || bnd.value == false){	bnd='';	}
				else{bnd='						<p class="m1">'+bnd+'</p>';}
				k = k+1; 

				mtxt=oRange.Cells(k,11);//설명글
				if(mtxt=='.' || mtxt.value == undefined || mtxt.value == null || mtxt.value == false){	mtxt='';	}
				else{mtxt='						<p class="m2">'+mtxt+'</p>';}
				k = k+1; 

				mttl1=oRange.Cells(k,11);//타이틀1
				if(mttl1=='.' || mttl1.value == undefined || mttl1.value == null || mttl1.value == false){	mttl1='';	}
				else{mttl1='						<p class="m3">'+mttl1+'</p>';}
				k = k+1; 

				mttl2=oRange.Cells(k,11);//타이틀2
				if(mttl2=='.' || mttl2.value == undefined || mttl2.value == null || mttl2.value == false){	mttl2='';	}
				else{mttl2='						<p class="m4">'+mttl2+'</p>';}
				k = k+1; 

				mimg1=oRange.Cells(k,36);//이미지1
				if(mimg1=='.' || mimg1.value == undefined || mimg1.value == null || mimg1.value == false){	mimg1='';	}
				else{mimg1='								<td><img src="'+mimg1+'" alt="" /></td>';}
				k = k+1; 

				mimg2=oRange.Cells(k,36);//이미지2
				if(mimg2=='.' || mimg2.value == undefined || mimg2.value == null || mimg2.value == false){	mimg2='';	}
				else{mimg2='								<td><img src="'+mimg2+'" alt="" /></td>';}
				k = k+1; 

				mimg3=oRange.Cells(k,36);//이미지3
				if(mimg3=='.' || mimg3.value == undefined || mimg3.value == null || mimg3.value == false){	mimg3='';	}
				else{mimg3='								<td><img src="'+mimg3+'" alt="" /></td>';}
				k = k+1; 

				mimg4=oRange.Cells(k,36);//이미지4
				if(mimg4=='.' || mimg4.value == undefined || mimg4.value == null || mimg4.value == false){	mimg4='';	}
				else{mimg4='								<td><img src="'+mimg4+'" alt="" /></td>';}
				//alert(k);				
				
				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="mini4 '+oRange.Cells(k,3)+'" style="'+mbg+'">'+
									'	<p class="top"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/mini_top.gif" alt="" /></p>'+
									'	<div class="cont">'+
									'		<table>'+
									'			<tr>'+
									'				<td>'+
									'					<div class="lft">'+
									logo+
									bnd+
									mtxt+
									mttl1+
									mttl2+
									'					</div>'+
									'				</td>'+
									'				<td class="rgt">'+
									'						<table class="imgBox">'+
									'							<tr>'+
									mimg1+
									'								<td class="emt"></td>'+
									mimg2+
									'							</tr>'+
									'							<tr>'+
									mimg3+
									'								<td class="emt"></td>'+
									mimg4+
									'							</tr>'+
									'						</table>'+
									'				</td>'+
									'			</tr>'+
									'		</table>'+		
									'	</div>'+
									'</div>'+
									'';			
			}
			
			
			else if(Type=='box1'){//박스1
				//1번 컨텐츠
				id1=oRange.Cells(k,45);//id값
				if(id1=='.' || id1.value == undefined || id1.value == null || id1.value == false){	id1='';	}
				else{id1='id="'+id1+'"';}
				secNme1=oRange.Cells(k,11);//타이틀명
				if(secNme1=='.' || secNme1.value == undefined || secNme1.value == null || secNme1.value == false){	secNme1='';	}
				secSNme1=oRange.Cells(k,9);//파일명
				secFne1=oRange.Cells(k,10);//확장자
				if(secFne1=='.' || secFne1.value == undefined || secFne1.value == null || secFne1.value == false){	secFne1='jpg';	}
				secLnk1=oRange.Cells(k,35);//URL
				if(secLnk1=='.' || secLnk1.value == undefined || secLnk1.value == null || secLnk1.value == false){	secLnk1='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/>';	}
				else{secLnk1='<a href="'+secLnk1+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/></a>';}
				
				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'	<p '+id1+' class="box1">'+secLnk1+'</p>'+
									 '</div>'+
									 '';				
			}

			else if(Type=='box2'){//박스2
				//1번 컨텐츠
				id1=oRange.Cells(k,45);//id값
				if(id1=='.' || id1.value == undefined || id1.value == null || id1.value == false){	id1='';	}
				else{id1='id="'+id1+'"';}
				secNme1=oRange.Cells(k,11);//타이틀명
				if(secNme1=='.' || secNme1.value == undefined || secNme1.value == null || secNme1.value == false){	secNme1='';	}
				secSNme1=oRange.Cells(k,9);//파일명
				secFne1=oRange.Cells(k,10);//확장자
				if(secFne1=='.' || secFne1.value == undefined || secFne1.value == null || secFne1.value == false){	secFne1='jpg';	}
				secLnk1=oRange.Cells(k,35);//URL
				if(secLnk1=='.' || secLnk1.value == undefined || secLnk1.value == null || secLnk1.value == false){	secLnk1='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/>';	}
				else{secLnk1='<a href="'+secLnk1+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/></a>';}
				
				k=k+1;

				//2번 컨텐츠
				id2=oRange.Cells(k,45);//id값
				if(id2=='.' || id2.value == undefined || id2.value == null || id2.value == false){	id2='';	}
				else{id2='id="'+id2+'"';}
				secNme2=oRange.Cells(k,11);//타이틀명
				if(secNme2=='.' || secNme2.value == undefined || secNme2.value == null || secNme2.value == false){	secNme2='';	}
				secSNme2=oRange.Cells(k,9);//파일명
				secFne2=oRange.Cells(k,10);//확장자
				if(secFne2=='.' || secFne2.value == undefined || secFne2.value == null || secFne2.value == false){	secFne2='jpg';	}
				secLnk2=oRange.Cells(k,35);//URL
				if(secLnk2=='.' || secLnk2.value == undefined || secLnk2.value == null || secLnk2.value == false){	secLnk2='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/>';	}
				else{secLnk2='<a href="'+secLnk2+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/></a>';}

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<ul '+id2+' class="box2">'+
											'	<li>'+secLnk1+'</li>'+
											'	<li>'+secLnk2+'</li>'+
										'</ul>'+
									 '</div>'+
									 '';				
			}
			else if(Type=='box3' || Type=='box3-1' || Type=='box3-2'){//박스3
				//1번 컨텐츠
				id1=oRange.Cells(k,45);//id값
				if(id1=='.' || id1.value == undefined || id1.value == null || id1.value == false){	id1='';	}
				else{id1='id="'+id1+'"';}
				secNme1=oRange.Cells(k,11);//타이틀명
				if(secNme1=='.' || secNme1.value == undefined || secNme1.value == null || secNme1.value == false){	secNme1='';	}
				secSNme1=oRange.Cells(k,9);//파일명
				secFne1=oRange.Cells(k,10);//확장자
				if(secFne1=='.' || secFne1.value == undefined || secFne1.value == null || secFne1.value == false){	secFne1='jpg';	}
				secLnk1=oRange.Cells(k,35);//URL
				if(secLnk1=='.' || secLnk1.value == undefined || secLnk1.value == null || secLnk1.value == false){	secLnk1='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/>';	}
				else{secLnk1='<a href="'+secLnk1+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/></a>';}
				
				k=k+1;

				//2번 컨텐츠
				id2=oRange.Cells(k,45);//id값
				if(id2=='.' || id2.value == undefined || id2.value == null || id2.value == false){	id2='';	}
				else{id2='id="'+id2+'"';}
				secNme2=oRange.Cells(k,11);//타이틀명
				if(secNme2=='.' || secNme2.value == undefined || secNme2.value == null || secNme2.value == false){	secNme2='';	}
				secSNme2=oRange.Cells(k,9);//파일명
				secFne2=oRange.Cells(k,10);//확장자
				if(secFne2=='.' || secFne2.value == undefined || secFne2.value == null || secFne2.value == false){	secFne2='jpg';	}
				secLnk2=oRange.Cells(k,35);//URL
				if(secLnk2=='.' || secLnk2.value == undefined || secLnk2.value == null || secLnk2.value == false){	secLnk2='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/>';	}
				else{secLnk2='<a href="'+secLnk2+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/></a>';}

				k=k+1;

				//3번 컨텐츠
				id3=oRange.Cells(k,45);//id값
				if(id3=='.' || id3.value == undefined || id3.value == null || id3.value == false){	id3='';	}
				else{id3='id="'+id3+'"';}
				secNme3=oRange.Cells(k,11);//타이틀명
				if(secNme3=='.' || secNme3.value == undefined || secNme3.value == null || secNme3.value == false){	secNme3='';	}
				secSNme3=oRange.Cells(k,9);//파일명
				secFne3=oRange.Cells(k,10);//확장자
				if(secFne3=='.' || secFne3.value == undefined || secFne3.value == null || secFne3.value == false){	secFne3='jpg';	}
				secLnk3=oRange.Cells(k,35);//URL
				if(secLnk3=='.' || secLnk3.value == undefined || secLnk3.value == null || secLnk3.value == false){	secLnk3='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/>';	}
				else{secLnk3='<a href="'+secLnk3+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/></a>';}

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<ul '+id3+' class="'+Type+'">'+
											'	<li>'+secLnk1+'</li>'+
											'	<li>'+secLnk2+'</li>'+
											'	<li>'+secLnk3+'</li>'+
										'</ul>'+
									 '</div>'+
									 '';				
			}
			else if(Type=='box4'){//박스4
				//1번 컨텐츠
				id1=oRange.Cells(k,45);//id값
				if(id1=='.' || id1.value == undefined || id1.value == null || id1.value == false){	id1='';	}
				else{id1='id="'+id1+'"';}
				secNme1=oRange.Cells(k,11);//타이틀명
				if(secNme1=='.' || secNme1.value == undefined || secNme1.value == null || secNme1.value == false){	secNme1='';	}
				secSNme1=oRange.Cells(k,9);//파일명
				secFne1=oRange.Cells(k,10);//확장자
				if(secFne1=='.' || secFne1.value == undefined || secFne1.value == null || secFne1.value == false){	secFne1='jpg';	}
				secLnk1=oRange.Cells(k,35);//URL
				if(secLnk1=='.' || secLnk1.value == undefined || secLnk1.value == null || secLnk1.value == false){	secLnk1='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/>';	}
				else{secLnk1='<a href="'+secLnk1+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/></a>';}
				
				k=k+1;

				//2번 컨텐츠
				id2=oRange.Cells(k,45);//id값
				if(id2=='.' || id2.value == undefined || id2.value == null || id2.value == false){	id2='';	}
				else{id2='id="'+id2+'"';}
				secNme2=oRange.Cells(k,11);//타이틀명
				if(secNme2=='.' || secNme2.value == undefined || secNme2.value == null || secNme2.value == false){	secNme2='';	}
				secSNme2=oRange.Cells(k,9);//파일명
				secFne2=oRange.Cells(k,10);//확장자
				if(secFne2=='.' || secFne2.value == undefined || secFne2.value == null || secFne2.value == false){	secFne2='jpg';	}
				secLnk2=oRange.Cells(k,35);//URL
				if(secLnk2=='.' || secLnk2.value == undefined || secLnk2.value == null || secLnk2.value == false){	secLnk2='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/>';	}
				else{secLnk2='<a href="'+secLnk2+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/></a>';}

				k=k+1;

				//3번 컨텐츠
				id3=oRange.Cells(k,45);//id값
				if(id3=='.' || id3.value == undefined || id3.value == null || id3.value == false){	id3='';	}
				else{id3='id="'+id3+'"';}
				secNme3=oRange.Cells(k,11);//타이틀명
				if(secNme3=='.' || secNme3.value == undefined || secNme3.value == null || secNme3.value == false){	secNme3='';	}
				secSNme3=oRange.Cells(k,9);//파일명
				secFne3=oRange.Cells(k,10);//확장자
				if(secFne3=='.' || secFne3.value == undefined || secFne3.value == null || secFne3.value == false){	secFne3='jpg';	}
				secLnk3=oRange.Cells(k,35);//URL
				if(secLnk3=='.' || secLnk3.value == undefined || secLnk3.value == null || secLnk3.value == false){	secLnk3='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/>';	}
				else{secLnk3='<a href="'+secLnk3+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/></a>';}

				k=k+1;

				//4번 컨텐츠
				id4=oRange.Cells(k,45);//id값
				if(id4=='.' || id4.value == undefined || id4.value == null || id4.value == false){	id4='';	}
				else{id4='id="'+id4+'"';}
				secNme4=oRange.Cells(k,11);//타이틀명
				if(secNme4=='.' || secNme4.value == undefined || secNme4.value == null || secNme4.value == false){	secNme4='';	}
				secSNme4=oRange.Cells(k,9);//파일명
				secFne4=oRange.Cells(k,10);//확장자
				if(secFne4=='.' || secFne4.value == undefined || secFne4.value == null || secFne4.value == false){	secFne4='jpg';	}
				secLnk4=oRange.Cells(k,35);//URL
				if(secLnk4=='.' || secLnk4.value == undefined || secLnk4.value == null || secLnk4.value == false){	secLnk4='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme4+'.'+secFne4+'"  alt="'+secNme4+'"/>';	}
				else{secLnk4='<a href="'+secLnk4+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme4+'.'+secFne4+'"  alt="'+secNme4+'"/></a>';}

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<ul '+id4+' class="box4">'+
											'	<li>'+secLnk1+'</li>'+
											'	<li>'+secLnk2+'</li>'+
											'	<li>'+secLnk3+'</li>'+
											'	<li>'+secLnk4+'</li>'+
										'</ul>'+
									 '</div>'+
									 '';				
			}
			else if(Type=='box5'){//박스5
				//1번 컨텐츠
				id1=oRange.Cells(k,45);//id값
				if(id1=='.' || id1.value == undefined || id1.value == null || id1.value == false){	id1='';	}
				else{id1='id="'+id1+'"';}
				secNme1=oRange.Cells(k,11);//타이틀명
				if(secNme1=='.' || secNme1.value == undefined || secNme1.value == null || secNme1.value == false){	secNme1='';	}
				secSNme1=oRange.Cells(k,9);//파일명
				secFne1=oRange.Cells(k,10);//확장자
				if(secFne1=='.' || secFne1.value == undefined || secFne1.value == null || secFne1.value == false){	secFne1='jpg';	}
				secLnk1=oRange.Cells(k,35);//URL
				if(secLnk1=='.' || secLnk1.value == undefined || secLnk1.value == null || secLnk1.value == false){	secLnk1='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/>';	}
				else{secLnk1='<a href="'+secLnk1+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/></a>';}
				
				k=k+1;

				//2번 컨텐츠
				id2=oRange.Cells(k,45);//id값
				if(id2=='.' || id2.value == undefined || id2.value == null || id2.value == false){	id2='';	}
				else{id2='id="'+id2+'"';}
				secNme2=oRange.Cells(k,11);//타이틀명
				if(secNme2=='.' || secNme2.value == undefined || secNme2.value == null || secNme2.value == false){	secNme2='';	}
				secSNme2=oRange.Cells(k,9);//파일명
				secFne2=oRange.Cells(k,10);//확장자
				if(secFne2=='.' || secFne2.value == undefined || secFne2.value == null || secFne2.value == false){	secFne2='jpg';	}
				secLnk2=oRange.Cells(k,35);//URL
				if(secLnk2=='.' || secLnk2.value == undefined || secLnk2.value == null || secLnk2.value == false){	secLnk2='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/>';	}
				else{secLnk2='<a href="'+secLnk2+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/></a>';}

				k=k+1;

				//3번 컨텐츠
				id3=oRange.Cells(k,45);//id값
				if(id3=='.' || id3.value == undefined || id3.value == null || id3.value == false){	id3='';	}
				else{id3='id="'+id3+'"';}
				secNme3=oRange.Cells(k,11);//타이틀명
				if(secNme3=='.' || secNme3.value == undefined || secNme3.value == null || secNme3.value == false){	secNme3='';	}
				secSNme3=oRange.Cells(k,9);//파일명
				secFne3=oRange.Cells(k,10);//확장자
				if(secFne3=='.' || secFne3.value == undefined || secFne3.value == null || secFne3.value == false){	secFne3='jpg';	}
				secLnk3=oRange.Cells(k,35);//URL
				if(secLnk3=='.' || secLnk3.value == undefined || secLnk3.value == null || secLnk3.value == false){	secLnk3='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/>';	}
				else{secLnk3='<a href="'+secLnk3+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/></a>';}

				k=k+1;

				//4번 컨텐츠
				id4=oRange.Cells(k,45);//id값
				if(id4=='.' || id4.value == undefined || id4.value == null || id4.value == false){	id4='';	}
				else{id4='id="'+id4+'"';}
				secNme4=oRange.Cells(k,11);//타이틀명
				if(secNme4=='.' || secNme4.value == undefined || secNme4.value == null || secNme4.value == false){	secNme4='';	}
				secSNme4=oRange.Cells(k,9);//파일명
				secFne4=oRange.Cells(k,10);//확장자
				if(secFne4=='.' || secFne4.value == undefined || secFne4.value == null || secFne4.value == false){	secFne4='jpg';	}
				secLnk4=oRange.Cells(k,35);//URL
				if(secLnk4=='.' || secLnk4.value == undefined || secLnk4.value == null || secLnk4.value == false){	secLnk4='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme4+'.'+secFne4+'"  alt="'+secNme4+'"/>';	}
				else{secLnk4='<a href="'+secLnk4+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme4+'.'+secFne4+'"  alt="'+secNme4+'"/></a>';}

				k=k+1;

				//5번 컨텐츠
				id5=oRange.Cells(k,45);//id값
				if(id5=='.' || id5.value == undefined || id5.value == null || id5.value == false){	id5='';	}
				else{id5='id="'+id5+'"';}
				secNme5=oRange.Cells(k,11);//타이틀명
				if(secNme5=='.' || secNme5.value == undefined || secNme5.value == null || secNme5.value == false){	secNme5='';	}
				secSNme5=oRange.Cells(k,9);//파일명
				secFne5=oRange.Cells(k,10);//확장자
				if(secFne5=='.' || secFne5.value == undefined || secFne5.value == null || secFne5.value == false){	secFne5='jpg';	}
				secLnk5=oRange.Cells(k,35);//URL
				if(secLnk5=='.' || secLnk5.value == undefined || secLnk5.value == null || secLnk5.value == false){	secLnk5='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme5+'.'+secFne5+'"  alt="'+secNme5+'"/>';	}
				else{secLnk5='<a href="'+secLnk5+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme5+'.'+secFne5+'"  alt="'+secNme5+'"/></a>';}

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<ul '+id5+' class="box5">'+
											'	<li>'+secLnk1+'</li>'+
											'	<li>'+secLnk2+'</li>'+
											'	<li>'+secLnk3+'</li>'+
											'	<li>'+secLnk4+'</li>'+
											'	<li>'+secLnk5+'</li>'+
										'</ul>'+
									 '</div>'+
									 '';				
			}

			else if(Type=='box6'){//박스6
				//1번 컨텐츠
				id1=oRange.Cells(k,45);//id값
				if(id1=='.' || id1.value == undefined || id1.value == null || id1.value == false){	id1='';	}
				else{id1='id="'+id1+'"';}
				secNme1=oRange.Cells(k,11);//타이틀명
				if(secNme1=='.' || secNme1.value == undefined || secNme1.value == null || secNme1.value == false){	secNme1='';	}
				secSNme1=oRange.Cells(k,9);//파일명
				secFne1=oRange.Cells(k,10);//확장자
				if(secFne1=='.' || secFne1.value == undefined || secFne1.value == null || secFne1.value == false){	secFne1='jpg';	}
				secLnk1=oRange.Cells(k,35);//URL
				if(secLnk1=='.' || secLnk1.value == undefined || secLnk1.value == null || secLnk1.value == false){	secLnk1='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/>';	}
				else{secLnk1='<a href="'+secLnk1+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme1+'.'+secFne1+'"  alt="'+secNme1+'"/></a>';}
				
				k=k+1;

				//2번 컨텐츠
				id2=oRange.Cells(k,45);//id값
				if(id2=='.' || id2.value == undefined || id2.value == null || id2.value == false){	id2='';	}
				else{id2='id="'+id2+'"';}
				secNme2=oRange.Cells(k,11);//타이틀명
				if(secNme2=='.' || secNme2.value == undefined || secNme2.value == null || secNme2.value == false){	secNme2='';	}
				secSNme2=oRange.Cells(k,9);//파일명
				secFne2=oRange.Cells(k,10);//확장자
				if(secFne2=='.' || secFne2.value == undefined || secFne2.value == null || secFne2.value == false){	secFne2='jpg';	}
				secLnk2=oRange.Cells(k,35);//URL
				if(secLnk2=='.' || secLnk2.value == undefined || secLnk2.value == null || secLnk2.value == false){	secLnk2='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/>';	}
				else{secLnk2='<a href="'+secLnk2+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme2+'.'+secFne2+'"  alt="'+secNme2+'"/></a>';}

				k=k+1;

				//3번 컨텐츠
				id3=oRange.Cells(k,45);//id값
				if(id3=='.' || id3.value == undefined || id3.value == null || id3.value == false){	id3='';	}
				else{id3='id="'+id3+'"';}
				secNme3=oRange.Cells(k,11);//타이틀명
				if(secNme3=='.' || secNme3.value == undefined || secNme3.value == null || secNme3.value == false){	secNme3='';	}
				secSNme3=oRange.Cells(k,9);//파일명
				secFne3=oRange.Cells(k,10);//확장자
				if(secFne3=='.' || secFne3.value == undefined || secFne3.value == null || secFne3.value == false){	secFne3='jpg';	}
				secLnk3=oRange.Cells(k,35);//URL
				if(secLnk3=='.' || secLnk3.value == undefined || secLnk3.value == null || secLnk3.value == false){	secLnk3='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/>';	}
				else{secLnk3='<a href="'+secLnk3+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme3+'.'+secFne3+'"  alt="'+secNme3+'"/></a>';}

				k=k+1;

				//4번 컨텐츠
				id4=oRange.Cells(k,45);//id값
				if(id4=='.' || id4.value == undefined || id4.value == null || id4.value == false){	id4='';	}
				else{id4='id="'+id4+'"';}
				secNme4=oRange.Cells(k,11);//타이틀명
				if(secNme4=='.' || secNme4.value == undefined || secNme4.value == null || secNme4.value == false){	secNme4='';	}
				secSNme4=oRange.Cells(k,9);//파일명
				secFne4=oRange.Cells(k,10);//확장자
				if(secFne4=='.' || secFne4.value == undefined || secFne4.value == null || secFne4.value == false){	secFne4='jpg';	}
				secLnk4=oRange.Cells(k,35);//URL
				if(secLnk4=='.' || secLnk4.value == undefined || secLnk4.value == null || secLnk4.value == false){	secLnk4='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme4+'.'+secFne4+'"  alt="'+secNme4+'"/>';	}
				else{secLnk4='<a href="'+secLnk4+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme4+'.'+secFne4+'"  alt="'+secNme4+'"/></a>';}

				k=k+1;

				//5번 컨텐츠
				id5=oRange.Cells(k,45);//id값
				if(id5=='.' || id5.value == undefined || id5.value == null || id5.value == false){	id5='';	}
				else{id5='id="'+id5+'"';}
				secNme5=oRange.Cells(k,11);//타이틀명
				if(secNme5=='.' || secNme5.value == undefined || secNme5.value == null || secNme5.value == false){	secNme5='';	}
				secSNme5=oRange.Cells(k,9);//파일명
				secFne5=oRange.Cells(k,10);//확장자
				if(secFne5=='.' || secFne5.value == undefined || secFne5.value == null || secFne5.value == false){	secFne5='jpg';	}
				secLnk5=oRange.Cells(k,35);//URL
				if(secLnk5=='.' || secLnk5.value == undefined || secLnk5.value == null || secLnk5.value == false){	secLnk5='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme5+'.'+secFne5+'"  alt="'+secNme5+'"/>';	}
				else{secLnk5='<a href="'+secLnk5+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme5+'.'+secFne5+'"  alt="'+secNme5+'"/></a>';}

				k=k+1;

				//6번 컨텐츠
				id6=oRange.Cells(k,45);//id값
				if(id6=='.' || id6.value == undefined || id6.value == null || id6.value == false){	id6='';	}
				else{id6='id="'+id6+'"';}
				secNme6=oRange.Cells(k,11);//타이틀명
				if(secNme6=='.' || secNme6.value == undefined || secNme6.value == null || secNme6.value == false){	secNme6='';	}
				secSNme6=oRange.Cells(k,9);//파일명
				secFne6=oRange.Cells(k,10);//확장자
				if(secFne6=='.' || secFne6.value == undefined || secFne6.value == null || secFne6.value == false){	secFne6='jpg';	}
				secLnk6=oRange.Cells(k,35);//URL
				if(secLnk6=='.' || secLnk6.value == undefined || secLnk6.value == null || secLnk6.value == false){	secLnk6='<img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme6+'.'+secFne6+'"  alt="'+secNme6+'"/>';	}
				else{secLnk6='<a href="'+secLnk6+'"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/'+dayCode+'/'+secSNme6+'.'+secFne6+'"  alt="'+secNme6+'"/></a>';}

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<ul '+id6+' class="box6">'+
											'	<li>'+secLnk1+'</li>'+
											'	<li>'+secLnk2+'</li>'+
											'	<li>'+secLnk3+'</li>'+
											'	<li>'+secLnk4+'</li>'+
											'	<li>'+secLnk5+'</li>'+
											'	<li>'+secLnk6+'</li>'+
										'</ul>'+
									 '</div>'+
									 '';	
			}

			else if(Type=='s1'){//컨텐츠1
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=='.' || secSold1.value == undefined || secSold1.value == null || secSold1.value == false || secSold1.value == 'NULL'){secSold1='';}
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				 if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';} 
				//secBrd1='['&secBrd1&']';
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'		</table>'+
									
									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									
									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1+ 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico12=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									
									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+dType+'">'+
										'<div class="s1">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'%</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+
									 '</div>'+
									 '';				
				icon1='';
			}


			else if(Type=='s2'){//컨텐츠2
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=='.' || secSold1.value == undefined || secSold1.value == null || secSold1.value == false || secSold1.value == 'NULL'){secSold1='';}
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				 if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';} 
				//secBrd1='['&secBrd1&']';
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									
									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'		</table>'+

									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico12=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									
									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;
				
				k=k+1;
				//2번 컨텐츠
				secSone2=oRange.Cells(k,38);//이미지위치
				if(secSone2=='.' || secSone2.value == undefined || secSone2.value == null || secSone2.value == false || secSone2.value == 'NULL'){secSone2='';}
				secZoom2=oRange.Cells(k,37);//확대축소
				if(secZoom2=='.' || secZoom2.value == undefined || secZoom2.value == null || secZoom2.value == false || secZoom2.value == 'NULL'){secZoom2='';}
				secSold2=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold2=='.' || secSold2.value == undefined || secSold2.value == null || secSold2.value == false || secSold2.value == 'NULL'){secSold2='';}
				if(secSold2=="soldout"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold2=="upcoming"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold2='';}
				secNme2=oRange.Cells(k,11);//타이틀명
				secBrd2=oRange.Cells(k,12);//브랜드
				 if(secBrd2=='.' || secBrd2.value == undefined || secBrd2.value == null || secBrd2.value == false || secBrd2.value == 'NULL'){secBrd2='';}
				else{secBrd2='['+secBrd2+']';} 
				//secBrd2='['+secBrd2+']';
				secSNme2=oRange.Cells(k,13);//설명글
				if(secSNme2=='.' || secSNme2.value == undefined || secSNme2.value == null || secSNme2.value == false){secSNme2='';}
				else{secSNme2='	<dd class="txt">'+secSNme2+'</dd>';};
				secPri2=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri2=oRange.Cells(k,15);//할인가
				secDPri2=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct21=round(oRange.Cells(k,24),0);//할인율
				secDct22=oRange.Cells(k,27);//그외값
				secico21=oRange.Cells(k,23);//아이콘1
				if(secico21=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct21+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico21=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+
									'';
				}
				else if(secico21=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico21=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									
									'';
				}
				else{};
				secico22=oRange.Cells(k,26);//아이콘2
				if(secico22=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct22+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico22=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico22=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+
									'';
				}
				else if(secico22=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									
									'';
				}
				else{};
				
				secCpn2=oRange.Cells(k,29);//쿠폰율
				secLnk2=oRange.Cells(k,35);//URL
				secPic2=oRange.Cells(k,36);//대표이미지
				picResize(secPic2);
				secPic2=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<div class="s2">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'%</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom2+' '+secSone2+'"><img src="'+secPic2+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold2+
											secSNme2+
											'	<dd class="ttl">'+'<span>'+secBrd2+'</span>'+secNme2+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri2+'</strike></span>'+
											'		<strong>'+secDPri2+'</strong>'+
											'	</dd>'+

											//'	<dd class="dc"><span class="'+secico21+'"><span><strong>'+secDct21+'</strong>%</span></span><span class="'+secico22+'">'+secDct22+'%</span></dd>'+ 
											'	<dd class="icoD">'+
											icon2+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk2+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+
									 '</div>'+
									 '';				
				icon1='';icon2='';
			}
			else if(Type=='s3' || Type=='s3-1' || Type=='s3-2' || Type=='s3-3'){//컨텐츠3
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=='.' || secSold1.value == undefined || secSold1.value == null || secSold1.value == false || secSold1.value == 'NULL'){secSold1='';}
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';}
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'		</table>'+
									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;
				
				k=k+1;
				//2번 컨텐츠
				secSone2=oRange.Cells(k,38);//이미지위치
				if(secSone2=='.' || secSone2.value == undefined || secSone2.value == null || secSone2.value == false || secSone2.value == 'NULL'){secSone2='';}
				secZoom2=oRange.Cells(k,37);//확대축소
				if(secZoom2=='.' || secZoom2.value == undefined || secZoom2.value == null || secZoom2.value == false || secZoom2.value == 'NULL'){secZoom2='';}
				secSold2=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold2=='.' || secSold2.value == undefined || secSold2.value == null || secSold2.value == false || secSold2.value == 'NULL'){secSold2='';}
				if(secSold2=="soldout"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold2=="upcoming"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold2='';}
				secNme2=oRange.Cells(k,11);//타이틀명
				secBrd2=oRange.Cells(k,12);//브랜드
				if(secBrd2=='.' || secBrd2.value == undefined || secBrd2.value == null || secBrd2.value == false || secBrd2.value == 'NULL'){secBrd2='';}
				else{secBrd2='['+secBrd2+']';}
				secSNme2=oRange.Cells(k,13);//설명글
				if(secSNme2=='.' || secSNme2.value == undefined || secSNme2.value == null || secSNme2.value == false){secSNme2='';}
				else{secSNme2='	<dd class="txt">'+secSNme2+'</dd>';};
				secPri2=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri2=oRange.Cells(k,15);//할인가
				secDPri2=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct21=round(oRange.Cells(k,24),0);//할인율
				secDct22=oRange.Cells(k,27);//그외값
				secico21=oRange.Cells(k,23);//아이콘1
				if(secico21=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct21+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+
									'';
				}
				else if(secico21=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico22=oRange.Cells(k,26);//아이콘2
				if(secico22=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct22+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn2=oRange.Cells(k,29);//쿠폰율
				secLnk2=oRange.Cells(k,35);//URL
				secPic2=oRange.Cells(k,36);//대표이미지
				picResize(secPic2);
				secPic2=temp;

				k=k+1;
				//3번 컨텐츠
				secSone3=oRange.Cells(k,38);//이미지위치
				if(secSone3=='.' || secSone3.value == undefined || secSone3.value == null || secSone3.value == false || secSone3.value == 'NULL'){secSone3='';}
				secZoom3=oRange.Cells(k,37);//확대축소
				if(secZoom3=='.' || secZoom3.value == undefined || secZoom3.value == null || secZoom3.value == false || secZoom3.value == 'NULL'){secZoom3='';}
				secSold3=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold3=="soldout"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold3=="upcoming"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold3='';}
				secNme3=oRange.Cells(k,11);//타이틀명
				secBrd3=oRange.Cells(k,12);//브랜드
				 if(secBrd3=='.' || secBrd3.value == undefined || secBrd3.value == null || secBrd3.value == false || secBrd3.value == 'NULL'){secBrd3='';}
				else{secBrd3='['+secBrd3+']';} 
				secSNme3=oRange.Cells(k,13);//설명글
				if(secSNme3=='.' || secSNme3.value == undefined || secSNme3.value == null || secSNme3.value == false){secSNme3='';}
				else{secSNme3='	<dd class="txt">'+secSNme3+'</dd>';};
				secPri3=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri3=oRange.Cells(k,15);//할인가
				secDPri3=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct31=round(oRange.Cells(k,24),0);//할인율
				secDct32=oRange.Cells(k,27);//그외값
				secico31=oRange.Cells(k,23);//아이콘1
				if(secico31=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct31+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico32=oRange.Cells(k,26);//아이콘2
				if(secico32=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct32+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn3=oRange.Cells(k,29);//쿠폰율
				secLnk3=oRange.Cells(k,35);//URL
				secPic3=oRange.Cells(k,36);//대표이미지
				picResize(secPic3);
				secPic3=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<div class="'+Type+'">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'%</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom2+' '+secSone2+'"><img src="'+secPic2+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold2+
											secSNme2+
											'	<dd class="ttl">'+'<span>'+secBrd2+'</span>'+secNme2+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri2+'</strike></span>'+
											'		<strong>'+secDPri2+'</strong>'+
											'	</dd>'+

											//'	<dd class="dc"><span class="'+secico21+'"><span><strong>'+secDct21+'</strong>%</span></span><span class="'+secico22+'">'+secDct22+'%</span></dd>'+ 
											'	<dd class="icoD">'+
											icon2+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk2+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom3+' '+secSone3+'"><img src="'+secPic3+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold3+
											secSNme3+
											'	<dd class="ttl">'+'<span>'+secBrd3+'</span>'+secNme3+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri3+'</strike></span>'+
											'		<strong>'+secDPri3+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico31+'"><span><strong>'+secDct31+'</strong>%</span></span><span class="'+secico32+'">'+secDct32+'%</span></dd>'+ 
											'	<dd class="icoD">'+
											icon3+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk3+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+
									 '</div>';			
				icon1='';icon2='';icon3='';
			}
			else if(Type=='s4'){//컨텐츠4
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				 if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';} 
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;
				
				k=k+1;
				//2번 컨텐츠
				secSone2=oRange.Cells(k,38);//이미지위치
				if(secSone2=='.' || secSone2.value == undefined || secSone2.value == null || secSone2.value == false || secSone2.value == 'NULL'){secSone2='';}
				secZoom2=oRange.Cells(k,37);//확대축소
				if(secZoom2=='.' || secZoom2.value == undefined || secZoom2.value == null || secZoom2.value == false || secZoom2.value == 'NULL'){secZoom2='';}
				secSold2=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold2=='.' || secSold2.value == undefined || secSold2.value == null || secSold2.value == false || secSold2.value == 'NULL'){secSold2='';}
				if(secSold2=="soldout"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold2=="upcoming"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold2='';}
				secNme2=oRange.Cells(k,11);//타이틀명
				secBrd2=oRange.Cells(k,12);//브랜드
				 if(secBrd2=='.' || secBrd2.value == undefined || secBrd2.value == null || secBrd2.value == false || secBrd2.value == 'NULL'){secBrd2='';}
				else{secBrd2='['+secBrd2+']';} 
				secSNme2=oRange.Cells(k,13);//설명글
				if(secSNme2=='.' || secSNme2.value == undefined || secSNme2.value == null || secSNme2.value == false){secSNme2='';}
				else{secSNme2='	<dd class="txt">'+secSNme2+'</dd>';};
				secPri2=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri2=oRange.Cells(k,15);//할인가
				secDPri2=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct21=round(oRange.Cells(k,24),0);//할인율
				secDct22=oRange.Cells(k,27);//그외값
				secico21=oRange.Cells(k,23);//아이콘1
				if(secico21=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct21+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico121=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico121+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico22=oRange.Cells(k,26);//아이콘2
				if(secico22=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct22+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico122+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn2=oRange.Cells(k,29);//쿠폰율
				secLnk2=oRange.Cells(k,35);//URL
				secPic2=oRange.Cells(k,36);//대표이미지
				picResize(secPic2);
				secPic2=temp;

				k=k+1;
				//3번 컨텐츠
				secSone3=oRange.Cells(k,38);//이미지위치
				if(secSone3=='.' || secSone3.value == undefined || secSone3.value == null || secSone3.value == false || secSone3.value == 'NULL'){secSone3='';}
				secZoom3=oRange.Cells(k,37);//확대축소
				if(secZoom3=='.' || secZoom3.value == undefined || secZoom3.value == null || secZoom3.value == false || secZoom3.value == 'NULL'){secZoom3='';}
				secSold3=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold3=="soldout"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold3=="upcoming"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold3='';}
				secNme3=oRange.Cells(k,11);//타이틀명
				secBrd3=oRange.Cells(k,12);//브랜드
				 if(secBrd3=='.' || secBrd3.value == undefined || secBrd3.value == null || secBrd3.value == false || secBrd3.value == 'NULL'){secBrd3='';}
				else{secBrd3='['+secBrd3+']';} 
				secSNme3=oRange.Cells(k,13);//설명글
				if(secSNme3=='.' || secSNme3.value == undefined || secSNme3.value == null || secSNme3.value == false){secSNme3='';}
				else{secSNme3='	<dd class="txt">'+secSNme3+'</dd>';};
				secPri3=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri3=oRange.Cells(k,15);//할인가
				secDPri3=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct31=round(oRange.Cells(k,24),0);//할인율
				secDct32=oRange.Cells(k,27);//그외값
				secico31=oRange.Cells(k,23);//아이콘1
				if(secico31=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct31+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico32=oRange.Cells(k,26);//아이콘2
				if(secico32=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct32+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn3=oRange.Cells(k,29);//쿠폰율
				secLnk3=oRange.Cells(k,35);//URL
				secPic3=oRange.Cells(k,36);//대표이미지
				picResize(secPic3);
				secPic3=temp;

				k=k+1;
				//4번 컨텐츠
				secSone4=oRange.Cells(k,38);//이미지위치
				if(secSone4=='.' || secSone4.value == undefined || secSone4.value == null || secSone4.value == false || secSone4.value == 'NULL'){secSone4='';}
				secZoom4=oRange.Cells(k,37);//확대축소
				if(secZoom4=='.' || secZoom4.value == undefined || secZoom4.value == null || secZoom4.value == false || secZoom4.value == 'NULL'){secZoom4='';}
				secSold4=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold4=="soldout"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold4=="upcoming"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold4='';}
				secNme4=oRange.Cells(k,11);//타이틀명
				secBrd4=oRange.Cells(k,12);//브랜드
				 if(secBrd4=='.' || secBrd4.value == undefined || secBrd4.value == null || secBrd4.value == false || secBrd4.value == 'NULL'){secBrd4='';}
				else{secBrd4='['+secBrd4+']';} 
				secSNme4=oRange.Cells(k,13);//설명글
				if(secSNme4=='.' || secSNme4.value == undefined || secSNme4.value == null || secSNme4.value == false){secSNme4='';}
				else{secSNme4='	<dd class="txt">'+secSNme4+'</dd>';};
				secPri4=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri4=oRange.Cells(k,15);//할인가
				secDPri4=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct41=round(oRange.Cells(k,24),0);//할인율
				secDct42=oRange.Cells(k,27);//그외값
				secico41=oRange.Cells(k,23);//아이콘1
				if(secico41=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct41+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico42=oRange.Cells(k,26);//아이콘2
				if(secico42=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct42+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn4=oRange.Cells(k,29);//쿠폰율
				secLnk4=oRange.Cells(k,35);//URL
				secPic4=oRange.Cells(k,36);//대표이미지
				picResize(secPic4);
				secPic4=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<div class="'+Type+'">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom2+' '+secSone2+'"><img src="'+secPic2+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold2+
											secSNme2+
											'	<dd class="ttl">'+'<span>'+secBrd2+'</span>'+secNme2+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri2+'</strike></span>'+
											'		<strong>'+secDPri2+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico21+'"><span><strong>'+secDct21+'</strong>%</span></span><span class="'+secico22+'">'+secDct22+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon2+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk2+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom3+' '+secSone3+'"><img src="'+secPic3+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold3+
											secSNme3+
											'	<dd class="ttl">'+'<span>'+secBrd3+'</span>'+secNme3+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri3+'</strike></span>'+
											'		<strong>'+secDPri3+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico31+'"><span><strong>'+secDct31+'</strong>%</span></span><span class="'+secico32+'">'+secDct32+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon3+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk3+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom4+' '+secSone4+'"><img src="'+secPic4+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold4+
											secSNme4+
											'	<dd class="ttl">'+'<span>'+secBrd4+'</span>'+secNme4+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri4+'</strike></span>'+
											'		<strong>'+secDPri4+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico41+'"><span><strong>'+secDct41+'</strong>%</span></span><span class="'+secico42+'">'+secDct42+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon4+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk4+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+
									 '</div>';				
				icon1='';icon2='';icon3='';icon4='';
			}


			else if(Type=='s5' || Type=='s5-1' || Type=='s5-2' || Type=='s5-3'){//컨텐츠5
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=='.' || secSold1.value == undefined || secSold1.value == null || secSold1.value == false || secSold1.value == 'NULL'){secSold1='';}
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				 if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';} 
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;
				
				k=k+1;
				//2번 컨텐츠
				secSone2=oRange.Cells(k,38);//이미지위치
				if(secSone2=='.' || secSone2.value == undefined || secSone2.value == null || secSone2.value == false || secSone2.value == 'NULL'){secSone2='';}
				secZoom2=oRange.Cells(k,37);//확대축소
				if(secZoom2=='.' || secZoom2.value == undefined || secZoom2.value == null || secZoom2.value == false || secZoom2.value == 'NULL'){secZoom2='';}
				secSold2=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold2=='.' || secSold2.value == undefined || secSold2.value == null || secSold2.value == false || secSold2.value == 'NULL'){secSold2='';}
				if(secSold2=="soldout"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold2=="upcoming"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold2='';}
				secNme2=oRange.Cells(k,11);//타이틀명
				secBrd2=oRange.Cells(k,12);//브랜드
				 if(secBrd2=='.' || secBrd2.value == undefined || secBrd2.value == null || secBrd2.value == false || secBrd2.value == 'NULL'){secBrd2='';}
				else{secBrd2='['+secBrd2+']';} 
				secSNme2=oRange.Cells(k,13);//설명글
				if(secSNme2=='.' || secSNme2.value == undefined || secSNme2.value == null || secSNme2.value == false){secSNme2='';}
				else{secSNme2='	<dd class="txt">'+secSNme2+'</dd>';};
				secPri2=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri2=oRange.Cells(k,15);//할인가
				secDPri2=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct21=round(oRange.Cells(k,24),0);//할인율
				secDct22=oRange.Cells(k,27);//그외값
				secico21=oRange.Cells(k,23);//아이콘1
				if(secico21=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct21+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico22=oRange.Cells(k,26);//아이콘2
				if(secico22=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct22+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn2=oRange.Cells(k,29);//쿠폰율
				secLnk2=oRange.Cells(k,35);//URL
				secPic2=oRange.Cells(k,36);//대표이미지
				picResize(secPic2);
				secPic2=temp;

				k=k+1;
				//3번 컨텐츠
				secSone3=oRange.Cells(k,38);//이미지위치
				if(secSone3=='.' || secSone3.value == undefined || secSone3.value == null || secSone3.value == false || secSone3.value == 'NULL'){secSone3='';}
				secZoom3=oRange.Cells(k,37);//확대축소
				if(secZoom3=='.' || secZoom3.value == undefined || secZoom3.value == null || secZoom3.value == false || secZoom3.value == 'NULL'){secZoom3='';}
				secSold3=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold3=="soldout"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold3=="upcoming"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold3='';}
				secNme3=oRange.Cells(k,11);//타이틀명
				secBrd3=oRange.Cells(k,12);//브랜드
				 if(secBrd3=='.' || secBrd3.value == undefined || secBrd3.value == null || secBrd3.value == false || secBrd3.value == 'NULL'){secBrd3='';}
				else{secBrd3='['+secBrd3+']';} 
				secSNme3=oRange.Cells(k,13);//설명글
				if(secSNme3=='.' || secSNme3.value == undefined || secSNme3.value == null || secSNme3.value == false){secSNme3='';}
				else{secSNme3='	<dd class="txt">'+secSNme3+'</dd>';};
				secPri3=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri3=oRange.Cells(k,15);//할인가
				secDPri3=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct31=round(oRange.Cells(k,24),0);//할인율
				secDct32=oRange.Cells(k,27);//그외값
				secico31=oRange.Cells(k,23);//아이콘1
				if(secico31=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct31+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico32=oRange.Cells(k,26);//아이콘2
				if(secico32=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct32+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn3=oRange.Cells(k,29);//쿠폰율
				secLnk3=oRange.Cells(k,35);//URL
				secPic3=oRange.Cells(k,36);//대표이미지
				picResize(secPic3);
				secPic3=temp;

				k=k+1;
				//4번 컨텐츠
				secSone4=oRange.Cells(k,38);//이미지위치
				if(secSone4=='.' || secSone4.value == undefined || secSone4.value == null || secSone4.value == false || secSone4.value == 'NULL'){secSone4='';}
				secZoom4=oRange.Cells(k,37);//확대축소
				if(secZoom4=='.' || secZoom4.value == undefined || secZoom4.value == null || secZoom4.value == false || secZoom4.value == 'NULL'){secZoom4='';}
				secSold4=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold4=="soldout"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold4=="upcoming"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold4='';}
				secNme4=oRange.Cells(k,11);//타이틀명
				secBrd4=oRange.Cells(k,12);//브랜드
				 if(secBrd4=='.' || secBrd4.value == undefined || secBrd4.value == null || secBrd4.value == false || secBrd4.value == 'NULL'){secBrd4='';}
				else{secBrd4='['+secBrd4+']';} 
				secSNme4=oRange.Cells(k,13);//설명글
				if(secSNme4=='.' || secSNme4.value == undefined || secSNme4.value == null || secSNme4.value == false){secSNme4='';}
				else{secSNme4='	<dd class="txt">'+secSNme4+'</dd>';};
				secPri4=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri4=oRange.Cells(k,15);//할인가
				secDPri4=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct41=round(oRange.Cells(k,24),0);//할인율
				secDct42=oRange.Cells(k,27);//그외값
				secico41=oRange.Cells(k,23);//아이콘1
				if(secico41=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct41+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico42=oRange.Cells(k,26);//아이콘2
				if(secico42=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct42+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn4=oRange.Cells(k,29);//쿠폰율
				secLnk4=oRange.Cells(k,35);//URL
				secPic4=oRange.Cells(k,36);//대표이미지
				picResize(secPic4);
				secPic4=temp;

				k=k+1;
				//5번 컨텐츠
				secSone5=oRange.Cells(k,38);//이미지위치
				if(secSone5=='.' || secSone5.value == undefined || secSone5.value == null || secSone5.value == false || secSone5.value == 'NULL'){secSone5='';}
				secZoom5=oRange.Cells(k,37);//확대축소
				if(secZoom5=='.' || secZoom5.value == undefined || secZoom5.value == null || secZoom5.value == false || secZoom5.value == 'NULL'){secZoom5='';}
				secSold5=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold5=='.' || secSold5.value == undefined || secSold5.value == null || secSold5.value == false || secSold5.value == 'NULL'){secSold5='';}
				if(secSold5=="soldout"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold5=="upcoming"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold5='';}
				secNme5=oRange.Cells(k,11);//타이틀명
				secBrd5=oRange.Cells(k,12);//브랜드
				 if(secBrd5=='.' || secBrd5.value == undefined || secBrd5.value == null || secBrd5.value == false || secBrd5.value == 'NULL'){secBrd5='';}
				else{secBrd5='['+secBrd5+']';} 
				secSNme5=oRange.Cells(k,13);//설명글
				if(secSNme5=='.' || secSNme5.value == undefined || secSNme5.value == null || secSNme5.value == false){secSNme5='';}
				else{secSNme5='	<dd class="txt">'+secSNme5+'</dd>';};
				secPri5=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri5=oRange.Cells(k,15);//할인가
				secDPri5=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct51=round(oRange.Cells(k,24),0);//할인율
				secDct52=oRange.Cells(k,27);//그외값
				secico51=oRange.Cells(k,23);//아이콘1
				if(secico51=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct51+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico52=oRange.Cells(k,26);//아이콘2
				if(secico52=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct52+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn5=oRange.Cells(k,29);//쿠폰율
				secLnk5=oRange.Cells(k,35);//URL
				secPic5=oRange.Cells(k,36);//대표이미지
				picResize(secPic5);
				secPic5=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<div class="'+Type+'">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom2+' '+secSone2+'"><img src="'+secPic2+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold2+
											secSNme2+
											'	<dd class="ttl">'+'<span>'+secBrd2+'</span>'+secNme2+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri2+'</strike></span>'+
											'		<strong>'+secDPri2+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico21+'"><span><strong>'+secDct21+'</strong>%</span></span><span class="'+secico22+'">'+secDct22+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon2+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk2+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom3+' '+secSone3+'"><img src="'+secPic3+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold3+
											secSNme3+
											'	<dd class="ttl">'+'<span>'+secBrd3+'</span>'+secNme3+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri3+'</strike></span>'+
											'		<strong>'+secDPri3+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico31+'"><span><strong>'+secDct31+'</strong>%</span></span><span class="'+secico32+'">'+secDct32+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon3+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk3+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom4+' '+secSone4+'"><img src="'+secPic4+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold4+
											secSNme4+
											'	<dd class="ttl">'+'<span>'+secBrd4+'</span>'+secNme4+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri4+'</strike></span>'+
											'		<strong>'+secDPri4+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico41+'"><span><strong>'+secDct41+'</strong>%</span></span><span class="'+secico42+'">'+secDct42+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon4+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk4+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom5+' '+secSone5+'"><img src="'+secPic5+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold4+
											secSNme5+
											'	<dd class="ttl">'+'<span>'+secBrd5+'</span>'+secNme5+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri5+'</strike></span>'+
											'		<strong>'+secDPri5+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico51+'"><span><strong>'+secDct51+'</strong>%</span></span><span class="'+secico52+'">'+secDct52+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon5+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk5+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+
									 '</div>';				
				icon1='';icon2='';icon3='';icon4='';icon5='';
			}

			else if(Type=='s6' || Type=='s6-1' || Type=='s6-2' || Type=='s6-3'){//컨텐츠6
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				 if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';} 
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;
				
				k=k+1;
				//2번 컨텐츠
				secSone2=oRange.Cells(k,38);//이미지위치
				if(secSone2=='.' || secSone2.value == undefined || secSone2.value == null || secSone2.value == false || secSone2.value == 'NULL'){secSone2='';}
				secZoom2=oRange.Cells(k,37);//확대축소
				if(secZoom2=='.' || secZoom2.value == undefined || secZoom2.value == null || secZoom2.value == false || secZoom2.value == 'NULL'){secZoom2='';}
				secSold2=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold2=='.' || secSold2.value == undefined || secSold2.value == null || secSold2.value == false || secSold2.value == 'NULL'){secSold2='';}
				if(secSold2=="soldout"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold2=="upcoming"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold2='';}
				secNme2=oRange.Cells(k,11);//타이틀명
				secBrd2=oRange.Cells(k,12);//브랜드
				 if(secBrd2=='.' || secBrd2.value == undefined || secBrd2.value == null || secBrd2.value == false || secBrd2.value == 'NULL'){secBrd2='';}
				else{secBrd2='['+secBrd2+']';} 
				secSNme2=oRange.Cells(k,13);//설명글
				if(secSNme2=='.' || secSNme2.value == undefined || secSNme2.value == null || secSNme2.value == false){secSNme2='';}
				else{secSNme2='	<dd class="txt">'+secSNme2+'</dd>';};
				secPri2=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri2=oRange.Cells(k,15);//할인가
				secDPri2=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct21=round(oRange.Cells(k,24),0);//할인율
				secDct22=oRange.Cells(k,27);//그외값
				secico21=oRange.Cells(k,23);//아이콘1
				if(secico21=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct21+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico22=oRange.Cells(k,26);//아이콘2
				if(secico22=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct22+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn2=oRange.Cells(k,29);//쿠폰율
				secLnk2=oRange.Cells(k,35);//URL
				secPic2=oRange.Cells(k,36);//대표이미지
				picResize(secPic2);
				secPic2=temp;

				k=k+1;
				//3번 컨텐츠
				secSone3=oRange.Cells(k,38);//이미지위치
				if(secSone3=='.' || secSone3.value == undefined || secSone3.value == null || secSone3.value == false || secSone3.value == 'NULL'){secSone3='';}
				secZoom3=oRange.Cells(k,37);//확대축소
				if(secZoom3=='.' || secZoom3.value == undefined || secZoom3.value == null || secZoom3.value == false || secZoom3.value == 'NULL'){secZoom3='';}
				secSold3=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold3=="soldout"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold3=="upcoming"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold3='';}
				secNme3=oRange.Cells(k,11);//타이틀명
				secBrd3=oRange.Cells(k,12);//브랜드
				 if(secBrd3=='.' || secBrd3.value == undefined || secBrd3.value == null || secBrd3.value == false || secBrd3.value == 'NULL'){secBrd3='';}
				else{secBrd3='['+secBrd3+']';} 
				secSNme3=oRange.Cells(k,13);//설명글
				if(secSNme3=='.' || secSNme3.value == undefined || secSNme3.value == null || secSNme3.value == false){secSNme3='';}
				else{secSNme3='	<dd class="txt">'+secSNme3+'</dd>';};
				secPri3=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri3=oRange.Cells(k,15);//할인가
				secDPri3=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct31=round(oRange.Cells(k,24),0);//할인율
				secDct32=oRange.Cells(k,27);//그외값
				secico31=oRange.Cells(k,23);//아이콘1
				if(secico31=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct31+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico32=oRange.Cells(k,26);//아이콘2
				if(secico32=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct32+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn3=oRange.Cells(k,29);//쿠폰율
				secLnk3=oRange.Cells(k,35);//URL
				secPic3=oRange.Cells(k,36);//대표이미지
				picResize(secPic3);
				secPic3=temp;

				k=k+1;
				//4번 컨텐츠
				secSone4=oRange.Cells(k,38);//이미지위치
				if(secSone4=='.' || secSone4.value == undefined || secSone4.value == null || secSone4.value == false || secSone4.value == 'NULL'){secSone4='';}
				secZoom4=oRange.Cells(k,37);//확대축소
				if(secZoom4=='.' || secZoom4.value == undefined || secZoom4.value == null || secZoom4.value == false || secZoom4.value == 'NULL'){secZoom4='';}
				secSold4=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold4=="soldout"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold4=="upcoming"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold4='';}
				secNme4=oRange.Cells(k,11);//타이틀명
				secBrd4=oRange.Cells(k,12);//브랜드
				 if(secBrd4=='.' || secBrd4.value == undefined || secBrd4.value == null || secBrd4.value == false || secBrd4.value == 'NULL'){secBrd4='';}
				else{secBrd4='['+secBrd4+']';} 
				secSNme4=oRange.Cells(k,13);//설명글
				if(secSNme4=='.' || secSNme4.value == undefined || secSNme4.value == null || secSNme4.value == false){secSNme4='';}
				else{secSNme4='	<dd class="txt">'+secSNme4+'</dd>';};
				secPri4=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri4=oRange.Cells(k,15);//할인가
				secDPri4=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct41=round(oRange.Cells(k,24),0);//할인율
				secDct42=oRange.Cells(k,27);//그외값
				secico41=oRange.Cells(k,23);//아이콘1
				if(secico41=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct41+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico42=oRange.Cells(k,26);//아이콘2
				if(secico42=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct42+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn4=oRange.Cells(k,29);//쿠폰율
				secLnk4=oRange.Cells(k,35);//URL
				secPic4=oRange.Cells(k,36);//대표이미지
				picResize(secPic4);
				secPic4=temp;

				k=k+1;
				//5번 컨텐츠
				secSone5=oRange.Cells(k,38);//이미지위치
				if(secSone5=='.' || secSone5.value == undefined || secSone5.value == null || secSone5.value == false || secSone5.value == 'NULL'){secSone5='';}
				secZoom5=oRange.Cells(k,37);//확대축소
				if(secZoom5=='.' || secZoom5.value == undefined || secZoom5.value == null || secZoom5.value == false || secZoom5.value == 'NULL'){secZoom5='';}
				secSold5=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold5=='.' || secSold5.value == undefined || secSold5.value == null || secSold5.value == false || secSold5.value == 'NULL'){secSold5='';}
				if(secSold5=="soldout"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold5=="upcoming"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold5='';}
				secNme5=oRange.Cells(k,11);//타이틀명
				secBrd5=oRange.Cells(k,12);//브랜드
				 if(secBrd5=='.' || secBrd5.value == undefined || secBrd5.value == null || secBrd5.value == false || secBrd5.value == 'NULL'){secBrd5='';}
				else{secBrd5='['+secBrd5+']';} 
				secSNme5=oRange.Cells(k,13);//설명글
				if(secSNme5=='.' || secSNme5.value == undefined || secSNme5.value == null || secSNme5.value == false){secSNme5='';}
				else{secSNme5='	<dd class="txt">'+secSNme5+'</dd>';};
				secPri5=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri5=oRange.Cells(k,15);//할인가
				secDPri5=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct51=round(oRange.Cells(k,24),0);//할인율
				secDct52=oRange.Cells(k,27);//그외값
				secico51=oRange.Cells(k,23);//아이콘1
				if(secico51=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct51+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico52=oRange.Cells(k,26);//아이콘2
				if(secico52=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct52+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn5=oRange.Cells(k,29);//쿠폰율
				secLnk5=oRange.Cells(k,35);//URL
				secPic5=oRange.Cells(k,36);//대표이미지
				picResize(secPic5);
				secPic5=temp;

				k=k+1;
				//6번 컨텐츠
				secSone6=oRange.Cells(k,38);//이미지위치
				if(secSone6=='.' || secSone6.value == undefined || secSone6.value == null || secSone6.value == false || secSone6.value == 'NULL'){secSone6='';}
				secZoom6=oRange.Cells(k,37);//확대축소
				if(secZoom6=='.' || secZoom6.value == undefined || secZoom6.value == null || secZoom6.value == false || secZoom6.value == 'NULL'){secZoom6='';}
				secSold6=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold6=="soldout"){secSold6='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold6=="upcoming"){secSold6='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold6='';}
				secNme6=oRange.Cells(k,11);//타이틀명
				secBrd6=oRange.Cells(k,12);//브랜드
				 if(secBrd6=='.' || secBrd6.value == undefined || secBrd6.value == null || secBrd6.value == false || secBrd6.value == 'NULL'){secBrd6='';}
				else{secBrd6='['+secBrd6+']';} 
				secSNme6=oRange.Cells(k,13);//설명글
				if(secSNme6=='.' || secSNme6.value == undefined || secSNme6.value == null || secSNme6.value == false){secSNme6='';}
				else{secSNme6='	<dd class="txt">'+secSNme6+'</dd>';};
				secPri6=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri6=oRange.Cells(k,15);//할인가
				secDPri6=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct61=round(oRange.Cells(k,24),0);//할인율
				secDct62=oRange.Cells(k,27);//그외값
				secico61=oRange.Cells(k,23);//아이콘1
				if(secico61=='ico00'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct61+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico01'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico02'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico03'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico62=oRange.Cells(k,26);//아이콘2
				if(secico62=='ico00'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct62+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico01'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico02'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico03'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn7=oRange.Cells(k,29);//쿠폰율
				secLnk6=oRange.Cells(k,35);//URL
				secPic6=oRange.Cells(k,36);//대표이미지
				picResize(secPic6);
				secPic6=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<div class="'+Type+'">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom2+' '+secSone2+'"><img src="'+secPic2+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold2+
											secSNme2+
											'	<dd class="ttl">'+'<span>'+secBrd2+'</span>'+secNme2+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri2+'</strike></span>'+
											'		<strong>'+secDPri2+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico21+'"><span><strong>'+secDct21+'</strong>%</span></span><span class="'+secico22+'">'+secDct22+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon2+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk2+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom3+' '+secSone3+'"><img src="'+secPic3+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold3+
											secSNme3+
											'	<dd class="ttl">'+'<span>'+secBrd3+'</span>'+secNme3+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri3+'</strike></span>'+
											'		<strong>'+secDPri3+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico31+'"><span><strong>'+secDct31+'</strong>%</span></span><span class="'+secico32+'">'+secDct32+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon3+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk3+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom4+' '+secSone4+'"><img src="'+secPic4+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold4+
											secSNme4+
											'	<dd class="ttl">'+'<span>'+secBrd4+'</span>'+secNme4+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri4+'</strike></span>'+
											'		<strong>'+secDPri4+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico41+'"><span><strong>'+secDct41+'</strong>%</span></span><span class="'+secico42+'">'+secDct42+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon4+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk4+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom5+' '+secSone5+'"><img src="'+secPic5+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold5+
											secSNme5+
											'	<dd class="ttl">'+'<span>'+secBrd5+'</span>'+secNme5+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri5+'</strike></span>'+
											'		<strong>'+secDPri5+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico51+'"><span><strong>'+secDct51+'</strong>%</span></span><span class="'+secico52+'">'+secDct52+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon5+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk5+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom6+' '+secSone6+'"><img src="'+secPic6+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold6+
											secSNme6+
											'	<dd class="ttl">'+'<span>'+secBrd6+'</span>'+secNme6+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri6+'</strike></span>'+
											'		<strong>'+secDPri6+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico61+'"><span><strong>'+secDct61+'</strong>%</span></span><span class="'+secico62+'">'+secDct62+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon6+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk6+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+				
									 '</div>';				
				icon1='';icon2='';icon3='';icon4='';icon5='';icon6='';
			}

			else if(Type=='s7' || Type=='s7-1' || Type=='s7-2' || Type=='s7-3' || Type=='s7-4'){//컨텐츠7
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				 if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';} 
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;
				
				k=k+1;
				//2번 컨텐츠
				secSone2=oRange.Cells(k,38);//이미지위치
				if(secSone2=='.' || secSone2.value == undefined || secSone2.value == null || secSone2.value == false || secSone2.value == 'NULL'){secSone2='';}
				secZoom2=oRange.Cells(k,37);//확대축소
				if(secZoom2=='.' || secZoom2.value == undefined || secZoom2.value == null || secZoom2.value == false || secZoom2.value == 'NULL'){secZoom2='';}
				secSold2=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold2=='.' || secSold2.value == undefined || secSold2.value == null || secSold2.value == false || secSold2.value == 'NULL'){secSold2='';}
				if(secSold2=="soldout"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold2=="upcoming"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold2='';}
				secNme2=oRange.Cells(k,11);//타이틀명
				secBrd2=oRange.Cells(k,12);//브랜드
				 if(secBrd2=='.' || secBrd2.value == undefined || secBrd2.value == null || secBrd2.value == false || secBrd2.value == 'NULL'){secBrd2='';}
				else{secBrd2='['+secBrd2+']';} 
				secSNme2=oRange.Cells(k,13);//설명글
				if(secSNme2=='.' || secSNme2.value == undefined || secSNme2.value == null || secSNme2.value == false){secSNme2='';}
				else{secSNme2='	<dd class="txt">'+secSNme2+'</dd>';};
				secPri2=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri2=oRange.Cells(k,15);//할인가
				secDPri2=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct21=round(oRange.Cells(k,24),0);//할인율
				secDct22=oRange.Cells(k,27);//그외값
				secico21=oRange.Cells(k,23);//아이콘1
				if(secico21=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct21+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico22=oRange.Cells(k,26);//아이콘2
				if(secico22=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct22+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico03'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn2=oRange.Cells(k,29);//쿠폰율
				secLnk2=oRange.Cells(k,35);//URL
				secPic2=oRange.Cells(k,36);//대표이미지
				picResize(secPic2);
				secPic2=temp;

				k=k+1;
				//3번 컨텐츠
				secSone3=oRange.Cells(k,38);//이미지위치
				if(secSone3=='.' || secSone3.value == undefined || secSone3.value == null || secSone3.value == false || secSone3.value == 'NULL'){secSone3='';}
				secZoom3=oRange.Cells(k,37);//확대축소
				if(secZoom3=='.' || secZoom3.value == undefined || secZoom3.value == null || secZoom3.value == false || secZoom3.value == 'NULL'){secZoom3='';}
				secSold3=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold3=="soldout"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold3=="upcoming"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold3='';}
				secNme3=oRange.Cells(k,11);//타이틀명
				secBrd3=oRange.Cells(k,12);//브랜드
				 if(secBrd3=='.' || secBrd3.value == undefined || secBrd3.value == null || secBrd3.value == false || secBrd3.value == 'NULL'){secBrd3='';}
				else{secBrd3='['+secBrd3+']';} 
				secSNme3=oRange.Cells(k,13);//설명글
				if(secSNme3=='.' || secSNme3.value == undefined || secSNme3.value == null || secSNme3.value == false){secSNme3='';}
				else{secSNme3='	<dd class="txt">'+secSNme3+'</dd>';};
				secPri3=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri3=oRange.Cells(k,15);//할인가
				secDPri3=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct31=round(oRange.Cells(k,24),0);//할인율
				secDct32=oRange.Cells(k,27);//그외값
				secico31=oRange.Cells(k,23);//아이콘1
				if(secico31=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct31+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico32=oRange.Cells(k,26);//아이콘2
				if(secico32=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct32+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn3=oRange.Cells(k,29);//쿠폰율
				secLnk3=oRange.Cells(k,35);//URL
				secPic3=oRange.Cells(k,36);//대표이미지
				picResize(secPic3);
				secPic3=temp;

				k=k+1;
				//4번 컨텐츠
				secSone4=oRange.Cells(k,38);//이미지위치
				if(secSone4=='.' || secSone4.value == undefined || secSone4.value == null || secSone4.value == false || secSone4.value == 'NULL'){secSone4='';}
				secZoom4=oRange.Cells(k,37);//확대축소
				if(secZoom4=='.' || secZoom4.value == undefined || secZoom4.value == null || secZoom4.value == false || secZoom4.value == 'NULL'){secZoom4='';}
				secSold4=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold4=="soldout"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold4=="upcoming"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold4='';}
				secNme4=oRange.Cells(k,11);//타이틀명
				secBrd4=oRange.Cells(k,12);//브랜드
				 if(secBrd4=='.' || secBrd4.value == undefined || secBrd4.value == null || secBrd4.value == false || secBrd4.value == 'NULL'){secBrd4='';}
				else{secBrd4='['+secBrd4+']';} 
				secSNme4=oRange.Cells(k,13);//설명글
				if(secSNme4=='.' || secSNme4.value == undefined || secSNme4.value == null || secSNme4.value == false){secSNme4='';}
				else{secSNme4='	<dd class="txt">'+secSNme4+'</dd>';};
				secPri4=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri4=oRange.Cells(k,15);//할인가
				secDPri4=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct41=round(oRange.Cells(k,24),0);//할인율
				secDct42=oRange.Cells(k,27);//그외값
				secico41=oRange.Cells(k,23);//아이콘1
				if(secico41=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct41+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico42=oRange.Cells(k,26);//아이콘2
				if(secico42=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct42+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn4=oRange.Cells(k,29);//쿠폰율
				secLnk4=oRange.Cells(k,35);//URL
				secPic4=oRange.Cells(k,36);//대표이미지
				picResize(secPic4);
				secPic4=temp;

				k=k+1;
				//5번 컨텐츠
				secSone5=oRange.Cells(k,38);//이미지위치
				if(secSone5=='.' || secSone5.value == undefined || secSone5.value == null || secSone5.value == false || secSone5.value == 'NULL'){secSone5='';}
				secZoom5=oRange.Cells(k,37);//확대축소
				if(secZoom5=='.' || secZoom5.value == undefined || secZoom5.value == null || secZoom5.value == false || secZoom5.value == 'NULL'){secZoom5='';}
				secSold5=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold5=='.' || secSold5.value == undefined || secSold5.value == null || secSold5.value == false || secSold5.value == 'NULL'){secSold5='';}
				if(secSold5=="soldout"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold5=="upcoming"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold5='';}
				secNme5=oRange.Cells(k,11);//타이틀명
				secBrd5=oRange.Cells(k,12);//브랜드
				 if(secBrd5=='.' || secBrd5.value == undefined || secBrd5.value == null || secBrd5.value == false || secBrd5.value == 'NULL'){secBrd5='';}
				else{secBrd5='['+secBrd5+']';} 
				secSNme5=oRange.Cells(k,13);//설명글
				if(secSNme5=='.' || secSNme5.value == undefined || secSNme5.value == null || secSNme5.value == false){secSNme5='';}
				else{secSNme5='	<dd class="txt">'+secSNme5+'</dd>';};
				secPri5=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri5=oRange.Cells(k,15);//할인가
				secDPri5=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct51=round(oRange.Cells(k,24),0);//할인율
				secDct52=oRange.Cells(k,27);//그외값
				secico51=oRange.Cells(k,23);//아이콘1
				if(secico51=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct51+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico52=oRange.Cells(k,26);//아이콘2
				if(secico52=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct52+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn5=oRange.Cells(k,29);//쿠폰율
				secLnk5=oRange.Cells(k,35);//URL
				secPic5=oRange.Cells(k,36);//대표이미지
				picResize(secPic5);
				secPic5=temp;

				k=k+1;
				//6번 컨텐츠
				secSone6=oRange.Cells(k,38);//이미지위치
				if(secSone6=='.' || secSone6.value == undefined || secSone6.value == null || secSone6.value == false || secSone6.value == 'NULL'){secSone6='';}
				secZoom6=oRange.Cells(k,37);//확대축소
				if(secZoom6=='.' || secZoom6.value == undefined || secZoom6.value == null || secZoom6.value == false || secZoom6.value == 'NULL'){secZoom6='';}
				secSold6=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold6=="soldout"){secSold6='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold6=="upcoming"){secSold6='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold6='';}
				secNme6=oRange.Cells(k,11);//타이틀명
				secBrd6=oRange.Cells(k,12);//브랜드
				 if(secBrd6=='.' || secBrd6.value == undefined || secBrd6.value == null || secBrd6.value == false || secBrd6.value == 'NULL'){secBrd6='';}
				else{secBrd6='['+secBrd6+']';} 
				secSNme6=oRange.Cells(k,13);//설명글
				if(secSNme6=='.' || secSNme6.value == undefined || secSNme6.value == null || secSNme6.value == false){secSNme6='';}
				else{secSNme6='	<dd class="txt">'+secSNme6+'</dd>';};
				secPri6=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri6=oRange.Cells(k,15);//할인가
				secDPri6=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct61=round(oRange.Cells(k,24),0);//할인율
				secDct62=oRange.Cells(k,27);//그외값
				secico61=oRange.Cells(k,23);//아이콘1
				if(secico61=='ico00'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct61+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico01'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico02'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico03'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico62=oRange.Cells(k,26);//아이콘2
				if(secico62=='ico00'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct62+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico01'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico02'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico03'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn6=oRange.Cells(k,29);//쿠폰율
				secLnk6=oRange.Cells(k,35);//URL
				secPic6=oRange.Cells(k,36);//대표이미지
				picResize(secPic6);
				secPic6=temp;

				k=k+1;
				//7번 컨텐츠
				secSone7=oRange.Cells(k,38);//이미지위치
				if(secSone7=='.' || secSone7.value == undefined || secSone7.value == null || secSone7.value == false || secSone7.value == 'NULL'){secSone7='';}
				secZoom7=oRange.Cells(k,37);//확대축소
				if(secZoom7=='.' || secZoom7.value == undefined || secZoom7.value == null || secZoom7.value == false || secZoom7.value == 'NULL'){secZoom7='';}
				secSold7=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold7=="soldout"){secSold7='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold7=="upcoming"){secSold7='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold7='';}
				secNme7=oRange.Cells(k,11);//타이틀명
				secBrd7=oRange.Cells(k,12);//브랜드
				 if(secBrd7=='.' || secBrd7.value == undefined || secBrd7.value == null || secBrd7.value == false || secBrd7.value == 'NULL'){secBrd7='';}
				else{secBrd7='['+secBrd7+']';} 
				secSNme7=oRange.Cells(k,13);//설명글
				if(secSNme7=='.' || secSNme7.value == undefined || secSNme7.value == null || secSNme7.value == false){secSNme7='';}
				else{secSNme7='	<dd class="txt">'+secSNme7+'</dd>';};
				secPri7=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri7=oRange.Cells(k,15);//할인가
				secDPri7=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct71=round(oRange.Cells(k,24),0);//할인율
				secDct72=oRange.Cells(k,27);//그외값
				secico71=oRange.Cells(k,23);//아이콘1
				if(secico71=='ico00'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct71+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico71=='ico01'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico71=='ico02'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico71=='ico03'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico72=oRange.Cells(k,26);//아이콘2
				if(secico72=='ico00'){
					icon7 = icon7 + 
									'		<table class="'+secico72+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct72+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico72=='ico01'){
					icon7 = icon7 + 
									'		<table class="'+secico72+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico72=='ico02'){
					icon7 = icon7 + 
									'		<table class="'+secico72+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico71=='ico03'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn7=oRange.Cells(k,29);//쿠폰율
				secLnk7=oRange.Cells(k,35);//URL
				secPic7=oRange.Cells(k,36);//대표이미지
				picResize(secPic7);
				secPic7=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<div class="'+Type+'">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom2+' '+secSone2+'"><img src="'+secPic2+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold2+
											secSNme2+
											'	<dd class="ttl">'+'<span>'+secBrd2+'</span>'+secNme2+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri2+'</strike></span>'+
											'		<strong>'+secDPri2+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico21+'"><span><strong>'+secDct21+'</strong>%</span></span><span class="'+secico22+'">'+secDct22+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon2+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk2+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom3+' '+secSone3+'"><img src="'+secPic3+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold3+
											secSNme3+
											'	<dd class="ttl">'+'<span>'+secBrd3+'</span>'+secNme3+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri3+'</strike></span>'+
											'		<strong>'+secDPri3+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico31+'"><span><strong>'+secDct31+'</strong>%</span></span><span class="'+secico32+'">'+secDct32+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon3+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk3+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom4+' '+secSone4+'"><img src="'+secPic4+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold4+
											secSNme4+
											'	<dd class="ttl">'+'<span>'+secBrd4+'</span>'+secNme4+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri4+'</strike></span>'+
											'		<strong>'+secDPri4+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico41+'"><span><strong>'+secDct41+'</strong>%</span></span><span class="'+secico42+'">'+secDct42+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon4+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk4+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom5+' '+secSone5+'"><img src="'+secPic5+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold5+
											secSNme5+
											'	<dd class="ttl">'+'<span>'+secBrd5+'</span>'+secNme5+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri5+'</strike></span>'+
											'		<strong>'+secDPri5+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico51+'"><span><strong>'+secDct51+'</strong>%</span></span><span class="'+secico52+'">'+secDct52+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon5+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk5+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom6+' '+secSone6+'"><img src="'+secPic6+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold6+
											secSNme6+
											'	<dd class="ttl">'+'<span>'+secBrd6+'</span>'+secNme6+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri6+'</strike></span>'+
											'		<strong>'+secDPri6+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico61+'"><span><strong>'+secDct61+'</strong>%</span></span><span class="'+secico62+'">'+secDct62+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon6+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk6+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom7+' '+secSone7+'"><img src="'+secPic7+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold7+
											secSNme7+
											'	<dd class="ttl">'+'<span>'+secBrd7+'</span>'+secNme7+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri7+'</strike></span>'+
											'		<strong>'+secDPri7+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico71+'"><span><strong>'+secDct71+'</strong>%</span></span><span class="'+secico72+'">'+secDct72+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon7+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk7+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+				
									 '</div>';				
				icon1='';icon2='';icon3='';icon4='';icon5='';icon6='';icon7='';
			}

			/* 테스트 */
			else if(Type=='s8' || Type=='s8-1' || Type=='s8-2' || Type=='s8-3' || Type=='s8-4'){//컨텐츠8
				//1번 컨텐츠
				secSone1=oRange.Cells(k,38);//이미지위치
				if(secSone1=='.' || secSone1.value == undefined || secSone1.value == null || secSone1.value == false || secSone1.value == 'NULL'){secSone1='';}
				secZoom1=oRange.Cells(k,37);//확대축소
				if(secZoom1=='.' || secZoom1.value == undefined || secZoom1.value == null || secZoom1.value == false || secZoom1.value == 'NULL'){secZoom1='';}
				secSold1=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold1=="soldout"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold1=="upcoming"){secSold1='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold1='';}
				secNme1=oRange.Cells(k,11);//타이틀명
				secBrd1=oRange.Cells(k,12);//브랜드
				 if(secBrd1=='.' || secBrd1.value == undefined || secBrd1.value == null || secBrd1.value == false || secBrd1.value == 'NULL'){secBrd1='';}
				else{secBrd1='['+secBrd1+']';} 
				secSNme1=oRange.Cells(k,13);//설명글
				if(secSNme1=='.' || secSNme1.value == undefined || secSNme1.value == null || secSNme1.value == false){secSNme1='';}
				else{secSNme1='	<dd class="txt">'+secSNme1+'</dd>';};
				secPri1=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri1=oRange.Cells(k,15);//할인가
				secDPri1=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct11=round(oRange.Cells(k,24),0);//할인율
				secDct12=oRange.Cells(k,27);//그외값
				secico11=oRange.Cells(k,23);//아이콘1
				if(secico11=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct11+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico11=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico11+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico12=oRange.Cells(k,26);//아이콘2
				if(secico12=='ico00'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct12+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico01'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico02'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico12=='ico03'){
					icon1 = icon1 + 
									'		<table class="'+secico12+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn1=oRange.Cells(k,29);//쿠폰율
				secLnk1=oRange.Cells(k,35);//URL
				secPic1=oRange.Cells(k,36);//대표이미지
				picResize(secPic1);
				secPic1=temp;
				
				k=k+1;
				//2번 컨텐츠
				secSone2=oRange.Cells(k,38);//이미지위치
				if(secSone2=='.' || secSone2.value == undefined || secSone2.value == null || secSone2.value == false || secSone2.value == 'NULL'){secSone2='';}
				secZoom2=oRange.Cells(k,37);//확대축소
				if(secZoom2=='.' || secZoom2.value == undefined || secZoom2.value == null || secZoom2.value == false || secZoom2.value == 'NULL'){secZoom2='';}
				secSold2=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold2=='.' || secSold2.value == undefined || secSold2.value == null || secSold2.value == false || secSold2.value == 'NULL'){secSold2='';}
				if(secSold2=="soldout"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold2=="upcoming"){secSold2='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold2='';}
				secNme2=oRange.Cells(k,11);//타이틀명
				secBrd2=oRange.Cells(k,12);//브랜드
				 if(secBrd2=='.' || secBrd2.value == undefined || secBrd2.value == null || secBrd2.value == false || secBrd2.value == 'NULL'){secBrd2='';}
				else{secBrd2='['+secBrd2+']';} 
				secSNme2=oRange.Cells(k,13);//설명글
				if(secSNme2=='.' || secSNme2.value == undefined || secSNme2.value == null || secSNme2.value == false){secSNme2='';}
				else{secSNme2='	<dd class="txt">'+secSNme2+'</dd>';};
				secPri2=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri2=oRange.Cells(k,15);//할인가
				secDPri2=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct21=round(oRange.Cells(k,24),0);//할인율
				secDct22=oRange.Cells(k,27);//그외값
				secico21=oRange.Cells(k,23);//아이콘1
				if(secico21=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct21+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico21=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico21+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico22=oRange.Cells(k,26);//아이콘2
				if(secico22=='ico00'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct22+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico01'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico22=='ico02'){
					icon2 = icon2 + 
									'		<table class="'+secico22+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn2=oRange.Cells(k,29);//쿠폰율
				secLnk2=oRange.Cells(k,35);//URL
				secPic2=oRange.Cells(k,36);//대표이미지
				picResize(secPic2);
				secPic2=temp;

				k=k+1;
				//3번 컨텐츠
				secSone3=oRange.Cells(k,38);//이미지위치
				if(secSone3=='.' || secSone3.value == undefined || secSone3.value == null || secSone3.value == false || secSone3.value == 'NULL'){secSone3='';}
				secZoom3=oRange.Cells(k,37);//확대축소
				if(secZoom3=='.' || secZoom3.value == undefined || secZoom3.value == null || secZoom3.value == false || secZoom3.value == 'NULL'){secZoom3='';}
				secSold3=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold3=="soldout"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold3=="upcoming"){secSold3='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold3='';}
				secNme3=oRange.Cells(k,11);//타이틀명
				secBrd3=oRange.Cells(k,12);//브랜드
				 if(secBrd3=='.' || secBrd3.value == undefined || secBrd3.value == null || secBrd3.value == false || secBrd3.value == 'NULL'){secBrd3='';}
				else{secBrd3='['+secBrd3+']';} 
				secSNme3=oRange.Cells(k,13);//설명글
				if(secSNme3=='.' || secSNme3.value == undefined || secSNme3.value == null || secSNme3.value == false){secSNme3='';}
				else{secSNme3='	<dd class="txt">'+secSNme3+'</dd>';};
				secPri3=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri3=oRange.Cells(k,15);//할인가
				secDPri3=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct31=round(oRange.Cells(k,24),0);//할인율
				secDct32=oRange.Cells(k,27);//그외값
				secico31=oRange.Cells(k,23);//아이콘1
				if(secico31=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct31+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico32=oRange.Cells(k,26);//아이콘2
				if(secico32=='ico00'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct32+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico01'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico32=='ico02'){
					icon3 = icon3 + 
									'		<table class="'+secico32+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico31=='ico03'){
					icon3 = icon3 + 
									'		<table class="'+secico31+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn3=oRange.Cells(k,29);//쿠폰율
				secLnk3=oRange.Cells(k,35);//URL
				secPic3=oRange.Cells(k,36);//대표이미지
				picResize(secPic3);
				secPic3=temp;

				k=k+1;
				//4번 컨텐츠
				secSone4=oRange.Cells(k,38);//이미지위치
				if(secSone4=='.' || secSone4.value == undefined || secSone4.value == null || secSone4.value == false || secSone4.value == 'NULL'){secSone4='';}
				secZoom4=oRange.Cells(k,37);//확대축소
				if(secZoom4=='.' || secZoom4.value == undefined || secZoom4.value == null || secZoom4.value == false || secZoom4.value == 'NULL'){secZoom4='';}
				secSold4=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold4=="soldout"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold4=="upcoming"){secSold4='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold4='';}
				secNme4=oRange.Cells(k,11);//타이틀명
				secBrd4=oRange.Cells(k,12);//브랜드
				 if(secBrd4=='.' || secBrd4.value == undefined || secBrd4.value == null || secBrd4.value == false || secBrd4.value == 'NULL'){secBrd4='';}
				else{secBrd4='['+secBrd4+']';} 
				secSNme4=oRange.Cells(k,13);//설명글
				if(secSNme4=='.' || secSNme4.value == undefined || secSNme4.value == null || secSNme4.value == false){secSNme4='';}
				else{secSNme4='	<dd class="txt">'+secSNme4+'</dd>';};
				secPri4=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri4=oRange.Cells(k,15);//할인가
				secDPri4=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct41=round(oRange.Cells(k,24),0);//할인율
				secDct42=oRange.Cells(k,27);//그외값
				secico41=oRange.Cells(k,23);//아이콘1
				if(secico41=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct41+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico42=oRange.Cells(k,26);//아이콘2
				if(secico42=='ico00'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct42+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico01'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico42=='ico02'){
					icon4 = icon4 + 
									'		<table class="'+secico42+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico41=='ico03'){
					icon4 = icon4 + 
									'		<table class="'+secico41+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn4=oRange.Cells(k,29);//쿠폰율
				secLnk4=oRange.Cells(k,35);//URL
				secPic4=oRange.Cells(k,36);//대표이미지
				picResize(secPic4);
				secPic4=temp;

				k=k+1;
				//5번 컨텐츠
				secSone5=oRange.Cells(k,38);//이미지위치
				if(secSone5=='.' || secSone5.value == undefined || secSone5.value == null || secSone5.value == false || secSone5.value == 'NULL'){secSone5='';}
				secZoom5=oRange.Cells(k,37);//확대축소
				if(secZoom5=='.' || secZoom5.value == undefined || secZoom5.value == null || secZoom5.value == false || secZoom5.value == 'NULL'){secZoom5='';}
				secSold5=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold5=='.' || secSold5.value == undefined || secSold5.value == null || secSold5.value == false || secSold5.value == 'NULL'){secSold5='';}
				if(secSold5=="soldout"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold5=="upcoming"){secSold5='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold5='';}
				secNme5=oRange.Cells(k,11);//타이틀명
				secBrd5=oRange.Cells(k,12);//브랜드
				 if(secBrd5=='.' || secBrd5.value == undefined || secBrd5.value == null || secBrd5.value == false || secBrd5.value == 'NULL'){secBrd5='';}
				else{secBrd5='['+secBrd5+']';} 
				secSNme5=oRange.Cells(k,13);//설명글
				if(secSNme5=='.' || secSNme5.value == undefined || secSNme5.value == null || secSNme5.value == false){secSNme5='';}
				else{secSNme5='	<dd class="txt">'+secSNme5+'</dd>';};
				secPri5=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri5=oRange.Cells(k,15);//할인가
				secDPri5=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct51=round(oRange.Cells(k,24),0);//할인율
				secDct52=oRange.Cells(k,27);//그외값
				secico51=oRange.Cells(k,23);//아이콘1
				if(secico51=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct51+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico52=oRange.Cells(k,26);//아이콘2
				if(secico52=='ico00'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct52+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico01'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico52=='ico02'){
					icon5 = icon5 + 
									'		<table class="'+secico52+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico51=='ico03'){
					icon5 = icon5 + 
									'		<table class="'+secico51+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn5=oRange.Cells(k,29);//쿠폰율
				secLnk5=oRange.Cells(k,35);//URL
				secPic5=oRange.Cells(k,36);//대표이미지
				picResize(secPic5);
				secPic5=temp;

				k=k+1;
				//6번 컨텐츠
				secSone6=oRange.Cells(k,38);//이미지위치
				if(secSone6=='.' || secSone6.value == undefined || secSone6.value == null || secSone6.value == false || secSone6.value == 'NULL'){secSone6='';}
				secZoom6=oRange.Cells(k,37);//확대축소
				if(secZoom6=='.' || secZoom6.value == undefined || secZoom6.value == null || secZoom6.value == false || secZoom6.value == 'NULL'){secZoom6='';}
				secSold6=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold6=="soldout"){secSold6='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold6=="upcoming"){secSold6='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold6='';}
				secNme6=oRange.Cells(k,11);//타이틀명
				secBrd6=oRange.Cells(k,12);//브랜드
				 if(secBrd6=='.' || secBrd6.value == undefined || secBrd6.value == null || secBrd6.value == false || secBrd6.value == 'NULL'){secBrd6='';}
				else{secBrd6='['+secBrd6+']';} 
				secSNme6=oRange.Cells(k,13);//설명글
				if(secSNme6=='.' || secSNme6.value == undefined || secSNme6.value == null || secSNme6.value == false){secSNme6='';}
				else{secSNme6='	<dd class="txt">'+secSNme6+'</dd>';};
				secPri6=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri6=oRange.Cells(k,15);//할인가
				secDPri6=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct61=round(oRange.Cells(k,24),0);//할인율
				secDct62=oRange.Cells(k,27);//그외값
				secico61=oRange.Cells(k,23);//아이콘1
				if(secico61=='ico00'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct61+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico01'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico02'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico61=='ico03'){
					icon6 = icon6 + 
									'		<table class="'+secico61+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico62=oRange.Cells(k,26);//아이콘2
				if(secico62=='ico00'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct62+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico01'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico02'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico62=='ico03'){
					icon6 = icon6 + 
									'		<table class="'+secico62+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn6=oRange.Cells(k,29);//쿠폰율
				secLnk6=oRange.Cells(k,35);//URL
				secPic6=oRange.Cells(k,36);//대표이미지
				picResize(secPic6);
				secPic6=temp;

				k=k+1;
				//7번 컨텐츠
				secSone7=oRange.Cells(k,38);//이미지위치
				if(secSone7=='.' || secSone7.value == undefined || secSone7.value == null || secSone7.value == false || secSone7.value == 'NULL'){secSone7='';}
				secZoom7=oRange.Cells(k,37);//확대축소
				if(secZoom7=='.' || secZoom7.value == undefined || secZoom7.value == null || secZoom7.value == false || secZoom7.value == 'NULL'){secZoom7='';}
				secSold7=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold7=="soldout"){secSold7='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"></dd>'}
				else if(secSold7=="upcoming"){secSold7='<dd class="soldout"><img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"></dd>'}
				else{secSold7='';}
				secNme7=oRange.Cells(k,11);//타이틀명
				secBrd7=oRange.Cells(k,12);//브랜드
				 if(secBrd7=='.' || secBrd7.value == undefined || secBrd7.value == null || secBrd7.value == false || secBrd7.value == 'NULL'){secBrd7='';}
				else{secBrd7='['+secBrd7+']';} 
				secSNme7=oRange.Cells(k,13);//설명글
				if(secSNme7=='.' || secSNme7.value == undefined || secSNme7.value == null || secSNme7.value == false){secSNme7='';}
				else{secSNme7='	<dd class="txt">'+secSNme7+'</dd>';};
				secPri7=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri7=oRange.Cells(k,15);//할인가
				secDPri7=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct71=round(oRange.Cells(k,24),0);//할인율
				secDct72=oRange.Cells(k,27);//그외값
				secico71=oRange.Cells(k,23);//아이콘1
				if(secico71=='ico00'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct71+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico71=='ico01'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico71=='ico02'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico71=='ico03'){
					icon7 = icon7 + 
									'		<table class="'+secico71+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico72=oRange.Cells(k,26);//아이콘2
				if(secico72=='ico00'){
					icon7 = icon7 + 
									'		<table class="'+secico72+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct72+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico72=='ico01'){
					icon7 = icon7 + 
									'		<table class="'+secico72+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico72=='ico02'){
					icon7 = icon7 + 
									'		<table class="'+secico72+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico72=='ico03'){
					icon7 = icon7 + 
									'		<table class="'+secico72+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn7=oRange.Cells(k,29);//쿠폰율
				secLnk7=oRange.Cells(k,35);//URL
				secPic7=oRange.Cells(k,36);//대표이미지
				picResize(secPic7);
				secPic7=temp;

				k=k+1;
				//8번 컨텐츠
				secSone8=oRange.Cells(k,38);//이미지위치
				if(secSone8=='.' || secSone8.value == undefined || secSone8.value == null || secSone8.value == false || secSone8.value == 'NULL'){secSone8='';}
				secZoom8=oRange.Cells(k,37);//확대축소
				if(secZoom8=='.' || secZoom8.value == undefined || secZoom8.value == null || secZoom8.value == false || secZoom8.value == 'NULL'){secZoom8='';}
				secSold8=oRange.Cells(k,22);//솔드아웃,판매예정
				if(secSold8=="soldout"){secSold8='<dd class="soldout"><!-- <img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/soldout.png"> --></dd>'}
				else if(secSold8=="upcoming"){secSold8='<dd class="soldout"><!-- <img alt="" src="http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/upcoming.png"> --></dd>'}
				else{secSold8='';}
				secNme8=oRange.Cells(k,11);//타이틀명
				secBrd8=oRange.Cells(k,12);//브랜드
				 if(secBrd8=='.' || secBrd8.value == undefined || secBrd8.value == null || secBrd8.value == false || secBrd8.value == 'NULL'){secBrd8='';}
				else{secBrd8='['+secBrd8+']';} 
				secSNme8=oRange.Cells(k,13);//설명글
				if(secSNme8=='.' || secSNme8.value == undefined || secSNme8.value == null || secSNme8.value == false){secSNme8='';}
				else{secSNme8='	<dd class="txt">'+secSNme8+'</dd>';};
				secPri8=addComma(Number(oRange.Cells(k,14)));//정상가
				secDcPri8=oRange.Cells(k,15);//할인가
				secDPri8=addComma(Number(oRange.Cells(k,16)));//최종가
				secDct81=round(oRange.Cells(k,24),0);//할인율
				secDct82=oRange.Cells(k,27);//그외값
				secico81=oRange.Cells(k,23);//아이콘1
				if(secico81=='ico00'){
					icon8 = icon8 + 
									'		<table class="'+secico81+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct81+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico81=='ico01'){
					icon8 = icon8 + 
									'		<table class="'+secico81+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico81=='ico02'){
					icon8 = icon8 + 
									'		<table class="'+secico81+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico81=='ico03'){
					icon8 = icon8 + 
									'		<table class="'+secico81+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secico82=oRange.Cells(k,26);//아이콘2
				if(secico82=='ico00'){
					icon8 = icon8 + 
									'		<table class="'+secico82+'">'+
									'				<tr>'+
									'					<td><span><strong>'+secDct82+'</strong>%</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico82=='ico01'){
					icon8 = icon8 + 
									'		<table class="'+secico82+'">'+
									'				<tr>'+
									'					<td><span>무료<br/>배송</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico82=='ico02'){
					icon8 = icon8 + 
									'		<table class="'+secico82+'">'+
									'				<tr>'+
									'					<td><span>신상품</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else if(secico81=='ico03'){
					icon8 = icon8 + 
									'		<table class="'+secico81+'">'+
									'				<tr>'+
									'					<td><span>예약<br/>판매</span></td>'+
									'				</tr>'+
									'			</table>'+

									'';
				}
				else{};
				secCpn8=oRange.Cells(k,29);//쿠폰율
				secLnk8=oRange.Cells(k,35);//URL
				secPic8=oRange.Cells(k,36);//대표이미지
				picResize(secPic8);
				secPic8=temp;

				allCode = allCode+'<!-- // 가이드 영역 시작 부분 입니다.  -->'+
									'<div class="'+oRange.Cells(k,3)+'">'+
										'<div class="'+Type+'">'+
											'<dl>'+
											'	<dt class="pic '+secZoom1+' '+secSone1+'"><img src="'+secPic1+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold1+
											secSNme1+
											'	<dd class="ttl">'+'<span>'+secBrd1+'</span>'+secNme1+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri1+'</strike></span>'+
											'		<strong>'+secDPri1+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico11+'"><span><strong>'+secDct11+'</strong>%</span></span><span class="'+secico12+'">'+secDct12+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon1+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk1+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom2+' '+secSone2+'"><img src="'+secPic2+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold2+
											secSNme2+
											'	<dd class="ttl">'+'<span>'+secBrd2+'</span>'+secNme2+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri2+'</strike></span>'+
											'		<strong>'+secDPri2+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico21+'"><span><strong>'+secDct21+'</strong>%</span></span><span class="'+secico22+'">'+secDct22+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon2+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk2+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom3+' '+secSone3+'"><img src="'+secPic3+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold3+
											secSNme3+
											'	<dd class="ttl">'+'<span>'+secBrd3+'</span>'+secNme3+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri3+'</strike></span>'+
											'		<strong>'+secDPri3+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico31+'"><span><strong>'+secDct31+'</strong>%</span></span><span class="'+secico32+'">'+secDct32+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon3+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk3+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom4+' '+secSone4+'"><img src="'+secPic4+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold4+
											secSNme4+
											'	<dd class="ttl">'+'<span>'+secBrd4+'</span>'+secNme4+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri4+'</strike></span>'+
											'		<strong>'+secDPri4+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico41+'"><span><strong>'+secDct41+'</strong>%</span></span><span class="'+secico42+'">'+secDct42+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon4+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk4+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom5+' '+secSone5+'"><img src="'+secPic5+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold5+
											secSNme5+
											'	<dd class="ttl">'+'<span>'+secBrd5+'</span>'+secNme5+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri5+'</strike></span>'+
											'		<strong>'+secDPri5+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico51+'"><span><strong>'+secDct51+'</strong>%</span></span><span class="'+secico52+'">'+secDct52+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon5+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk5+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom6+' '+secSone6+'"><img src="'+secPic6+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold6+
											secSNme6+
											'	<dd class="ttl">'+'<span>'+secBrd6+'</span>'+secNme6+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri6+'</strike></span>'+
											'		<strong>'+secDPri6+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico61+'"><span><strong>'+secDct61+'</strong>%</span></span><span class="'+secico62+'">'+secDct62+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon6+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk6+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom7+' '+secSone7+'"><img src="'+secPic7+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold7+
											secSNme7+
											'	<dd class="ttl">'+'<span>'+secBrd7+'</span>'+secNme7+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri7+'</strike></span>'+
											'		<strong>'+secDPri7+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico71+'"><span><strong>'+secDct71+'</strong>%</span></span><span class="'+secico72+'">'+secDct72+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon7+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk7+'">링크</a></dd>'+
											'</dl>'+
											'<dl>'+
											'	<dt class="pic '+secZoom8+' '+secSone8+'"><img src="'+secPic8+'" onerror="this.src=\'http://mall.hyundailivart.co.kr/UserFiles/Image/01_promotion/common/error.jpg\';" alt=""/></dt>'+
											secSold8+
											secSNme8+
											'	<dd class="ttl">'+'<span>'+secBrd8+'</span>'+secNme8+'</dd>'+
											'	<dd class="price">'+
											'		<span><strike>'+secPri8+'</strike></span>'+
											'		<strong>'+secDPri8+'</strong>'+
											'	</dd>'+
											//'	<dd class="dc"><span class="'+secico81+'"><span><strong>'+secDct81+'</strong>%</span></span><span class="'+secico82+'">'+secDct82+'</span></dd>'+ 
											'	<dd class="icoD">'+
											icon8+
											'	</dd>'+
											'	<dd class="lnk"><a href="'+secLnk8+'">링크</a></dd>'+
											'</dl>'+
										'</div>'+				
									 '</div>';				
				icon1='';icon2='';icon3='';icon4='';icon5='';icon6='';icon7='';icon8='';
			}
			/* 테스트 */

			//alert(all.innerHTML);
		
			all.innerHTML = '<div class="promotion v2">'+allCode+'</div>';
		}

		

		


	/* } */
}


function nowLoc(i){
	if(i==1){
		loc = 'D:/Guide/ja';
	}
	else if(i==2){
		loc = 'D:/Guide/ja/';
		//loc = '\\\\172.17.14.220\\Lguidesystem\\ja\\';
	}

	return(loc);
}

//alert(clock());