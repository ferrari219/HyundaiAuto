//------------------------ ���� -------------------------------//
//�⺻����
var i; //��� ���� ����
var temp;
var txt;
var wSize;
var mall; //������
var scp; //CSS & font�� ����

//�����Լ�
function round(num, ja) {
	ja = Math.pow(10, ja);
	return Math.round(num * ja) / ja;
}
function ceil(num, ja) {
	ja = Math.pow(10, ja);
	return Math.ceil(num * ja) / ja;
}
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

function Mid(Str, Num1, Num2) {
	//alert(Num1);
	return String(Str).substr(Num1, Num2);
}

function addComma(n) {
	if (isNaN(n)) {
		return 0;
	}
	var reg = /(^[+-]?\d+)(\d{3})/;
	n += '';
	while (reg.test(n)) n = n.replace(reg, '$1' + ',' + '$2');
	return n;
}

function cntCng(k) {
	if (k < 10) {
		k = '0' + k;
	} else {
		k = k;
	}
	return k;
}

String.prototype.trim = function () {
	return this.replace(/\s/g, '');
};

//���ó�¥ �ڵ��Է�
window.onload = function () {
	now = new Date();
	var year = Right(now.getYear(), 2);
	var mon = now.getMonth() + 1;
	var day = now.getDate();

	//alert(monC);
	if (String(mon).length == 1) {
		mon = '0' + mon;
	}
	if (String(day).length == 1) {
		day = '0' + day;
	}

	//alert(year+mon+day);
	//return year+mon+day;
	document.frm.today.value = '161231_340_basic';
	//document.frm.today.value = document.frm.today.value;
	//document.frm.today.value = year+mon+day+'_��������';
	dayCode = document.frm.today.value;
};

/*HTML ����*/
function jsCopy(obj) {
	var copyObj = obj;
	var copyStr = document.getElementById(copyObj).innerHTML;

	window.clipboardData.setData('text', copyStr);
	alert('Html �ҽ��� ����Ǿ����ϴ�.');
}

// �̹��� ��ȯ
function picResize(secPic) {
	/* alert(secPic);*/
	temp = String(secPic);
	temp = temp.replace('img_95', 'img_615'); //ġȯ
	temp = temp.replace('img_140', 'img_615');
	temp = temp.replace('img_150', 'img_615');
	temp = temp.replace('img_170', 'img_615');
	temp = temp.replace('img_220', 'img_615');
	temp = temp.replace('_detail1_5', '_detail1_ORIGIN');
	temp = temp.replace('_detail2_5', '_detail2_ORIGIN');
	temp = temp.replace('_detail3_5', '_detail3_ORIGIN');
	temp = temp.replace('_detail4_5', '_detail4_ORIGIN');
	temp = temp.replace('_detail5_5', '_detail5_ORIGIN');

	temp = temp.replace('_detail1_4', '_detail1_ORIGIN');
	temp = temp.replace('_detail2_4', '_detail2_ORIGIN');
	temp = temp.replace('_detail3_4', '_detail3_ORIGIN');
	temp = temp.replace('_detail4_4', '_detail4_ORIGIN');
	temp = temp.replace('_detail5_4', '_detail5_ORIGIN');

	temp = temp.replace('_detail1_3', '_detail1_ORIGIN');
	temp = temp.replace('_detail2_3', '_detail2_ORIGIN');
	temp = temp.replace('_detail3_3', '_detail3_ORIGIN');
	temp = temp.replace('_detail4_3', '_detail4_ORIGIN');
	temp = temp.replace('_detail5_3', '_detail5_ORIGIN');

	temp = temp.replace('_detail1_2', '_detail1_ORIGIN');
	temp = temp.replace('_detail2_2', '_detail2_ORIGIN');
	temp = temp.replace('_detail3_2', '_detail3_ORIGIN');
	temp = temp.replace('_detail4_2', '_detail4_ORIGIN');
	temp = temp.replace('_detail5_2', '_detail5_ORIGIN');

	temp = temp.replace('_detail1_1', '_detail1_ORIGIN');
	temp = temp.replace('_detail2_1', '_detail2_ORIGIN');
	temp = temp.replace('_detail3_1', '_detail3_ORIGIN');
	temp = temp.replace('_detail4_1', '_detail4_ORIGIN');
	temp = temp.replace('_detail5_1', '_detail5_ORIGIN');

	temp = temp.replace('THUMB_1', 'ORIGIN');
	temp = temp.replace('THUMB_2', 'ORIGIN');
	temp = temp.replace('THUMB_3', 'ORIGIN');
	temp = temp.replace('THUMB_4', 'ORIGIN');
	temp = temp.replace('THUMB_5', 'ORIGIN');

	return temp;
}

// �ؽ�Ʈ ����
function txtCng(k) {
	/* alert(secPic);*/
	cng = String(k);
	//alert(cng);
	cng = cng.replace(/;/gi, '<br/>');
	cng = cng.replace(/_/gi, '&nbsp;&nbsp;&nbsp;&nbsp;');
	cng = cng.replace('[', '<span>');
	cng = cng.replace(']', '</span>');
	//cng = ""+cng;

	return cng;
}

//------------------------- ���� ���� --------------------------------//
//���� ��¥ �ݿ�
function todayCode() {
	dayCode = document.frm.today.value; //���³�¥ ��ǲ�� �ݿ�
	//alert(dayCode+'�� ã�ƺ��Կ�!');
}

//------------------------- ��Ÿ ���� --------------------------------//

//���� �ҷ�����
function excel(Location, i) {
	//������ �Էµ� Ÿ��, ��ȹ��Ÿ�Թ�ȣ

	todayCode();

	//alert(i);
	cTemp = 'border-left:1px dotted #ccc;border-right:1px dotted #ccc;border-bottom:1px dotted #ccc;';
	switch (i) {
		//����
		case '0':
			mall = 'cjmall';
			outPut = '_cj_deal';
			wSize = '758';
			break;
		case '1':
			mall = 'hmall';
			outPut = '_h_deal';
			wSize = '778';
			break;
		case '2':
			mall = 'gsshop';
			outPut = '_gs_deal';
			wSize = '778';
			break;
		case '3':
			mall = 'lottei';
			outPut = '_lottei_deal';
			wSize = '778';
			break;
		case '4':
			mall = 'naver';
			outPut = '_naver_deal';
			wSize = '758';
			break;
		case '5':
			mall = '11st';
			outPut = '_e_deal';
			wSize = '858';
			break; //11����
		case '6':
			mall = 'auction';
			outPut = '_a_deal';
			wSize = '858';
			break; //����
		case '7':
			mall = 'gmarket';
			outPut = '_g_deal';
			wSize = '858';
			break; //G����
		case '8':
			mall = 'g9';
			outPut = '_g9_deal';
			wSize = '778';
			break;
		case '9':
			mall = 'tmon';
			outPut = '_tmon_deal';
			wSize = '768';
			break;
		case '99':
			mall = '';
			outPut = '';
			wSize = '758';
			break; //etc
		default:
			alert('�ٽ� ���� ���ּ���.');
			break;
	}
	//alert(wSize);

	//������ ���� ����
	var Excel, oRange;

	// ���� ��ü����
	Excel = new ActiveXObject('Excel.Application');
	Excel.Application.Workbooks.Open('\\\\100.10.10.200\\Lguidesystem\\ip\\' + dayCode + outPut + '.xls');
	//Excel.Application.Workbooks.Open('C:/Guide/ip/'+dayCode+outPut+'.xls');
	//Excel.Application.Workbooks.Open('http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/excel/'+dayCode+outPut+'.xls');
	//Excel.Application.Workbooks.Open('http://jpub.cafe24.com/G_v02/excel/ip/'+dayCode+outPut+'.xls');

	//����ǥ��
	Excel.Application.Visible = true;
	//sheet1 ����
	Excel.Application.Worksheets('sheet1').Activate;
	// sheet1�� �۾������� ��ü�� ����
	oRange = Excel.Application.ActiveSheet.UsedRange; //oRange���� ���� ������ �ִ� �κи� ���õ�.
	//oRange�κ��� ����ǥ��
	oRange.Select();

	//oRange�κ��� �������� ���
	var Type;
	var all = document.getElementById('dev_all'); //ã��
	all.innerHTML = ''; //�ʱ�ȭ

	var sum = ''; //�հ谪
	var oBox = '';
	var oBoxC = ''; //�ɼǿ���
	var oSec = '';
	var oSecC = ''; //�ɼǱ��п���
	var tBox = '';
	var tBoxC = ''; //Ÿ��Ʋ����
	var tSec = '';
	var tSecC = ''; //Ÿ��Ʋ���п���
	var cnt = 0;
	var cntT; //ī��Ʈ ����
	var sNum; // ��ȣ
	var ttl; //Ÿ��Ʋ
	var pic; //����
	var dc; //����
	var pc; //����
	var pct1;
	var pct2;
	var pct3; //���ݸ�
	var pcn1;
	var pcn2;
	var pcn3; //����
	var pc1;
	var pc2;
	var pc3; //���� ��ü��
	var ee; //������
	var tt; //Ÿ��Ʋ��
	var ss; //�ɼǰ�
	var stNum; //Ÿ��Ʋ��ȣ
	var tTtl; //Ÿ��Ʋ
	var tpic; //����
	var tdc; //����
	var tpc; //����
	var tpct1;
	var tpct2;
	var tpct3; //���ݸ�
	var tpcn1;
	var tpcn2;
	var tpcn3; //����
	var tpc1;
	var tpc2;
	var tpc3; //���� ��ü��
	var part1;
	var part2;
	var part3;
	var tpart1;
	var tpart2;
	var tpart3;
	var emt = ''; //�󿵿� ��

	var ex = new Array();

	//�� �ޱ�
	for (var e = 12; e <= oRange.Rows.count; e++) {
		ex[e] = new Array(); //excel�� �迭 ����
		for (var j = 1; j <= oRange.Columns.count; j++) {
			ex[e][j] = oRange.Cells(e, j);

			/*unde���� �� �⺻�� ����*/
			if (ex[e][j] == '.' || ex[e][j].value == undefined || ex[e][j].value == null || ex[e][j].value == false) {
				switch (j) {
					case 2:
						ex[e][j] = 'd0';
						break; //������
					case 4:
						ex[e][j] = 's2';
						break; //��ǰ ���� ���
					case 20:
						ex[e][j] = 's2';
						break; //����ǥ��
					case 36:
						ex[e][j] = 'h0';
						break; //�̹��� Ȯ��
					case 37:
						ex[e][j] = 'h2';
						break; //�ɼ� ����
					case 38:
						ex[e][j] = 'h1';
						break; //Ÿ��Ʋ ����
					default:
						ex[e][j] = '';
						break;
				} /*switch (j) */
			} /*unde���� �� �⺻�� ����*/
			//sum=sum+ex[e][j]+' ';

			if (e >= 13) {
				/* �� �Ҵ� */
				if (j == 59) {
					if (ex[e][11] == '') {
						ee = '';
					} //������ �籸��
					else if (ex[e][11] == '@') {
						ee = '				<p class="e emt"></p>';
					} else {
						ee = '				<p class="e">' + txtCng(ex[e][11]) + '</p>';
					}
					if (ex[e][15] == '') {
						tt = '';
					} //Ÿ��Ʋ �籸��
					else {
						tt = '				<p class="t">' + txtCng(ex[e][15]) + '</p>';
					}
					if (ex[e][16] == '') {
						ss = '';
					} //�ɼ� �籸��
					else {
						ss = '				<p class="s">' + txtCng(ex[e][16]) + '</p>';
					}

					if (ex[e][21] == '') {
						pct1 = '';
					} //���ݸ�1 �籸��
					else {
						pct1 = '				<span>' + ex[e][21] + '</span>';
					}
					if (ex[e][22] == '') {
						pcn1 = '';
					} //����1 �籸��
					else {
						pcn1 = '				<em>' + addComma(ex[e][22]) + '<span>��</span></em>';
					}

					if (ex[e][23] == '') {
						pct2 = '';
					} //���ݸ�2 �籸��
					else {
						pct2 = '				<span>' + ex[e][23] + '</span>';
					}
					if (ex[e][24] == '') {
						pcn2 = '';
					} //����2 �籸��
					else {
						pcn2 = '				<em>' + addComma(ex[e][24]) + '<span>��</span></em>';
					}

					if (ex[e][25] == '') {
						pct3 = '';
					} //���ݸ�3 �籸��
					else {
						pct3 = '				<span>' + ex[e][25] + '</span>';
					}
					if (ex[e][26] == '') {
						pcn3 = '';
					} //����3 �籸��
					else {
						pcn3 = '				<em>' + addComma(ex[e][26]) + '<span>��</span></em>';
					}

					ex[e][32] = round(ex[e][32], 2); //

					if (pct1 == '' && pcn1 == '') {
						pc1 = '';
					} else if (String(ex[e][2]) == 'd0' || String(ex[e][2]) == 'tmon0') {
						//�⺻���� ��� reverse
						pc1 = '' + '				<div class="pc1">' + pcn1 + pct1 + '				</div>';
					} else {
						pc1 = '' + '				<div class="pc1">' + pct1 + pcn1 + '				</div>';
					}
					if (pct2 == '' && pcn2 == '') {
						pc2 = '';
					} else if (String(ex[e][2]) == 'd0' || String(ex[e][2]) == 'tmon0') {
						//�⺻���� ��� reverse
						pc2 = '' + '				<div class="pc2">' + pcn2 + pct2 + '				</div>';
					} else {
						pc2 = '' + '				<div class="pc2">' + pct2 + pcn2 + '				</div>';
					}
					if (pct3 == '' && pcn3 == '') {
						pc3 = '';
					} else if (String(ex[e][2]) == 'd0' || String(ex[e][2]) == 'tmon0') {
						//�⺻���� ��� reverse
						pc3 = '' + '				<div class="pc3">' + pcn3 + pct3 + '				</div>';
					} else {
						pc3 = '' + '				<div class="pc3">' + pct3 + pcn3 + '				</div>';
					}
					if (ex[e][27] == '') {
						etc = '';
					} else {
						etc = '' + '				<div class="etc">' + ex[e][27] + '				</div>';
					}

					sNum = '' + '<!-- ���ڿ��� sNum -->' + '<div class="sNum ' + ex[e][37] + '">' + '	<table class="typeA">' + '		<tr>' + '			<td>' + '				<span>' + ex[e][9] + '</span>' + '				<strong>' + ex[e][10] + '</strong>' + '			</td>' + '		</tr>' + '	</table>' + '</div>' + '';
					ttl = '' + '<!-- Ÿ��Ʋ���� -->' + '<div class="ttl ' + ex[e][37] + '">' + '	<table class="typeA">' + '		<tr>' + '			<td>' + ee + tt + ss + '			</td>' + '		</tr>' + '	</table>' + '</div>' + '';
					if (ex[e][31] == '' && ex[e][32] == '' && ex[e][33] == '') {
						dc = '';
					} else {
						dc =
							'' +
							'<!-- ���ο��� -->' +
							'<div class="dc">' +
							'	<table class="typeA">' +
							'		<tr>' +
							'			<td>' +
							'				<em>' +
							ex[e][31] +
							'</em>' + //�����ܸ�
							'				<strong>' +
							ex[e][32] +
							'</strong>' + //��ġ
							'				<span>' +
							ex[e][33] +
							'</span>' + //����
							'			</td>' +
							'		</tr>' +
							'	</table>' +
							'</div>' +
							'';
					}
					pic = '' + '<!-- �������� -->' + '<div class="pic ' + ex[e][36] + '">' + '	<p><img src="' + picResize(ex[e][34]) + '" alt="" /></p>' + '</div>' + '';
					pc = '' + '<!-- ���ݿ��� -->' + '<div class="prc">' + '	<table class="typeA">' + '		<tr>' + '			<td>' + '				<!-- ���ݿ��� -->' + pc1 + pc2 + pc3 + etc + '			</td>' + '		</tr>' + '	</table>' + '</div>' + '';

					/* Ÿ��Ʋ �� ���� �Ҵ� */
					if (ex[e][9] == '@' || ex[e][10] == '@') {
						sNum = '' + '<!-- ���ڿ��� sNum -->' + '<div class="sNum emt ' + ex[e][37] + '">' + '</div>' + '';
						stNum = '' + '<div class="sNum emt  ' + ex[e][38] + '">' + '</div>' + '';
					} else if (ex[e][9] == '' && ex[e][10] == '') {
						sNum = '' + '<!-- ���ڿ��� sNum -->' + '<div class="sNum emt ' + ex[e][37] + '">' + '</div>' + '';
						stNum =
							'' +
							//'<div class="sNum emt  '+ex[e][38]+'">'+
							//'</div>'+
							'';
					} else {
						stNum = '' + '<!-- �ѹ� ���� -->' + '<div class="sNum ' + ex[e][38] + '">' + '	<table class="typeA">' + '		<tr>' + '			<td>' + '				<em>' + ex[e][9] + '</em><span>' + ex[e][10] + '</span>' + '			</td>' + '		</tr>' + '	</table>' + '</div>' + '';
					}
					tTtl = '' + '<div class="ttl ' + ex[e][37] + '">' + '	<table class="typeA">' + '		<tr>' + '			<td>' + ee + tt + ss + '			</td>' + '		</tr>' + '	</table>' + '</div>' + '';
					if (ex[e][22] == '' && ex[e][24] == '' && ex[e][26] == '') {
						pc = '';
						tpc = '';
					} else {
						tpc = '' + '<div class="prc ' + ex[e][20] + ' ' + ex[e][38] + '">' + '	<table class="typeA">' + '		<tr>' + '			<td>' + '				<!-- ���ݿ��� -->' + pc1 + pc2 + pc3 + etc + '			</td>' + '		</tr>' + '	</table>' + '</div>' + '';
					}
				} //�� �Ҵ�
				/* ���� */
				if (j == 60) {
					/* Css & Font */
					if (e == 13) {
						var css; //CSS ���� ����
						if (ex[e][2] == 'l1' || ex[e][2] == 'l2') {
							css = 'l0';
						} else {
							css = ex[e][2];
						}
					} // e���� 12�϶� �ѹ��� ����
					scp = '<link rel="stylesheet" type="text/css" href="css/' + css + '.css"/> '; //�߰� ��ũ��Ʈ��

					//alert(ex[e][2]);
					switch (String(ex[e][2])) {
						case 'gs0':
							scp = scp + '<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>' + '<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>' + '<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //�߰� ��ũ��Ʈ��
							part1 = '' + '<div class="part1">' + sNum + ttl + '</div>' + '';
							part2 = '' + '<div class="part2">' + pic + dc + '</div>' + '';
							part3 = '' + '<div class="part3">' + pc + '</div>' + '';
							tpart1 = '' + '<div class="part1">' + stNum + '</div>' + '';
							tpart2 = '' + '<div class="part2">' + tTtl + '</div>' + '';
							tpart3 = '' + '';
							break;
						case 'gs1':
							scp = scp + '<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>' + '<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>' + '<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //�߰� ��ũ��Ʈ��
							part1 = '' + '<div class="part1">' + sNum + ttl + '</div>' + '';
							part2 = '' + '<div class="part2">' + pic + dc + '</div>' + '';
							part3 = '' + '<div class="part3">' + pc + '</div>' + '';
							tpart1 = '' + '<div class="part1">' + stNum + '</div>' + '';
							tpart2 = '' + '<div class="part2">' + tTtl + '</div>' + '';
							tpart3 = '' + '';
							break;
						case 'l0':
							scp = scp + '<link href="https://cdn.rawgit.com/theeluwin/NotoSansKR-Hestia/master/stylesheets/NotoSansKR-Hestia.css" rel="stylesheet" type="text/css"/>' + ''; //�߰� ��ũ��Ʈ��
							part1 = '' + '<div class="part1">' + sNum + pic + '</div>' + '';
							part2 = '' + '<div class="part2">' + ttl + dc + '</div>' + '';
							part3 = '' + '<div class="part3">' + pc + '</div>' + '';
							tpart1 = '' + '<div class="part1">' + stNum + '</div>' + '';
							tpart2 = '' + '<div class="part2">' + tTtl + '</div>' + '';
							tpart3 = '' + '<div class="part3">' + tpc + '</div>' + '';
							break;
						case 'a0':
							scp = scp + '<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>' + '<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //�߰� ��ũ��Ʈ��
							part1 = '' + '<div class="part1">' + sNum + ttl + '</div>' + '';
							part2 = '' + '<div class="part2">' + pc + dc + '</div>' + '';
							part3 = '' + '<div class="part3">' + pic + '</div>' + '';
							tpart1 = '' + '<div class="part1">' + stNum + '</div>' + '';
							tpart2 = '' + '<div class="part2">' + tTtl + '</div>' + '';
							tpart3 = '' + '<div class="part3">' + tpc + dc + '</div>' + '';
							break;
						case 'tmon0':
							scp = scp + '<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>' + '<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>' + '<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //�߰� ��ũ��Ʈ��
							part1 = '' + '<div class="part1">' + sNum + ttl + '</div>' + '';
							part2 = '' + '<div class="part2">' + pic + '</div>' + '';
							part3 = '' + '<div class="part3">' + dc + pc + '</div>' + '';
							tpart1 = '' + '<div class="part1">' + stNum + '</div>' + '';
							tpart2 = '' + '<div class="part2">' + tTtl + '</div>' + '';
							tpart3 = '' + '<div class="part3">' + tpc + '</div>' + '';
							break;
						default:
							scp = scp + '<link rel="stylesheet" type="text/css" href="css/NanumBarunGothic.css"/>' + '<link rel="stylesheet" type="text/css" href="css/whitney-Bold.css"/>' + '<link rel="stylesheet" type="text/css" href="css/OldSansBlack.css"/>'; //�߰� ��ũ��Ʈ��
							part1 = '' + '<div class="part1">' + sNum + ttl + '</div>' + '';
							part2 = '' + '<div class="part2">' + pic + '</div>' + '';
							part3 = '' + '<div class="part3">' + dc + pc + '</div>' + '';
							tpart1 = '' + '<div class="part1">' + stNum + '</div>' + '';
							tpart2 = '' + '<div class="part2">' + tTtl + '</div>' + '';
							tpart3 = '' + '<div class="part3">' + tpc + '</div>' + '';
							break;
					}
					oSec = oSec + '<div class="oSec">' + '	<a href="' + ex[e][35] + '" target="_blank">' + part1 + part2 + part3 + '	</a>' + '</div>' + '';

					tSec =
						'' +
						'<!-- Ÿ��Ʋ ���� -->' +
						'<div class="tSec">' +
						tpart1 +
						tpart2 +
						tpart3 +
						'</div>' +
						//'<p style="clear:both;margin:0;padding:0;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/03_ipjum/'+ex[e][50]+'.jpg" style="display:block;margin:0 auto;" border="0" alt=""/></p>'+
						'';

					if (String(e - 12).length == 1) {
						cntT = '0' + String(e - 12);
					} else {
						cntT = String(e - 12);
					}
					tSecC =
						'' +
						'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/' +
						mall +
						'/' +
						dayCode +
						'_deal/guide/t' +
						cntT +
						'.jpg" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0"></p>' +
						'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/' +
						mall +
						'/' +
						dayCode +
						'_deal/b' +
						cntT +
						'.jpg" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0"></p>' +
						'';

					if (ex[13][4] == 's2' && e % 2 == 0) {
						//¦���϶��� ����
						oBox = oBox + '<!-- �ɼǸ���Ʈ ���� -->' + '<div class="oBox ' + ex[13][4] + '">' + oSec + '</div>' + '';
						oSec = ''; //���а� ���� �ʱ�ȭ
						cnt = cnt + 1; //ī��Ʈ
						//�ڸ��� ����
						if (String(cnt).length == 1) {
							cntC = '0' + String(cnt);
						} else {
							cntC = cnt;
						}
						oSecC =
							oSecC +
							'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/' +
							mall +
							'/' +
							dayCode +
							'_deal/guide/o' +
							cntC +
							'.jpg" usemap="#lvtMap' +
							cntC +
							'" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0">' +
							'	<map name="lvtMap' +
							cntC +
							'" id="lvtMap' +
							cntC +
							'">' +
							//'	  <area shape="rect" coords="0,0,387,648" href='+ex[e-1][35]+' title="'+ex[e-1][15]+'" target="_blank" />'+
							//'	  <area shape="rect" coords="392,4,777,645" href='+ex[e][35]+' title="'+ex[e][15]+'" target="_blank" />'+
							'	</map>' +
							'</p>' +
							'';
					} //if ¦�����϶��� �ݴ´�
					else if (e == oRange.Rows.count && e % 2 == 1) {
						//Ȧ����ó��
						oSec = '' + '<div class="oSec">' + '	<a href="' + ex[e][35] + '" target="_blank">' + part1 + part2 + part3 + '	</a>' + '</div>' + '';
						emt =
							'<div class="oSec empty">' +
							'	<table class="typeA">' +
							'		<tr>' +
							'			<td>' +
							'				<img src="http://mall.hyundailivart.co.kr/image/mobile/renewal/logo.png" alt="���븮��Ʈ"/>' +
							'			</td>' +
							'		</tr>' +
							'	</table>' +
							//'	<p>����پƾ�</p>'+
							'</div>';
						emt = '' + '<!-- �ɼǸ���Ʈ ���� -->' + '<div class="oBox ' + ex[13][4] + '">' + oSec + emt + '</div>' + '';
						oSec = '';
					}

					if (ex[13][4] == 's2' && e == oRange.Rows.count && e % 2 == 1) {
						cntC = cnt + 1;
						if (String(cntC).length == 1) {
							cntC = '0' + String(cntC);
						}
						emtC =
							'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/' +
							mall +
							'/' +
							dayCode +
							'_deal/guide/o' +
							cntC +
							'.jpg" usemap="#lvtMap' +
							cntC +
							'" style="margin: 0px auto; border: 0px currentColor; border-image: none; display: block;" border="0">' +
							'	<map name="lvtMap' +
							cntC +
							'" id="lvtMap' +
							cntC +
							'">' +
							//'	  <area shape="rect" coords="0,0,387,648"  href='+ex[e-1][35]+' title="'+ex[e-1][15]+'" target="_blank" />'+ //��ũ��
							'	</map>' +
							'</p>' +
							'';
					} else {
						emtC = '';
					} //if emtC �� ����

					tBox = tBox + tSec;

					oBoxC = oSecC + emtC;
					tBoxC = tBoxC + tSecC;
				} //j40
			} //if e>=13
		} //for j
	} //for e //for e

	//tBox='';
	/* �� �籸�� */
	/*for(var e=10;e<=oRange.Rows.count;e++){ 
		for(var j=1;j<=oRange.Columns.count;j++){ 
			
		};//for j
	};*/ /*���� ����� ����*/
	sum =
		'' +
		scp + //Css, Font ȣ�Ⱚ
		'<div class="oWrap" style="width:' +
		wSize +
		'px;">' +
		oBox +
		emt +
		'</div>' +
		'<!-- Ÿ��Ʋ ���� -->' +
		'<div class="tBox" style="width:' +
		wSize +
		'px;">' +
		tBox +
		'</div>' +
		'<div id="dev_allC" style="display:none;">' + //
		'	<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/03_ipjum/' +
		mall +
		'/' +
		dayCode +
		'_deal/a01.jpg" style="display:block;margin:0 auto;" border="0"/></p>' +
		oBoxC +
		tBoxC +
		'<p style="margin: 0px; padding: 0px;"><img src="http://mall.hyundailivart.co.kr/UserFiles/Image/00_product/01_online/01_ismine/01_common/99_ship/01.jpg" style="display:block;margin:0 auto;" border="0"/></p>' + // �������
		'</div>' +
		//'test'+
		'';
	//alert(sum);
	/*���� ����� ����*/

	all.innerHTML = sum;
}
