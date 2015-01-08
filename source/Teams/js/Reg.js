var nametitle='�������û�����';
var profile_username_toolong = ' �����û������� 15 ���ַ���������һ���϶̵��û�����';
var profile_username_tooshort = ' ��������û���С��3���ַ�, ������һ���ϳ����û�����';
var profile_username_illegal = ' �û������������ַ���ϵͳ���Σ���������д��';
var keyuser = '����У����Ժ�' ;
var keynums = '�����������λ������3���ַ������������롣' ;
var pkeyook = '�����ʽ��ȷ�������' ;
var profile_passwd_illegal = ' ����ջ�����Ƿ��ַ�����������д��';
var twopassok = 'ȷ��������ȷ�������';
var profile_passwd_notmatch = ' ������������벻һ�£���������ԡ�';
var profile_email_illegal = ' ����д��ȷ�ĵ����ʼ���ַ ';
var emailok = '�����ʼ���ַ��ȷ�������' ;
var profile_seccode_invalid = ' ��֤��Ϊ��λ���������֣�����ȷ��д��';
var reg_1 = 0;	//�û���
var reg_2 = 0;	//����
var reg_3 = 0;	//ȷ������
var reg_4 = 0;	//�ʼ���ַ
var reg_5 = 0;	//��ȫ���ʺͻش�
var reg_6 = 0;	//��֤��

function trim(str) {
	return str.replace(/^\s*(.*?)[\s\n]*$/g, '$1');
}

function setfocus(a){
	$(a).className = "focus";
}

function setblur(a){
	$(a).className = "blur";
}

function checkfocus(s) {
	var form = $("tform");
	var username = trim(form.username.value);
	var unlen = username.replace(/[^\x00-\xff]/g, "**").length;
	if(unlen < 3 || unlen > 20) {
		$(s).innerHTML = (unlen < 3 ? profile_username_tooshort : profile_username_toolong);
		$(s).className = 'fall';
		return false;
	}
	var objname='ukey';
	var x = new Ajax('HTML', objname);
	x.get('ajax.asp?checksubmit=yes&action=checkusername&username='+username, function(s){
		var obj = $(objname);
		if (trim(s)=='num1'){
			obj.innerHTML = '��������û���['+username+']����ʹ�ã��������д����ѡ�';
			obj.className = 'true';
		} else if (trim(s)=='num2'){
			obj.innerHTML = '��������û���['+username+']������ѡ�������û�����';
			obj.className = 'fall';
			return false;
		} else if (trim(s)=='num3'){
			obj.innerHTML = '��������û���['+username+']���в�����ע����ַ�����ѡ�������û�����';
			obj.className = 'fall';
			return false;
		}
	});
	reg_1 = 1;

	//�����û������в�����ע����ַ������޸ĺ������ύ
}

function checkpass(s) {
	var form = $("tform");
	var password = trim(form.password.value);
	var unlen = password.replace(/[^\x00-\xff]/g, "**").length;
	if(unlen < 3 ) {
		$(s).innerHTML = keynums ;
		$(s).className = 'fall';
		return false;
	}
	if(password == '' || /[\'\"\\]/.test(password)) {
		$(s).innerHTML = profile_passwd_illegal;
		$(s).className = 'fall';
		return false;
	} else {
		$(s).innerHTML = pkeyook;
		$(s).className = 'true';
	}
	reg_2 = 1;
}

function checkpass1(s) {
	var form = $("tform");
	var password = trim(form.password.value);
	var password2 = trim(form.password2.value);
	if(password == '' || (s && password2 == '')) {
		$(s).className = 'fall';
		return false;
	}
	if(password != password2) {
		$(s).innerHTML = profile_passwd_notmatch ;
		$(s).className = 'fall';
		return false;
	} else {
		$(s).innerHTML = twopassok ;
		$(s).className = 'true';
	}
	reg_3 = 1;
}

function checkemail(s) {
	var form = $("tform");
	var email = trim(form.email.value);
	if(!emailValidate(email)) {
		$(s).innerHTML = profile_email_illegal ;
		$(s).className = 'fall';
		return false;
	} else {
		if (dmail==0){
			var objname='emailkey';
			var x = new Ajax('HTML', objname);
			x.get('ajax.asp?checksubmit=yes&action=checkemail&email='+email, function(s){
				var obj = $(objname);
				if (trim(s)=='true'){
					obj.innerHTML = '��������ʼ���ַ'+email+'����ʹ�á�';
					obj.className = 'true';
				} else {
					obj.innerHTML = '��������ʼ���ַ'+email+'������Ѿ��������û�ʹ�ã���ѡ�������û�����';
					obj.className = 'fall';
					return false;
				}
			});
		} else {
			$(s).innerHTML = emailok ;
			$(s).className = 'true';
		}
	}
	reg_4 = 1;
}

function checkquest(s) {
	var form = $("tform");
	var questionid = trim(form.questionid.value);
	var answer = trim(form.answer.value);
	if(questionid != '') {
		if (answer=='') {
			$(s).innerHTML = '����д��' ;
			$(s).className = 'fall';
			return false;
		}
		$(s).innerHTML = '���ס����д������ʹ𰸡�' ;
		$(s).className = 'true';
		$('questkey').className = 'true';
	} else {
		if (answer !=''){
			$(s).innerHTML = '��û����д��ȫ���ʡ�' ;
			$(s).className = 'fall';
			return false;
		}
	}
	reg_5 = 1;
}

function checkcode(s) {
	var form = $("tform");
	var code = trim(form.code.value);
	if (code ==''){
		$(s).innerHTML = '����д��֤�롣' ;
		$(s).className = 'fall';
		return false;
	};
	if (code.length!=4){
		$(s).innerHTML = profile_seccode_invalid ;
		$(s).className = 'fall';
		return false;	
	}
	$(s).innerHTML = '��֤��������ȷ���������';
	$(s).className = 'true';
	//�������ajax��֤���ᵼ��session���ڡ��ύ����֤����������
			
	//var objname='codekey';
	//var x = new Ajax('HTML', objname);
	//x.get("ajax.asp?checksubmit=yes&action=checkseccode&seccodeverify="+code, function(s){
		//var obj = $(objname);
		//if (trim(s)=='true'){
		//	obj.innerHTML = '��֤��������ȷ���������';
		//} else {
		//	obj.innerHTML = '����д��ȷ����֤�롣';
		//	return false;
		//}
	//});	
	reg_6 = 1;
}

function validate(theform) {
	checkfocus('ukey');
	checkpass('pkey');
	checkpass1('pkey1');
	checkemail('emailkey');
	checkquest('answerkey');
	checkcode('codekey');
	if(reg_1 == 0){
		return false;
	}else if(reg_2 == 0){
		return false;
	}else if(reg_3 == 0){
		return false;
	}else if(reg_4 ==0){
		return false;
	}else if(reg_5 ==0){
		return false;	
	}else if(reg_6 ==0){
		return false;
	}else{
		return true;
	}
}


function emailValidate(emailStr) {
	var checkTLD=1;
	var knownDomsPat=/^(com|net|org|edu|int|mil|gov|arpa|biz|aero|name|coop|info|pro|museum|mobi)$/;
	var emailPat=/^(.+)@(.+)$/;
	var specialChars="\\(\\)><@,;:\\\\\\\"\\.\\[\\]";
	var validChars="\[^\\s" + specialChars + "\]";
	var quotedUser="(\"[^\"]*\")";
	var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/;
	var atom=validChars + '+';
	var word="(" + atom + "|" + quotedUser + ")";
	var userPat=new RegExp("^" + word + "(\\." + word + ")*$");
	var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$");
	var matchArray=emailStr.match(emailPat);
	if (matchArray==null) {
		return false;
	}
	var user=matchArray[1];
	var domain=matchArray[2];
	for (i=0; i<user.length; i++) {
		if (user.charCodeAt(i)>127) {
			return false;
		}
	}
	for (i=0; i<domain.length; i++) {
		if (domain.charCodeAt(i)>127) {
			return false;
		}
	}
	if (user.match(userPat)==null) {
		return false;
	}
	var IPArray=domain.match(ipDomainPat);
	if (IPArray!=null) {
	// this is an IP address
		for (var i=1;i<=4;i++) {
			if (IPArray[i]>255) {
				return false;
			}
		}
		return true;
	}
	// Domain is symbolic name.  Check if it's valid.
	var atomPat=new RegExp("^" + atom + "$");
	var domArr=domain.split(".");
	var len=domArr.length;
	for (i=0;i<len;i++) {
		if (domArr[i].search(atomPat)==-1) {
			return false;
		}
	}
	if (checkTLD && domArr[domArr.length-1].length!=2 && domArr[domArr.length-1].search(knownDomsPat)==-1) {
		return false;
	}
	if (len<2) {
		return false;
	}
	return true;
}

function showadv() {
  if(document.tform.advshow.checked == true) {
      document.getElementById("adv").style.display = "";
   } else {
      document.getElementById("adv").style.display = "none";
  }
}