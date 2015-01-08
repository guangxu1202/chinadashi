var nametitle='请输入用户名。';
var profile_username_toolong = ' 您的用户名超过 15 个字符，请输入一个较短的用户名。';
var profile_username_tooshort = ' 您输入的用户名小于3个字符, 请输入一个较长的用户名。';
var profile_username_illegal = ' 用户名包含敏感字符或被系统屏蔽，请重新填写。';
var keyuser = '检测中，请稍后。' ;
var keynums = '您输入的密码位数少于3个字符，请重新输入。' ;
var pkeyook = '密码格式正确，请继续' ;
var profile_passwd_illegal = ' 密码空或包含非法字符，请重新填写。';
var twopassok = '确认密码正确，请继续';
var profile_passwd_notmatch = ' 两次输入的密码不一致，请检查后重试。';
var profile_email_illegal = ' 请填写正确的电子邮件地址 ';
var emailok = '电子邮件地址正确，请继续' ;
var profile_seccode_invalid = ' 验证码为四位阿拉伯数字，请正确填写。';
var reg_1 = 0;	//用户名
var reg_2 = 0;	//密码
var reg_3 = 0;	//确认密码
var reg_4 = 0;	//邮件地址
var reg_5 = 0;	//安全提问和回答
var reg_6 = 0;	//验证码

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
			obj.innerHTML = '您输入的用户名['+username+']可以使用，请继续填写其他选项。';
			obj.className = 'true';
		} else if (trim(s)=='num2'){
			obj.innerHTML = '您输入的用户名['+username+']错误，请选择其他用户名。';
			obj.className = 'fall';
			return false;
		} else if (trim(s)=='num3'){
			obj.innerHTML = '您输入的用户名['+username+']含有不允许注册的字符，请选择其他用户名。';
			obj.className = 'fall';
			return false;
		}
	});
	reg_1 = 1;

	//您的用户名含有不允许注册的字符，请修改后重新提交
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
					obj.innerHTML = '您输入的邮件地址'+email+'可以使用。';
					obj.className = 'true';
				} else {
					obj.innerHTML = '您输入的邮件地址'+email+'错误或已经被其他用户使用，请选择其他用户名。';
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
			$(s).innerHTML = '请填写答案' ;
			$(s).className = 'fall';
			return false;
		}
		$(s).innerHTML = '请记住您填写的问题和答案。' ;
		$(s).className = 'true';
		$('questkey').className = 'true';
	} else {
		if (answer !=''){
			$(s).innerHTML = '您没有填写安全提问。' ;
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
		$(s).innerHTML = '请填写验证码。' ;
		$(s).className = 'fall';
		return false;
	};
	if (code.length!=4){
		$(s).innerHTML = profile_seccode_invalid ;
		$(s).className = 'fall';
		return false;	
	}
	$(s).innerHTML = '验证码输入正确，请继续。';
	$(s).className = 'true';
	//如果开启ajax验证，会导致session过期。提交后验证码会产生错误。
			
	//var objname='codekey';
	//var x = new Ajax('HTML', objname);
	//x.get("ajax.asp?checksubmit=yes&action=checkseccode&seccodeverify="+code, function(s){
		//var obj = $(objname);
		//if (trim(s)=='true'){
		//	obj.innerHTML = '验证码输入正确，请继续。';
		//} else {
		//	obj.innerHTML = '请填写正确的验证码。';
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