/******************************************************************************
  team Board - modify for Team's daymoon
  Copyright 2005-2006 team studio. (http://www.team5.cn)
*******************************************************************************/
var sPop = null;
var postSubmited = false;
var userAgent = navigator.userAgent.toLowerCase();
var is_ie = (userAgent.indexOf('msie') != -1 && !is_opera && !is_saf && !is_webtv) && userAgent.substr(userAgent.indexOf('msie') + 5, 3);
var is_moz = (navigator.product == 'Gecko' && !is_saf) && userAgent.substr(userAgent.indexOf('firefox') + 8, 3);
var is_opera = userAgent.indexOf('opera') != -1 && opera.version();
var is_saf = userAgent.indexOf('applewebkit') != -1 || navigator.vendor == 'Apple Computer, Inc.';
var is_webtv = userAgent.indexOf('webtv') != -1;
var is_kon = userAgent.indexOf('konqueror') != -1;
var is_ns = userAgent.indexOf('compatible') == -1 && userAgent.indexOf('mozilla') != -1 && !is_opera && !is_webtv && !is_saf;
var is_mac = userAgent.indexOf('mac') != -1;

// ==menu 's == //
var menuOffX=0	//菜单距连接文字最左端距离
var menuOffY=20	//菜单距连接文字顶端距离
function showmenu(e,vmenu,mod){
	which=vmenu
	menuobj=document.getElementById("popmenu")
	menuobj.thestyle=menuobj.style
	menuobj.innerHTML=which
	menuobj.contentwidth=menuobj.offsetWidth
	eventX=e.clientX
	eventY=e.clientY
	var rightedge=document.body.clientWidth-eventX
	var bottomedge=document.body.clientHeight-eventY
	if (rightedge<menuobj.contentwidth)
		{
		menuobj.thestyle.left=document.body.scrollLeft+eventX-menuobj.contentwidth+menuOffX
		}
	else
		{
		menuobj.thestyle.left=is_ie? ie_x(event.srcElement)+menuOffX : is_ns? window.pageXOffset+eventX : eventX
		}
	if (bottomedge<menuobj.contentheight&&mod!=0)
		{
		menuobj.thestyle.top=document.body.scrollTop+eventY-menuobj.contentheight-event.offsetY+menuOffY-23
		}
	else
		{
		menuobj.thestyle.top=is_ie? ie_y(event.srcElement)+menuOffY : is_ns? window.pageYOffset+eventY+10 : eventY
		}
	menuobj.thestyle.visibility="visible"
}

function ie_y(e){  
	var t=e.offsetTop;  
	while(e=e.offsetParent){  
		t+=e.offsetTop;  
	}  
	return t;  
}  
function ie_x(e){  
	var l=e.offsetLeft;  
	while(e=e.offsetParent){  
		l+=e.offsetLeft;  
	}  
	return l;  
}

function highlightmenu(e,state){
	if (document.all)
		source_el=event.srcElement
		while(source_el.id!="popmenu"){
			source_el=document.getElementById? source_el.parentNode : source_el.parentElement
			if (source_el.className=="menuitems"){
				source_el.id=(state=="on")? "mouseoverstyle" : ""
		}
	}
}

function InsertHTML(s){
	var oEditor = FCKeditorAPI.GetInstance('message') ;
	var strHtml = '';
	var iTitle = prompt('标签内容', "需要显示在标签内的内容");
	switch ( s ){
		case 1 : strHtml = '[code]'+iTitle+'[/code]' ; break ;
		case 2 : strHtml = '[REPLAYVIEW]'+iTitle+'[/REPLAYVIEW]' ; break ;
		case 3 : strHtml = '[QQ]'+iTitle+'[/QQ]' ; break ;
		case 4 : strHtml = '[Quote]'+iTitle+'[/Quote]' ; break ;
		case 5 : {
			var ibuy = prompt('请输入购买金额', "10");
			strHtml = '[buy='+ibuy+']'+iTitle+'[/buy]' ; break ;
		}
		case 6 : strHtml = '[Qvod]'+iTitle+'[/Qvod]' ; break ;
	}
	oEditor.InsertHtml( strHtml ) ;
}


function hidemenu(){if (window.menuobj)menuobj.thestyle.visibility="hidden"}
function dynamichide(e){if ((is_ie||is_ns)&&!menuobj.contains(e.toElement))hidemenu()}
document.onclick=hidemenu
document.write("<div class=menuskin id=popmenu onmouseover=highlightmenu(event,'on') onmouseout=highlightmenu(event,'off');dynamichide(event)></div>")
// 菜单END

function getpass(s,b){
	if (unescape(s)==readCookie("ipass"+ b )){
		writeCookie("ulookpass"+ b,"ok","720")
		alert("密码输入正确,请刷新后查看内容")
	}else {
		alert("密码输入错误,请重新输入");
	}
	//alert (readCookie("ipass"+ b ));
}

function readCookie(name) {
	    var cookieValue = "";
	    var search = name + "=";
	    if(document.cookie.length > 0) { 
		    offset = document.cookie.indexOf(search);
		    if (offset != -1) { 
			    offset += search.length;
			    end = document.cookie.indexOf(";", offset);
			    if (end == -1) end = document.cookie.length;
			        cookieValue = unescape(document.cookie.substring(offset, end))
	        }
	    }
	    return cookieValue;
	}

 function writeCookie(name, value, hours) {
	    var expire = "";
	    if(hours != null) {
		    expire = new Date((new Date()).getTime() + hours * 3600000);
		    expire = "; expires=" + expire.toGMTString();
	    }
	    document.cookie = name + "=" + escape(value) + expire;
}


function countDown(secs){
	$("stime").innerText = secs;
	if(--secs>=0) setTimeout("countDown("+secs+")",1000);
}

function countup(secs){
	$("stime").innerText = secs;
	if(++secs>=0) setTimeout("countup("+secs+")",1000);
}

//拷贝地址
function copyurls(s){
    var clipBoardContent= document.title.split("|")[0]+"\n"+location.href  //定义变量内容
    window.clipboardData.setData("Text",clipBoardContent);  //赋值
    alert("复制本站链接成功!");  //弹出提示
}

//改变字体大小
function doZoom(size,name){
	document.getElementById('items_'+ name).style.fontSize=size+'px';
}

function cloneObj(oClone, oParent, count) {
	if(oParent.childNodes.length < count) {
		var newNode = oClone.cloneNode(true);
		oParent.appendChild(newNode);
		
		return true;
	} 
	return false;	
}

function delObj(oParent, count) {
	if(oParent.childNodes.length > count) {
		oParent.removeChild(oParent.lastChild);
		return true;
	}
	return false;
}

function Insertcccode(htmlstr){  
	var oEditor = FCKeditorAPI.GetInstance('message') ;
 	oEditor.InsertHtml( htmlstr );
}

function setfileid(maxup){
	str='';
	if(!document.getElementById("upcount").value){
		document.getElementById("upcount").value=1;
	}
	if(document.getElementById("upcount").value > maxup){
		alert('您最多只能同时上传 '+maxup+' 个文件!');
		document.getElementById("upcount").value = maxup;
		setfileid();
	}
	else
	{
		for(i=1;i<=document.getElementById("upcount").value;i++)
			str+='<div id="divfileItem" style="padding-top:4px"> 上传附件: <input type="file" name="fileitemid'+i+'" size="60" onBlur=this.className="colorblur"; onfocus=this.className="colorfocus"; class="colorblur"></div>';
			document.getElementById("divfileItem").innerHTML=str;
		}
}

function clonePoll(maxpoll){
	if(!cloneObj(document.getElementById('divPollItem'), document.getElementById('polloptions') ,maxpoll)){
		alert('投票项不能多于 ' + maxpoll + ' 个');
	}
	document.all("pollitemid")[document.all('pollitemid').length-1].value = "";
}

// ==msg 's== //
function checkclick(msg){
	if(confirm(msg))
		{
		event.returnValue=true;}else{event.returnValue=false;
		}
	}

// ==form 's== //
function checkall(form, prefix, checkall) {
	var checkall = checkall ? checkall : 'chkall';
	for(var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if(e.name != checkall && (!prefix || (prefix && e.name.match(prefix))))
		{
		e.checked = form.elements[checkall].checked;
		}
	}
}

function copycode(obj) {
	var rng = document.body.createTextRange();
	rng.moveToElementText(obj);
	rng.scrollIntoView();
	rng.select();
	rng.execCommand("Copy");
	rng.collapse(false);
}

function toggle_collapse(objname, unfolded) {
	if(typeof unfolded == 'undefined') {
		var unfolded = 1;
	}
	var obj = $(objname);
	var oldstatus = obj.style.display;
	var collapsed = getcookie('team_collapse');
	var cookie_start = collapsed ? collapsed.indexOf(objname) : -1;
	var cookie_end = cookie_start + objname.length + 1;

	obj.style.display = oldstatus == 'none' ? '' : 'none';
	collapsed = cookie_start != -1 && ((unfolded && oldstatus == 'none') || (!unfolded && oldstatus == '')) ?
			collapsed.substring(0, cookie_start) + collapsed.substring(cookie_end, collapsed.length) : (
			cookie_start == -1 && ((unfolded && oldstatus == '') || (!unfolded && oldstatus == 'none')) ?
			collapsed + objname + ' ' : collapsed);

	expires = new Date();
	expires.setTime(expires.getTime() + (collapsed ? 86400 * 30 : -(86400 * 30 * 1000)));
	document.cookie = 'team_collapse=' + escape(collapsed) + '; expires=' + expires.toGMTString() + '; path=/';

	var img = $(objname + '_img');
	var img_regexp = new RegExp((oldstatus == 'none' ? '_yes' : '_no') + '\\.gif$');
	var img_re = oldstatus == 'none' ? '_no.gif' : '_yes.gif'
	if(img) {
		img.src = img.src.replace(img_regexp, img_re);
	}
}

function imgzoom(o) {
	if(event.ctrlKey) {
		var zoom = parseInt(o.style.zoom, 10) || 100;
		zoom -= event.wheelDelta / 12;
		if(zoom > 0) {
			o.style.zoom = zoom + '%';
		}
		return false;
	} else {
		return true;
	}
}

function insertSmiley(smilieid) {
	var oEditor = FCKeditorAPI.GetInstance('message') ;
	var src = $('smilie_' + smilieid).src;
	if ( oEditor.EditMode == FCK_EDITMODE_WYSIWYG ){
		//oEditor.InsertHtml( '<img src="' + src + '" border="0" alt="" />&nbsp;' ) ;
		oEditor.InsertHtml( '[em' + smilieid + ']' ) ;
	}
	else
		alert( 'You must be on WYSIWYG mode!' ) ;

}

function expandoptions(id)
{
	var a = document.getElementById(id);
	if(a.style.display=='')
	{
		a.style.display='none';
	}
	else
	{
		a.style.display='';
	}
}

function announcement() {
	$('announcement').innerHTML = '<marquee style="filter:progid:DXImageTransform.Microsoft.Alpha(startX=0, startY=0, finishX=10, finishY=100,style=1,opacity=0,finishOpacity=100); margin: 0px" direction="left" scrollamount="2" scrolldelay="1" onMouseOver="this.stop();" onMouseOut="this.start();">' +
	$('announcement').innerHTML + '</marquee>';
	$('announcement').style.display = 'block';
}

function $(id) {
	return document.getElementById(id);
}

// ==page's== //
function showPages(name) { //初始化属性
 this.name = name;      //对象名称
 this.page = 1;         //当前页数
 this.pageCount = 1;    //总页数
 this.dispCount = 1;    //所有贴数
 this.argName = 'page'; //参数名
 this.showTimes = 1;    //打印次数
}

showPages.prototype.getPage = function(){ //丛url获得当前页数,如果变量重复只获取最后一个
 var args = location.search;
 var reg = new RegExp('[\?&]?' + this.argName + '=([^&]*)[&$]?', 'gi');
 var chk = args.match(reg);
 this.page = RegExp.$1;
}

showPages.prototype.checkPages = function(){ //进行当前页数和总页数的验证
 if (isNaN(parseInt(this.page))) this.page = 1;
 if (isNaN(parseInt(this.pageCount))) this.pageCount = 1;
 if (this.page < 1) this.page = 1;
 if (this.pageCount < 1) this.pageCount = 1;
 this.page = parseInt(this.page);
 this.pageCount = parseInt(this.pageCount);
 if (this.page > this.pageCount) this.page = this.pageCount;
}

showPages.prototype.createHtml = function(mode){ //生成html代码

	var ispages = parseInt(this.page);
     var strHtml = '', prevPage = ispages - 1, nextPage = ispages + 1;
	 strHtml += '<div class="p_bar"><a class="p_total">' + this.dispCount + '</a> <a class="p_pages">' + ispages + '/' + this.pageCount + '</a>';
     if (prevPage < 1) {
       strHtml += '<a class="p_num">&#171;</a>';
       strHtml += '<a class="p_num">&#139;</a>';
     } else {
       strHtml += '<a href="javascript:' + this.name + '.toPage(1);" class="p_num">&#171;</a>';
       strHtml += '<a href="javascript:' + this.name + '.toPage(' + prevPage + ');" class="p_num">&#139;</a>';
     }
     if (ispages % 10 ==0) {
       var startPage = ispages - 9;
     } else {
       var startPage = ispages - ispages % 10 + 1;
     }
     if (startPage > 10) strHtml += '<a href="javascript:' + this.name + '.toPage(' + (startPage - 1) + ');" class="p_num">...</a>';
     for (var i = startPage; i < startPage + 10; i++) {
       if (i > this.pageCount) break;
       if (i == ispages) {
         strHtml += '<a class="p_curpage" title="Page ' + i + '" >' + i + '</a>';
       } else {
         strHtml += '<span title="Page ' + i + '"><a href="javascript:' + this.name + '.toPage(' + i + ');" class="p_num">' + i + '</a></span>';
       }
     }
     if (this.pageCount >= startPage + 10) strHtml += '<span title="Next 10 Pages"><a href="javascript:' + this.name + '.toPage(' + (startPage + 10) + ');" class="p_num">...</a></span>';
     if (nextPage > this.pageCount) {
       strHtml += '<a class="p_num" title="Next Page">&#155;</a>';
       strHtml += '<a class="p_num" title="Last Page">&#187;</a>';
     } else {
       strHtml += '<span  title="Next Page"><a class="p_num" href="javascript:' + this.name + '.toPage(' + nextPage + ');">&#155;</a></span>';
       strHtml += '<span title="Last Page"><a class="p_num" href="javascript:' + this.name + '.toPage(' + this.pageCount + ');">&#187;</a></span>';
     }
	 strHtml += '</div>';
 return strHtml;
}

showPages.prototype.createUrl = function (page) { //生成页面跳转url
 if (isNaN(parseInt(page))) page = 1;
 if (page < 1) page = 1;
 if (page > this.pageCount) page = this.pageCount;
 var url = location.protocol + '//' + location.host + location.pathname;
 var args = location.search;
 var reg = new RegExp('([\?&]?)' + this.argName + '=[^&]*[&$]?', 'gi');
 args = args.replace(reg,'$1');
 if (args == '' || args == null) {
   args += '?' + this.argName + '=' + page;
 } else if (args.substr(args.length - 1,1) == '?' || args.substr(args.length - 1,1) == '&') {
     args += this.argName + '=' + page;
 } else {
     args += '&' + this.argName + '=' + page;
 }
 return url + args;
}

showPages.prototype.toPage = function(page){ //页面跳转
 var turnTo = 1;
 if (typeof(page) == 'object') {
   turnTo = page.options[page.selectedIndex].value;
 } else {
   turnTo = page;
 }
 self.location.href = this.createUrl(turnTo);
}

showPages.prototype.printHtml = function(mode){ //显示html代码
 this.getPage();
 this.checkPages();
 this.showTimes += 1;
 document.write('<div id="pages_' + this.name + '_' + this.showTimes + '" class="a4"></div>');
 document.getElementById('pages_' + this.name + '_' + this.showTimes).innerHTML = this.createHtml(mode);
}

showPages.prototype.formatInputPage = function(e){ //限定输入页数格式
 var ie = navigator.appName=="Microsoft Internet Explorer"?true:false;
 if(!ie) var key = e.which;
 else var key = event.keyCode;
 if (key == 8 || key == 46 || (key >= 48 && key <= 57)) return true;
 return false;
}

function attachimg(obj, action, text) {
	if(action == 'load') {
		if(obj.width > screen.width * 0.7) {
			obj.resized = true;
			obj.width = screen.width * 0.7;
			obj.alt = text;
		}
		obj.onload = null;
	} else if(action == 'mouseover') {
		if(obj.resized) {
			obj.style.cursor = 'hand';
		}
	} else if(action == 'click') {
		if(!obj.resized) {
			return false;
		} else {
			window.open(text);
		}
	}
}

function seccheck(theform, seccodecheck , previewpost) {
	if(!previewpost && seccodecheck ) {
		var url = 'ajax.asp?checksubmit=yes&action=';
		if(seccodecheck) {
			var x = new Ajax('XML', '');
			x.get(url + 'checkseccode&seccodeverify=' + $('code').value, function(s) {
				if(s != 'succeed') {
					alert(s);
					$('code').focus();
				}  else {
					postsubmit(theform);
				}
			});
		} 
	} else {
		postsubmit(theform, previewpost);
	}
}

function postsubmit(theform, previewpost) {
	if(!previewpost) {
		theform.topicsubmit.disabled = true;
		theform.submit();
	}
}


function get_Code(){
	if(document.getElementById("imgid"))
		document.getElementById("imgid").innerHTML = '<img src="inc/code.asp?t='+Math.random()+'" alt="点击刷新验证码" style="cursor:pointer;border:0;vertical-align:middle;" onclick="this.src=\'inc/code.asp?t=\'+Math.random()" />'
}

function loadsmile(inum,page){
	var objname='smilieslist';
	var x = new Ajax('HTML', objname);
    x.get('ajax.asp?checksubmit=yes&action=smilies&inum='+inum+'&page='+ page, function(s){
		var obj = $(objname);
		obj.style.display = '';
		obj.innerHTML = s;
	});
}


function loadtopicstop(uname,tid){
	var objname='loadtopicstoplist';
	var x = new Ajax('HTML', objname);
    x.get('ajax.asp?checksubmit=yes&action=loadtopics&uname='+uname+'&tid='+tid, function(s){
		var obj = $(objname);
		obj.style.display = '';
		obj.innerHTML = s;
	});
}

function Ajax(recvType, statusId) {
	var aj = new Object();
	aj.statusId = statusId ? document.getElementById(statusId) : null;
	aj.targetUrl = '';
	aj.sendString = '';
	aj.recvType = recvType ? recvType : 'HTML';
	aj.resultHandle = null;
	aj.createXMLHttpRequest = function() {
		var request = false;
		if(window.XMLHttpRequest) {
			request = new XMLHttpRequest();
			if(request.overrideMimeType) {
				request.overrideMimeType('text/xml');
			}
		} else if(window.ActiveXObject) {
			var versions = ['Microsoft.XMLHTTP', 'MSXML.XMLHTTP', 'Microsoft.XMLHTTP', 'Msxml2.XMLHTTP.7.0', 'Msxml2.XMLHTTP.6.0', 'Msxml2.XMLHTTP.5.0', 'Msxml2.XMLHTTP.4.0', 'MSXML2.XMLHTTP.3.0', 'MSXML2.XMLHTTP'];
			for(var i=0; i<versions.length; i++) {
				try {
					request = new ActiveXObject(versions[i]);
					if(request) {
						return request;
					}
				} catch(e) {
					alert(e.message);
				}
			}
		}
		return request;
	}
	aj.XMLHttpRequest = aj.createXMLHttpRequest();
	aj.processHandle = function() {
		if(aj.XMLHttpRequest.readyState == 1 && aj.statusId) {
			aj.statusId.innerHTML = '正在建立连接';
		} else if(aj.XMLHttpRequest.readyState == 2 && aj.statusId) {
			aj.statusId.innerHTML = '正在接受数据';
		} else if(aj.XMLHttpRequest.readyState == 3 && aj.statusId) {
			aj.statusId.innerHTML = '正在接受数据';
		} else if(aj.XMLHttpRequest.readyState == 4) {
			if(aj.XMLHttpRequest.status == 200) {
				if(aj.recvType == 'HTML') {
					aj.resultHandle(aj.XMLHttpRequest.responseText);
				} else if(aj.recvType == 'XML') {
					aj.resultHandle(aj.XMLHttpRequest.responseXML.lastChild.firstChild.nodeValue);
				}
			} else {
				if(aj.statusId) {
					aj.statusId.innerHTML = "请求方未相应";
				}
			}
		}
	}
	aj.get = function(targetUrl, resultHandle) {
		aj.targetUrl = targetUrl;
		aj.XMLHttpRequest.onreadystatechange = aj.processHandle;
		aj.resultHandle = resultHandle;
		if(window.XMLHttpRequest) {
			aj.XMLHttpRequest.open('GET', aj.targetUrl);
			aj.XMLHttpRequest.send(null);
		} else {
		        aj.XMLHttpRequest.open("GET", targetUrl, true);
		        aj.XMLHttpRequest.send();
		}
	}

	aj.post = function(targetUrl, sendString, resultHandle) {
		aj.targetUrl = targetUrl;
		aj.sendString = sendString;
		aj.XMLHttpRequest.onreadystatechange = aj.processHandle;
		aj.resultHandle = resultHandle;
		aj.XMLHttpRequest.open('POST', targetUrl);
		aj.XMLHttpRequest.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
		aj.XMLHttpRequest.send(aj.sendString);
	}
	return aj;
}