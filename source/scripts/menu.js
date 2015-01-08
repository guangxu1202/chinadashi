// JavaScript Document
		var preloaded = [];

		for (var i = 1; i <= 7; i++) {
			preloaded[i] = [loadImage(i + "_0.gif"), loadImage(i + "_1.gif")];
		}

		function init() {

			if (mtDropDown.isSupported()) {
				mtDropDown.initialize();
			}
		}

		function loadImage(sFilename) {
			var img = new Image();
			img.src ="" + sFilename;
			return img;
		}

		function swapImage(imgName, sFilename) {
			document.images[imgName].src = sFilename;
		}


if (mtDropDown.isSupported()) {

// x轴：-10；y轴：-6 
		var ms = new mtDropDownSet(mtDropDown.direction.down, 0, -12  , mtDropDown.reference.bottomLeft);

		var menu1 = ms.addMenu(document.getElementById("menu1"));
		menu1.addItem("公司概况", "about.aspx");
		menu1.addItem("管理团队", "Management.aspx");
		menu1.addItem("组织架构", "structure.aspx");
		menu1.addItem("企业荣誉", "honor.aspx");
		menu1.addItem("发展历程", "history.aspx");
 
		var menu2 = ms.addMenu(document.getElementById("menu2"));
		menu2.addItem("公司新闻", "company_news.aspx");
		menu2.addItem("行业信息", "company_news.aspx?Fid=49");
		menu2.addItem("项目动态", "project_news.aspx");
		menu2.addItem("媒体聚焦", "company_news.aspx?Fid=50");
		menu2.addItem("视频新闻", "bulletin.aspx");
		menu2.addItem("中航地产报", "baokanpage.aspx");

		var menu3 = ms.addMenu(document.getElementById("menu3"));
		menu3.addItem("地产项目", "selling.aspx");
		menu3.addItem("物业业务", "property.aspx");
		menu3.addItem("酒店业务", "hotel.aspx");
		//menu3.addItem("其他业务", "other.aspx");

		var menu4 = ms.addMenu(document.getElementById("menu4"));
		menu4.addItem("文化体系", "Cultural.aspx");
		menu4.addItem("企业活动", "culture_active.aspx");
		menu4.addItem("企业公益", "enterprises.aspx");
		menu4.addItem("大事记", "events.aspx");

		var menu5 = ms.addMenu(document.getElementById("menu5"));
		menu5.addItem("中航会", "club.aspx");
		menu5.addItem("客户反馈", "feedback.aspx");
		menu5.addItem("链接你我", "contact.aspx");
		menu5.addItem("投诉与建议", "/complaints");
		//menu5.addItem("法律法规", "regulations.aspx");
		//menu5.addItem("投资者信箱", "exchange.aspx");
		//menu5.addItem("互动平台", "Interac.aspx");
		//menu5.addItem("业绩推介", "promote.aspx");
        
		var menu6 = ms.addMenu(document.getElementById("menu6"));
		menu6.addItem("人才理念", "hr_idea.aspx");
		menu6.addItem("社会招聘", "social.aspx");
		menu6.addItem("校园招聘", "campus.aspx");
		
		var menu7 = ms.addMenu(document.getElementById("menu7"));
		menu7.addItem("品牌管理", "brand_k.aspx");
		menu7.addItem("品牌活动", "brand_ac.aspx");
		menu7.addItem("形象画册", "MagzineView.html");
		menu7.addItem("媒体报道", "brand_media.aspx");
		menu7.addItem("新闻联络", "news_contact.aspx");
		

		mtDropDown.renderAll();
	}



mtDropDown.spacerGif = "images/menu_right.jpg"; 
mtDropDown.dingbatOn = ""; 
mtDropDown.dingbatOff = ""; 
mtDropDown.dingbatSize = 14; 
mtDropDown.menuPadding = 1; 
mtDropDown.itemPadding = 4; 
mtDropDown.shadowSize = 9; 
mtDropDown.shadowOffset = 3; 
mtDropDown.shadowColor = ""; 
mtDropDown.shadowPng = "images/menu_left.jpg"; 
mtDropDown.backgroundColor = ""; 
mtDropDown.backgroundPng = ""; 
mtDropDown.hideDelay = 200; 
mtDropDown.slideTime = 300; 

mtDropDown.reference = {topLeft:1,topRight:2,bottomLeft:3,bottomRight:4};
mtDropDown.direction = {down:1,right:2};
mtDropDown.registry = [];
mtDropDown._maxZ = 100;

mtDropDown.isSupported = function() {
	if (typeof mtDropDown.isSupported.r == "boolean") 
		return mtDropDown.isSupported.r;

	var ua = navigator.userAgent.toLowerCase();
	var an = navigator.appName;
	var r = false;

	if (ua.indexOf("gecko") > -1) r = true; 
	else if (an == "Microsoft Internet Explorer") {
		if (document.getElementById) r = true; 
	}

	mtDropDown.isSupported.r = r;
	return r;
}

mtDropDown.initialize = function() {
	for (var i = 0, menu = null; menu = this.registry[i]; i++) {
		menu.initialize();
	}
}

mtDropDown.renderAll = function() {
	var aMenuHtml = [];
	for (var i = 0, menu = null; menu = this.registry[i]; i++) {
		aMenuHtml[i] = menu.toString();
	}

	document.write(aMenuHtml.join(""));
}

/////////////////////////////// class mtDropDown BEGINS /////////////

function mtDropDown(oActuator, iDirection, iLeft, iTop, iReferencePoint, parentMenuSet) {

	this.addItem = addItem;
	this.addMenu = addMenu;
	this.toString = toString;
	this.initialize = initialize;
	this.isOpen = false;
	this.show = show;
	this.hide = hide;
	this.items = [];

	this.onactivate = new Function(); 
	this.ondeactivate = new Function(); 
	this.onmouseover = new Function(); 
	this.onqueue = new Function(); 

	this.index = mtDropDown.registry.length;
	mtDropDown.registry[this.index] = this;
	var id = "mtDropDown" + this.index;
	var contentHeight = null;
	var contentWidth = null;
	var childMenuSet = null;
	var animating = false;
	var childMenus = [];
	var slideAccel = -1;
	var elmCache = null;
	var ready = false;
	var _this = this;
	var a = null;
	var pos = iDirection == mtDropDown.direction.down ? "top" : "left";
	var dim = null;

	function addItem(sText, sUrl) {
		var item = new mtDropDownItem(sText, sUrl, this);
		item._index = this.items.length;
		this.items[item._index] = item;
	}

	function addMenu(oMenuItem) {
		if (!oMenuItem.parentMenu == this) throw new Error("Cannot add a menu here");
		if (childMenuSet == null) childMenuSet = new mtDropDownSet(mtDropDown.direction.right, -5, 2, mtDropDown.reference.topRight);
		var m = childMenuSet.addMenu(oMenuItem);
		childMenus[oMenuItem._index] = m;
		m.onmouseover = child_mouseover;
		m.ondeactivate = child_deactivate;
		m.onqueue = child_queue;
		return m;
	}

	function initialize() {
		initCache();
		initEvents();
		initSize();
		ready = true;
	}

	function show() {
		if (ready) {
			_this.isOpen = true;
			animating = true;
			setContainerPos();
			elmCache["clip"].style.visibility = "visible";
			elmCache["clip"].style.zIndex = mtDropDown._maxZ++;

			slideStart();
			_this.onactivate();
		}
	}

	function hide() {
		if (ready) {
			_this.isOpen = false;
			animating = true;
			for (var i = 0, item = null; item = elmCache.item[i]; i++) 
			dehighlight(item);
			if (childMenuSet) childMenuSet.hide();
			slideStart();
			_this.ondeactivate();
		}
	}

	function setContainerPos() {
		var sub = oActuator.constructor == mtDropDownItem; 
		var act = sub ? oActuator.parentMenu.elmCache["item"][oActuator._index] : oActuator; 
		var el = act;
		var x = 0;
		var y = 0;
		var minX = 0;
		var maxX = (window.innerWidth ? window.innerWidth : document.body.clientWidth) - parseInt(elmCache["clip"].style.width);
		var minY = 0;
		var maxY = (window.innerHeight ? window.innerHeight : document.body.clientHeight) - parseInt(elmCache["clip"].style.height);

		while (sub ? el.parentNode.className.indexOf("mtDropdownMenu") == -1 : el.offsetParent) {
			x += el.offsetLeft;
			y += el.offsetTop;
			if (el.scrollLeft) x -= el.scrollLeft;
			if (el.scrollTop) y -= el.scrollTop;
			el = el.offsetParent;
		}

		if (oActuator.constructor == mtDropDownItem) {
			x += parseInt(el.parentNode.style.left);
			y += parseInt(el.parentNode.style.top);
		}

		switch (iReferencePoint) {
			case mtDropDown.reference.topLeft:
			break;

			case mtDropDown.reference.topRight:
			x += act.offsetWidth;
			break;

			case mtDropDown.reference.bottomLeft:
			y += act.offsetHeight;
			break;

			case mtDropDown.reference.bottomRight:
			x += act.offsetWidth;
			y += act.offsetHeight;
			break;
		}

		x += iLeft;
		y += iTop;
		x = Math.max(Math.min(x, maxX), minX);
		y = Math.max(Math.min(y, maxY), minY);
		elmCache["clip"].style.left = x + "px";
		elmCache["clip"].style.top = y + "px";
	}

	function slideStart() {
		var x0 = parseInt(elmCache["content"].style[pos]);
		var x1 = _this.isOpen ? 0 : -dim;
		if (a != null) a.stop();
		a = new Accelimation(x0, x1, mtDropDown.slideTime, slideAccel);
		a.onframe = slideFrame;
		a.onend = slideEnd;
		a.start();
	}

	function slideFrame(x) {
		elmCache["content"].style[pos] = x + "px";
	}

	function slideEnd() {
		if (!_this.isOpen) elmCache["clip"].style.visibility = "hidden";
			animating = false;
	}

	function initSize() {

		var ow = elmCache["items"].offsetWidth;
		var oh = elmCache["items"].offsetHeight;
		var ua = navigator.userAgent.toLowerCase();

		elmCache["clip"].style.width = ow + mtDropDown.shadowSize + 2 + "px";
		elmCache["clip"].style.height = oh + mtDropDown.shadowSize + 2 + "px";

		elmCache["content"].style.width = ow + mtDropDown.shadowSize + "px";
		elmCache["content"].style.height = oh + mtDropDown.shadowSize + "px";
		contentHeight = oh + mtDropDown.shadowSize;
		contentWidth = ow + mtDropDown.shadowSize;
		dim = iDirection == mtDropDown.direction.down ? contentHeight : contentWidth;

		elmCache["content"].style[pos] = -dim - mtDropDown.shadowSize + "px";
		elmCache["clip"].style.visibility = "hidden";

		if (ua.indexOf("mac") == -1 || ua.indexOf("gecko") > -1) {

			elmCache["background"].style.width = ow + "px";
			elmCache["background"].style.height = oh + "px";
			elmCache["background"].style.backgroundColor = mtDropDown.backgroundColor;

			elmCache["shadowRight"].style.left = ow + "px";
			elmCache["shadowRight"].style.height = oh - (mtDropDown.shadowOffset - mtDropDown.shadowSize) + "px";
			elmCache["shadowRight"].style.backgroundColor = mtDropDown.shadowColor;



			//elmCache["shadowLeft"].style.left = ow + "px";
			elmCache["shadowLeft"].style.height = oh - (mtDropDown.shadowOffset - mtDropDown.shadowSize) + "px";
			elmCache["shadowLeft"].style.backgroundColor = mtDropDown.shadowColor;
						elmCache["shadowLeft"].firstChild.src = mtDropDown.shadowPng;
			//elmCache["shadowLeft"].style.left = ow + "px";
			elmCache["shadowLeft"].firstChild.width = mtDropDown.shadowSize;
			//elmCache["shadowLeft"].firstChild.height = oh - (mtDropDown.shadowOffset - mtDropDown.shadowSize);
		} else {
			elmCache["background"].firstChild.src = mtDropDown.backgroundPng;
			elmCache["background"].firstChild.width = ow;
			elmCache["background"].firstChild.height = oh;

			elmCache["shadowRight"].firstChild.src = mtDropDown.shadowPng;
			elmCache["shadowRight"].style.left = ow + "px";
			elmCache["shadowRight"].firstChild.width = mtDropDown.shadowSize;
			elmCache["shadowRight"].firstChild.height = oh - (mtDropDown.shadowOffset - mtDropDown.shadowSize);

			//elmCache["shadowLeft"].firstChild.src = mtDropDown.shadowPng;
			elmCache["shadowLeft"].style.left = ow + "px";
			elmCache["shadowLeft"].firstChild.width = mtDropDown.shadowSize;
			//elmCache["shadowLeft"].firstChild.height = oh - (mtDropDown.shadowOffset - mtDropDown.shadowSize);
		}
	}

	function initCache() {
		var menu = document.getElementById(id);
		var all = menu.all ? menu.all : menu.getElementsByTagName("*"); 
		elmCache = {};
		elmCache["clip"] = menu;
		elmCache["item"] = [];
		for (var i = 0, elm = null; elm = all[i]; i++) {
			switch (elm.className) {
				case "items":
				case "content":
				case "background":
				case "shadowRight":
				case "shadowLeft":
				elmCache[elm.className] = elm;
				break;
				case "item":
				elm._index = elmCache["item"].length;
				elmCache["item"][elm._index] = elm;
				break;
			}
		}

		_this.elmCache = elmCache;
	}

	function initEvents() {

		for (var i = 0, item = null; item = elmCache.item[i]; i++) {
			item.onmouseover = item_mouseover;
			item.onmouseout = item_mouseout;
			item.onclick = item_click;
		}

		if (typeof oActuator.tagName != "undefined") {
			oActuator.onmouseover = actuator_mouseover;
			oActuator.onmouseout = actuator_mouseout;
		}

		elmCache["content"].onmouseover = content_mouseover;
		elmCache["content"].onmouseout = content_mouseout;
	}

	function highlight(oRow) {
		oRow.className = "item hover";
		if (childMenus[oRow._index]) 
			oRow.lastChild.firstChild.src = mtDropDown.dingbatOn;
		}

	function dehighlight(oRow) {
		oRow.className = "item";
		if (childMenus[oRow._index]) 
			oRow.lastChild.firstChild.src = mtDropDown.dingbatOff;
	}

	function item_mouseover() {
		if (!animating) {
			highlight(this);

			if (childMenus[this._index]) 
				childMenuSet.showMenu(childMenus[this._index]);
			else if (childMenuSet) childMenuSet.hide();
		}
	}

	function item_mouseout() {
		if (!animating) {
			if (childMenus[this._index])
				childMenuSet.hideMenu(childMenus[this._index]);
			else 
				dehighlight(this);
		}
	}

	function item_click() {
		if (!animating) {
			//if (_this.items[this._index].url) 
	//			location.href = _this.items[this._index].url;
		if(_this.items[this._index].url!="/complaints"){
			location.href=_this.items[this._index].url;
		}else{
			window.open(_this.items[this._index].url,"投诉与建议");
		}
		}
	}

	function actuator_mouseover() {
		parentMenuSet.showMenu(_this);
	}

	function actuator_mouseout() {
		parentMenuSet.hideMenu(_this);
	}

	function content_mouseover() {
		if (!animating) {
			parentMenuSet.showMenu(_this);
			_this.onmouseover();
		}
	}

	function content_mouseout() {
		if (!animating) {
			parentMenuSet.hideMenu(_this);
		}
	}

	function child_mouseover() {
		if (!animating) {
			parentMenuSet.showMenu(_this);
		}
	}

	function child_deactivate() {
		for (var i = 0; i < childMenus.length; i++) {
			if (childMenus[i] == this) {
				dehighlight(elmCache["item"][i]);
				break;
			}
		}
	}

	function child_queue() {
		parentMenuSet.hideMenu(_this);
	}

	function toString() {
		var aHtml = [];
		var sClassName = "mtDropdownMenu" + (oActuator.constructor != mtDropDownItem ? " top" : "");
		for (var i = 0, item = null; item = this.items[i]; i++) {
			aHtml[i] = item.toString(childMenus[i]);
		}

		return '<div id="' + id + '" class="' + sClassName + '">' + 
		'<div class="content"><table class="items" cellpadding="0" cellspacing="1" border="0">' + 
		aHtml.join('') + 
		'</table>' + 
		'<div class="shadowLeft"><img src="' + mtDropDown.spacerGif + '" width="9" height="23"></div>' + 
		'<div class="shadowRight"><img src="' + mtDropDown.spacerGif + '" width="9" height="23"></div>' + 
		'<div class="background"><img src="' + mtDropDown.spacerGif + '" width="1" height="1"></div>' + 
		'</div></div>';
	}
}

/////////////////////////////// class mtDropDown ENDS /////////////

mtDropDownSet.registry = [];
function mtDropDownSet(iDirection, iLeft, iTop, iReferencePoint) {

	this.addMenu = addMenu;
	this.showMenu = showMenu;
	this.hideMenu = hideMenu;
	this.hide = hide;

	var menus = [];
	var _this = this;
	var current = null;
	this.index = mtDropDownSet.registry.length;
	mtDropDownSet.registry[this.index] = this;

	function addMenu(oActuator) {
		var m = new mtDropDown(oActuator, iDirection, iLeft, iTop, iReferencePoint, this);
		menus[menus.length] = m;
		return m;
	}

	function showMenu(oMenu) {
		if (oMenu != current) {
			if (current != null) hide(current); 
				current = oMenu;
			oMenu.show();
		} else {
			cancelHide(oMenu);
		}
	}

	function hideMenu(oMenu) {
		if (current == oMenu && oMenu.isOpen) {
			if (!oMenu.hideTimer) scheduleHide(oMenu);
		}
	}

	function scheduleHide(oMenu) {
		oMenu.onqueue();
		oMenu.hideTimer = window.setTimeout("mtDropDownSet.registry[" + _this.index + "].hide(mtDropDown.registry[" + oMenu.index + "])", mtDropDown.hideDelay);
	}

	function cancelHide(oMenu) {
		if (oMenu.hideTimer) {
			window.clearTimeout(oMenu.hideTimer);
			oMenu.hideTimer = null;
		}
	}

	function hide(oMenu) { 
		if (!oMenu && current) oMenu = current;
		if (oMenu && current == oMenu && oMenu.isOpen) {

			cancelHide(oMenu);
			current = null;
			oMenu.hideTimer = null;
			oMenu.hide();
		}
	}
}

function mtDropDownItem(sText, sUrl, oParent) {
	this.toString = toString;
	this.text = sText;
	this.url = sUrl;
	this.parentMenu = oParent;

	function toString(bDingbat) {
		var sDingbat = bDingbat ? mtDropDown.dingbatOff : mtDropDown.spacerGif;
		var iEdgePadding = mtDropDown.itemPadding + mtDropDown.menuPadding;
		var sPaddingLeft = "padding:" + mtDropDown.itemPadding + "px; padding-left:" + iEdgePadding + "px;"
		var sPaddingRight = "padding:" + mtDropDown.itemPadding + "px; padding-right:" + iEdgePadding + "px;"

///////////////////////////////////////////////////////////////// 横条
		return '<td class="item" nowrap>' + 
		sText + '</td>'; 
	}
}

function Accelimation(from, to, time, zip) {
	if (typeof zip == "undefined") zip = 0;
	if (typeof unit == "undefined") unit = "px";
	this.x0 = from;
	this.x1 = to;
	this.dt = time;
	this.zip = -zip;
	this.unit = unit;
	this.timer = null;
	this.onend = new Function();
	this.onframe = new Function();
}

Accelimation.prototype.start = function() {
	this.t0 = new Date().getTime();
	this.t1 = this.t0 + this.dt;
	var dx = this.x1 - this.x0;
	this.c1 = this.x0 + ((1 + this.zip) * dx / 3);
	this.c2 = this.x0 + ((2 + this.zip) * dx / 3);
	Accelimation._add(this);
}

Accelimation.prototype.stop = function() {
	Accelimation._remove(this);
}

Accelimation.prototype._paint = function(time) {
	if (time < this.t1) {
		var elapsed = time - this.t0;
		this.onframe(Accelimation._getBezier(elapsed/this.dt,this.x0,this.x1,this.c1,this.c2));
	}
	else this._end();
}

Accelimation.prototype._end = function() {
	Accelimation._remove(this);
	this.onframe(this.x1);
	this.onend();
}

Accelimation._add = function(o) {
	var index = this.instances.length;
	this.instances[index] = o;

	if (this.instances.length == 1) {
		this.timerID = window.setInterval("Accelimation._paintAll()", this.targetRes);
	}
}

Accelimation._remove = function(o) {
	for (var i = 0; i < this.instances.length; i++) {
		if (o == this.instances[i]) {
			this.instances = this.instances.slice(0,i).concat( this.instances.slice(i+1) );
			break;
		}
	}

	if (this.instances.length == 0) {
		window.clearInterval(this.timerID);
		this.timerID = null;
	}
}

Accelimation._paintAll = function() {
		var now = new Date().getTime();
		for (var i = 0; i < this.instances.length; i++) {
			this.instances[i]._paint(now);
	}
}

Accelimation._B1 = function(t) { return t*t*t }
Accelimation._B2 = function(t) { return 3*t*t*(1-t) }
Accelimation._B3 = function(t) { return 3*t*(1-t)*(1-t) }
Accelimation._B4 = function(t) { return (1-t)*(1-t)*(1-t) }

Accelimation._getBezier = function(percent,startPos,endPos,control1,control2) {
	return endPos * this._B1(percent) + control2 * this._B2(percent) + control1 * this._B3(percent) + startPos * this._B4(percent);
}

Accelimation.instances = [];
Accelimation.targetRes = 10;
Accelimation.timerID = null;
