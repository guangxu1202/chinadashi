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
		var ms = new mtDropDownSet(mtDropDown.direction.down, 0, 2  , mtDropDown.reference.bottomLeft);

		var menu1 = ms.addMenu(document.getElementById("menu1"));
		//menu1.addItem("公司概况", "about.aspx");
		//menu1.addItem("管理团队", "Management.aspx");
		//menu1.addItem("组织架构", "structure.aspx");
		//menu1.addItem("企业荣誉", "honor.aspx");
		//menu1.addItem("发展历程", "history.aspx");
 
		var menu2 = ms.addMenu(document.getElementById("menu2"));
		menu2.addItem("公司新闻&nbsp;&nbsp;&nbsp;&nbsp;|", "news.asp");
		menu2.addItem("产业新闻&nbsp;&nbsp;&nbsp;&nbsp;|", "news.asp?act=2");
		menu2.addItem("媒体报道", "news.asp?act=3");

		var menu3 = ms.addMenu(document.getElementById("menu3"));
		menu3.addItem("领导专栏&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp?act=1");
		menu3.addItem("大实简介&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp");
		menu3.addItem("大实荣誉&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp?act=2");
		menu3.addItem("大实架构&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp?act=3");
		menu3.addItem("大事记", "about_dsj.asp");

		var menu4 = ms.addMenu(document.getElementById("menu4"));
		menu4.addItem("大实理念&nbsp;&nbsp;&nbsp;&nbsp;|", "brand.asp");
		menu4.addItem("大实之道&nbsp;&nbsp;&nbsp;&nbsp;|", "brand.asp?act=2");
		menu4.addItem("大实文化&nbsp;&nbsp;&nbsp;&nbsp;|", "brand.asp?act=3");
		menu4.addItem("产品品牌", "brand_list.asp");

		var menu5 = ms.addMenu(document.getElementById("menu5"));
		menu5.addItem("房地产开发&nbsp;&nbsp;&nbsp;&nbsp;|", "estate.asp");
		menu5.addItem("其他产业", "estate.asp?act=1");
		//menu5.addItem("链接你我", "contact.aspx");
		//menu5.addItem("投诉与建议", "/complaints");
		//menu5.addItem("法律法规", "regulations.aspx");
		//menu5.addItem("投资者信箱", "exchange.aspx");
		//menu5.addItem("互动平台", "Interac.aspx");
		//menu5.addItem("业绩推介", "promote.aspx");
        
		
		var menu6 = ms.addMenu(document.getElementById("menu6"));
		menu6.addItem("职位信息&nbsp;&nbsp;&nbsp;&nbsp;|", "join.asp");
		//menu6.addItem("社会招聘", "social.aspx");
		menu6.addItem("校园招聘", "join_s.asp");
		
		var menu7 = ms.addMenu(document.getElementById("menu7"));
		//menu7.addItem("品牌管理", "brand_k.aspx");
		//menu7.addItem("品牌活动", "brand_ac.aspx");
		//menu7.addItem("形象画册", "MagzineView.html");
		//menu7.addItem("媒体报道", "brand_media.aspx");
		//menu7.addItem("新闻联络", "news_contact.aspx");
		

		mtDropDown.renderAll();
	}
