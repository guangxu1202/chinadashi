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

// x�᣺-10��y�᣺-6 
		var ms = new mtDropDownSet(mtDropDown.direction.down, 0, 2  , mtDropDown.reference.bottomLeft);

		var menu1 = ms.addMenu(document.getElementById("menu1"));
		//menu1.addItem("��˾�ſ�", "about.aspx");
		//menu1.addItem("�����Ŷ�", "Management.aspx");
		//menu1.addItem("��֯�ܹ�", "structure.aspx");
		//menu1.addItem("��ҵ����", "honor.aspx");
		//menu1.addItem("��չ����", "history.aspx");
 
		var menu2 = ms.addMenu(document.getElementById("menu2"));
		menu2.addItem("��˾����&nbsp;&nbsp;&nbsp;&nbsp;|", "news.asp");
		menu2.addItem("��ҵ����&nbsp;&nbsp;&nbsp;&nbsp;|", "news.asp?act=2");
		menu2.addItem("ý�屨��", "news.asp?act=3");

		var menu3 = ms.addMenu(document.getElementById("menu3"));
		menu3.addItem("�쵼ר��&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp?act=1");
		menu3.addItem("��ʵ���&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp");
		menu3.addItem("��ʵ����&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp?act=2");
		menu3.addItem("��ʵ�ܹ�&nbsp;&nbsp;&nbsp;&nbsp;|", "about.asp?act=3");
		menu3.addItem("���¼�", "about_dsj.asp");

		var menu4 = ms.addMenu(document.getElementById("menu4"));
		menu4.addItem("��ʵ����&nbsp;&nbsp;&nbsp;&nbsp;|", "brand.asp");
		menu4.addItem("��ʵ֮��&nbsp;&nbsp;&nbsp;&nbsp;|", "brand.asp?act=2");
		menu4.addItem("��ʵ�Ļ�&nbsp;&nbsp;&nbsp;&nbsp;|", "brand.asp?act=3");
		menu4.addItem("��ƷƷ��", "brand_list.asp");

		var menu5 = ms.addMenu(document.getElementById("menu5"));
		menu5.addItem("���ز�����&nbsp;&nbsp;&nbsp;&nbsp;|", "estate.asp");
		menu5.addItem("������ҵ", "estate.asp?act=1");
		//menu5.addItem("��������", "contact.aspx");
		//menu5.addItem("Ͷ���뽨��", "/complaints");
		//menu5.addItem("���ɷ���", "regulations.aspx");
		//menu5.addItem("Ͷ��������", "exchange.aspx");
		//menu5.addItem("����ƽ̨", "Interac.aspx");
		//menu5.addItem("ҵ���ƽ�", "promote.aspx");
        
		
		var menu6 = ms.addMenu(document.getElementById("menu6"));
		menu6.addItem("ְλ��Ϣ&nbsp;&nbsp;&nbsp;&nbsp;|", "join.asp");
		//menu6.addItem("�����Ƹ", "social.aspx");
		menu6.addItem("У԰��Ƹ", "join_s.asp");
		
		var menu7 = ms.addMenu(document.getElementById("menu7"));
		//menu7.addItem("Ʒ�ƹ���", "brand_k.aspx");
		//menu7.addItem("Ʒ�ƻ", "brand_ac.aspx");
		//menu7.addItem("���󻭲�", "MagzineView.html");
		//menu7.addItem("ý�屨��", "brand_media.aspx");
		//menu7.addItem("��������", "news_contact.aspx");
		

		mtDropDown.renderAll();
	}
