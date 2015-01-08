/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 */
// Disable the custom Enter Key Handler. This option will be removed in version 2.5.
FCKConfig.DisableEnterKeyHandler = false ;
FCKConfig.CustomConfigurationsPath = '' ;
FCKConfig.EditorAreaCSS = FCKConfig.BasePath + 'css/fck_editorarea.css' ;
FCKConfig.ToolbarComboPreviewCSS = '' ;
FCKConfig.DocType = '' ;
FCKConfig.BaseHref = '' ;
FCKConfig.FullPage = false ;
FCKConfig.Debug = false ;
FCKConfig.AllowQueryStringDebug = true ;
FCKConfig.SkinPath = FCKConfig.BasePath + 'skins/default/' ;
FCKConfig.PreloadImages = [ FCKConfig.SkinPath + 'images/toolbar.start.gif', FCKConfig.SkinPath + 'images/toolbar.buttonarrow.gif' ] ;
FCKConfig.PluginsPath = FCKConfig.BasePath + 'plugins/' ;
// FCKConfig.Plugins.Add( 'autogrow' ) ;
FCKConfig.AutoGrowMax = 400 ;
// FCKConfig.ProtectedSource.Add( /<%[\s\S]*?%>/g ) ;	// ASP style server side code <%...%>
// FCKConfig.ProtectedSource.Add( /<\?[\s\S]*?\?>/g ) ;	// PHP style server side code
// FCKConfig.ProtectedSource.Add( /(<asp:[^\>]+>[\s|\S]*?<\/asp:[^\>]+>)|(<asp:[^\>]+\/>)/gi ) ;	// ASP.Net style tags <asp:control>
FCKConfig.AutoDetectLanguage	= true ;
FCKConfig.DefaultLanguage		= 'en' ;
FCKConfig.ContentLangDirection	= 'ltr' ;
FCKConfig.ProcessHTMLEntities	= true ;
FCKConfig.IncludeLatinEntities	= true ;
FCKConfig.IncludeGreekEntities	= true ;
FCKConfig.ProcessNumericEntities = false ;
FCKConfig.AdditionalNumericEntities = ''  ;		// Single Quote: "'"
FCKConfig.FillEmptyBlocks	= true ;
FCKConfig.FormatSource		= true ;
FCKConfig.FormatOutput		= true ;
FCKConfig.FormatIndentator	= '    ' ;
FCKConfig.ForceStrongEm = true ;
FCKConfig.GeckoUseSPAN	= false ;
FCKConfig.StartupFocus	= false ;
FCKConfig.ForcePasteAsPlainText	= false ;
FCKConfig.AutoDetectPasteFromWord = true ;	// IE only.
FCKConfig.ForceSimpleAmpersand	= false ;
FCKConfig.TabSpaces		= 0 ;
FCKConfig.ShowBorders	= true ;
FCKConfig.SourcePopup	= false ;
FCKConfig.ToolbarStartExpanded	= true ;
FCKConfig.ToolbarCanCollapse	= true ;
FCKConfig.IgnoreEmptyParagraphValue = true ;
FCKConfig.PreserveSessionOnFileBrowser = false ;
FCKConfig.FloatingPanelsZIndex = 10000 ;
FCKConfig.TemplateReplaceAll = true ;
FCKConfig.TemplateReplaceCheckbox = true ;
FCKConfig.ToolbarLocation = 'In' ;
FCKConfig.ToolbarSets["Default"] = [
	['NewPage','Preview','Templates'],
	['Cut','Copy','Paste','PasteText','PasteWord'],
	['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat','-','Link','Unlink'],
	['SpecialChar','FitWindow'],
	'/',
	['Bold','Italic','Underline','StrikeThrough','-','Subscript','Superscript'],
	['OrderedList','UnorderedList','-','Outdent','Indent'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyFull'],
	['TextColor','BGColor','-','Image','Flash','MadMedia','Table','Rule','alipay'],
	'/',
	['Style','FontFormat','FontName','FontSize']
] ;
FCKConfig.ToolbarSets["Basic"] = [
	['NewPage','FontName','FontSize'],['TextColor','BGColor','-','Bold','Italic','Underline','OrderedList','UnorderedList','-','Link','Unlink','Image','Flash','-']
] ;
FCKConfig.EnterMode = 'p' ;			// p | div | br
FCKConfig.ShiftEnterMode = 'br' ;	// p | div | br
FCKConfig.Keystrokes = [
	[ CTRL + 65 /*A*/, true ],
	[ CTRL + 67 /*C*/, true ],
	[ CTRL + 70 /*F*/, true ],
	[ CTRL + 83 /*S*/, true ],
	[ CTRL + 88 /*X*/, true ],
	[ CTRL + 86 /*V*/, 'Paste' ],
	[ SHIFT + 45 /*INS*/, 'Paste' ],
	[ CTRL + 90 /*Z*/, 'Undo' ],
	[ CTRL + 89 /*Y*/, 'Redo' ],
	[ CTRL + SHIFT + 90 /*Z*/, 'Redo' ],
	[ CTRL + 76 /*L*/, 'Link' ],
	[ CTRL + 66 /*B*/, 'Bold' ],
	[ CTRL + 73 /*I*/, 'Italic' ],
	[ CTRL + 85 /*U*/, 'Underline' ],
	[ CTRL + SHIFT + 83 /*S*/, 'Save' ],
	[ CTRL + ALT + 13 /*ENTER*/, 'FitWindow' ],
	[ CTRL + 9 /*TAB*/, 'Source' ]
] ;
FCKConfig.ContextMenu = ['Generic','Link','Anchor','Image','Flash','BulletedList','NumberedList','Table'] ;
FCKConfig.FontColors = '000000,993300,333300,003300,003366,000080,333399,333333,800000,FF6600,808000,808080,008080,0000FF,666699,808080,FF0000,FF9900,99CC00,339966,33CCCC,3366FF,800080,999999,FF00FF,FFCC00,FFFF00,00FF00,00FFFF,00CCFF,993366,C0C0C0,FF99CC,FFCC99,FFFF99,CCFFCC,CCFFFF,99CCFF,CC99FF,FFFFFF' ;
FCKConfig.FontNames		= '宋体;黑体;隶书;楷体_GB2312;华文行楷;宋体;Arial;Comic Sans MS;Courier New;Tahoma;Times New Roman;Verdana' ;
FCKConfig.FontSizes		= '1/xx-small;2/x-small;3/small;4/medium;5/large;6/x-large;7/xx-large' ;
FCKConfig.FontFormats	= 'p;div;pre;address;h1;h2;h3;h4;h5;h6' ;
FCKConfig.StylesXmlPath		= FCKConfig.EditorPath + 'fckstyles.xml' ;
FCKConfig.TemplatesXmlPath	= FCKConfig.EditorPath + 'fcktemplates.xml' ;
FCKConfig.SpellChecker			= 'ieSpell' ;	// 'ieSpell' | 'SpellerPages'
FCKConfig.MaxUndoLevels = 15 ;
FCKConfig.DisableObjectResizing = false ;
FCKConfig.DisableFFTableHandles = true ;
FCKConfig.LinkDlgHideTarget		= false ;
FCKConfig.LinkDlgHideAdvanced	= false ;
FCKConfig.ImageDlgHideLink		= false ;
FCKConfig.ImageDlgHideAdvanced	= false ;
FCKConfig.FlashDlgHideAdvanced	= false ;
FCKConfig.ProtectedTags = '' ;
// This will be applied to the body element of the editor
FCKConfig.BodyId = '' ;
FCKConfig.BodyClass = '' ;
// The option switches between trying to keep the html structure or do the changes so the content looks like it was in Word
FCKConfig.CleanWordKeepsStructure = false ;
FCKConfig.SmileyPath	= FCKConfig.BasePath + '..\/Emotions\/' ;
FCKConfig.SmileyImages	= ['1.gif','2.gif','3.gif','4.gif','5.gif','6.gif','7.gif','8.gif','9.gif','10.gif','11.gif','12.gif','13.gif','14.gif','15.gif','16.gif','17.gif','18.gif','19.gif','20.gif','21.gif','22.gif','23.gif','24.gif','25.gif','26.gif','27.gif','28.gif','29.gif','30.gif','31.gif','32.gif','33.gif','34.gif','35.gif','36.gif','37.gif','38.gif','39.gif','40.gif','41.gif','42.gif','43.gif'] ; ;
FCKConfig.SmileyColumns = 6 ;
FCKConfig.SmileyWindowWidth		= 320 ;
FCKConfig.SmileyWindowHeight	= 520 ;
//plus return
var sOtherPluginPath = FCKConfig.BasePath.substr(0, FCKConfig.BasePath.length - 7) + 'editor/plugins/' ;
FCKConfig.Plugins.Add( 'alipay', 'zh-cn', sOtherPluginPath ) ;
