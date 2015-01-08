/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 * Plugin to insert "alipays" in the editor.
 */

// Register the related command.
FCKCommands.RegisterCommand( 'alipay', new FCKDialogCommand( 'alipay', FCKLang.alipayDlgTitle, FCKPlugins.Items['alipay'].Path + 'fck_alipay.html', 340, 550 ) ) ;

// Create the "alipay" toolbar button.
var oalipayItem = new FCKToolbarButton( 'alipay', FCKLang.alipayBtn ) ;
oalipayItem.IconPath = FCKPlugins.Items['alipay'].Path + 'alipay.gif' ;
FCKToolbarItems.RegisterItem( 'alipay', oalipayItem ) ;



var FCKalipays = new Object() ;
FCKalipays.Add = function( name ){
	var oSpan = FCK.CreateElement( 'SPAN' ) ;
	this.SetupSpan( oSpan, name ) ;
}

FCKalipays.SetupSpan = function( span, name ){
	span.innerHTML = name  ;
	span.style.backgroundColor = '#ffff00' ;
	span.style.color = '#000000' ;
	if ( FCKBrowserInfo.IsGecko )
		span.style.cursor = 'default' ;
	span._fckalipay = name ;
	span.contentEditable = false ;
	// To avoid it to be resized.
	span.onresizestart = function()
	{
		FCK.EditorWindow.event.returnValue = false ;
		return false ;
	}
}
