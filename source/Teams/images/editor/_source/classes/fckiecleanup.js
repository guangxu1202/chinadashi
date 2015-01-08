/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 * FCKIECleanup Class: a generic class used as a tool to remove IE leaks.
 */

var	FCKIECleanup = function( attachWindow )
{
	// If the attachWindow already have a cleanup object, jusgt use that one.
	if ( attachWindow._FCKCleanupObj )
		this.Items = attachWindow._FCKCleanupObj.Items ;
	else
	{
		this.Items = new Array() ;

		attachWindow._FCKCleanupObj = this ;
		FCKTools.AddEventListenerEx( attachWindow, 'unload', FCKIECleanup_Cleanup ) ;
//		attachWindow.attachEvent( 'onunload', FCKIECleanup_Cleanup ) ;
	}
}

FCKIECleanup.prototype.AddItem = function( dirtyItem, cleanupFunction )
{
	this.Items.push( [ dirtyItem, cleanupFunction ] ) ;
}

function FCKIECleanup_Cleanup()
{
	if ( !this._FCKCleanupObj )
		return ;

	var aItems = this._FCKCleanupObj.Items ;

	while ( aItems.length > 0 )
	{

		// It is important to remove from the end to the beginning (pop()),
		// because of the order things get created in the editor. In the code,
		// elements in deeper position in the DOM are placed at the end of the
		// cleanup function, so we must cleanup then first, otherwise IE could
		// throw some crazy memory errors (IE bug).
		var oItem = aItems.pop() ;
		if ( oItem )
			oItem[1].call( oItem[0] ) ;
	}

	this._FCKCleanupObj = null ;

	if ( CollectGarbage )
		CollectGarbage() ;
}