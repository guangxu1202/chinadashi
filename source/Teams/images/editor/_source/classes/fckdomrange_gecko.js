/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 */

FCKDomRange.prototype.MoveToSelection = function()
{
	this.Release( true ) ;

	var oSel = this.Window.getSelection() ;

	if ( oSel.rangeCount == 1 )
	{
		this._Range = FCKW3CRange.CreateFromRange( this.Window.document, oSel.getRangeAt(0) ) ;
		this._UpdateElementInfo() ;
	}
}

FCKDomRange.prototype.Select = function()
{
	var oRange = this._Range ;
	if ( oRange )
	{
		var oDocRange = this.Window.document.createRange() ;
		oDocRange.setStart( oRange.startContainer, oRange.startOffset ) ;

		try
		{
			oDocRange.setEnd( oRange.endContainer, oRange.endOffset ) ;
		}
		catch ( e )
		{
			// There is a bug in Firefox implementation (it would be too easy
			// otherwhise). The new start can't be after the end (W3C says it can).
			// So, let's create a new range and collapse it to the desired point.
			if ( e.toString().Contains( 'NS_ERROR_ILLEGAL_VALUE' ) )
			{
				oRange.collapse( true ) ;
				oDocRange.setEnd( oRange.endContainer, oRange.endOffset ) ;
			}
			else
				throw( e ) ;
		}

		var oSel = this.Window.getSelection() ;
		oSel.removeAllRanges() ;

		// We must add a clone otherwise Firefox will have rendering issues.
		oSel.addRange( oDocRange ) ;
	}
}
