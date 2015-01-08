/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 */

var FCKDocumentFragment = function( parentDocument )
{
	this._Document = parentDocument ;
	this.RootNode = parentDocument.createElement( 'div' ) ;
}

// Append the contents of this Document Fragment to another node.
FCKDocumentFragment.prototype =
{

	AppendTo : function( targetNode )
	{
		FCKDomTools.MoveChildren( this.RootNode, targetNode ) ;
	},

	AppendHtml : function( html )
	{
		var eTmpDiv = this._Document.createElement( 'div' ) ;
		eTmpDiv.innerHTML = html ;
		FCKDomTools.MoveChildren( eTmpDiv, this.RootNode ) ;
	},

	InsertAfterNode : function( existingNode )
	{
		var eRoot = this.RootNode ;
		var eLast ;

		while( ( eLast = eRoot.lastChild ) )
			FCKDomTools.InsertAfterNode( existingNode, eRoot.removeChild( eLast ) ) ;
	}
} ;