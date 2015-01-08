/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 */

var FCKDocumentFragment = function( parentDocument, baseDocFrag )
{
	this.RootNode = baseDocFrag || parentDocument.createDocumentFragment() ;
}

FCKDocumentFragment.prototype =
{

	// Append the contents of this Document Fragment to another element.
	AppendTo : function( targetNode )
	{
		targetNode.appendChild( this.RootNode ) ;
	},

	InsertAfterNode : function( existingNode )
	{
		FCKDomTools.InsertAfterNode( existingNode, this.RootNode ) ;
	}
}