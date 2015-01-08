/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 * FCKEvents Class: used to handle events is a advanced way.
 */

var FCKEvents = function( eventsOwner )
{
	this.Owner = eventsOwner ;
	this._RegisteredEvents = new Object() ;
}

FCKEvents.prototype.AttachEvent = function( eventName, functionPointer )
{
	var aTargets ;

	if ( !( aTargets = this._RegisteredEvents[ eventName ] ) )
		this._RegisteredEvents[ eventName ] = [ functionPointer ] ;
	else
		aTargets.push( functionPointer ) ;
}

FCKEvents.prototype.FireEvent = function( eventName, params )
{
	var bReturnValue = true ;

	var oCalls = this._RegisteredEvents[ eventName ] ;

	if ( oCalls )
	{
		for ( var i = 0 ; i < oCalls.length ; i++ )
			bReturnValue = ( oCalls[ i ]( this.Owner, params ) && bReturnValue ) ;
	}

	return bReturnValue ;
}
