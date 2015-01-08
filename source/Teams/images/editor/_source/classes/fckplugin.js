/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2007 Frederico Caldeira Knabben
 * FCKPlugin Class: Represents a single plugin.
 */

var FCKPlugin = function( name, availableLangs, basePath )
{
	this.Name = name ;
	this.BasePath = basePath ? basePath : FCKConfig.PluginsPath ;
	this.Path = this.BasePath + name + '/' ;

	if ( !availableLangs || availableLangs.length == 0 )
		this.AvailableLangs = new Array() ;
	else
		this.AvailableLangs = availableLangs.split(',') ;
}

FCKPlugin.prototype.Load = function()
{
	// Load the language file, if defined.
	if ( this.AvailableLangs.length > 0 )
	{
		var sLang ;

		// Check if the plugin has the language file for the active language.
		if ( this.AvailableLangs.IndexOf( FCKLanguageManager.ActiveLanguage.Code ) >= 0 )
			sLang = FCKLanguageManager.ActiveLanguage.Code ;
		else
			// Load the default language file (first one) if the current one is not available.
			sLang = this.AvailableLangs[0] ;

		// Add the main plugin script.
		LoadScript( this.Path + 'lang/' + sLang + '.js' ) ;
	}

	// Add the main plugin script.
	LoadScript( this.Path + 'fckplugin.js' ) ;
}