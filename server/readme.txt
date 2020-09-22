		LONGBOW SERVER - VERSION 1.0
		============================

Installation Guidelines
========================

1) Unzip to C:\LONGBOW
2) Run SETUP.BAT
3) Go into C:\longbow\server\longbow.vbp for the project

From what someone said last time, i'm workin on the interface :)


Understanding The Files In The /CONF Directory
==============================================

	BANIP.INI
		A List of IP's/hostnames that the server should prevent access
		 from.
	DIRCOLS.CFG
		Colour and font settings for the directory listings
	HTTP.CFG
		Misc server settings
	SCRIPTSEC.CFG
		2 Randomly generated security keys allowing scripts containing
		 these keys admin access to the server
	USERS.CFG
		List of the users allowed on the server
		USERNAME, PASSWORD, DIRECTORY, ENABLED?
	VDIR.CFG
		List of the servers virtual directories
	VHOST.CFG
		List of the servers virtual hosts
	DB_ACCESS.CFG
		Can Only Be Modified By The Database Admin Program

	!! It is recommended you make a backup of these files before modifying
		their contents !!

Scripting
=========

	The server allows scripts to be run which have been written using a completely
	wacked out and twisted language i've created :), look at the server code
	(script.cls) and the scripts provided and you should be able to work out
	some of the basic commands :)

	I've included some interesting examples of how the scripting language can be used,
	theres...
		update.html	- Example of file access and a counter

		lovepage.html	- Some shoddy thing someone asked me to do :)

		index.html	- Example of file access and information checking

SEC.IS2k Files
==============
	
	Any directory which contains one of these files is shared, although the level of access
	which is allowed to the directory is chosen by the information inside these files and the
	user account configuration file

