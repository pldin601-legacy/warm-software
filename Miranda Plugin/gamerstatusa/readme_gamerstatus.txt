GamerStatus plugin, v1.0.0.4
by mistag <mistag@baze.de> (previously by disq <disq@users.sf.net>)


-- Introduction --
gamerStatus checks if one of the configured programs is running and changes the status accordingly.
It can also perform some other actions then.

One thing, "Notify status change through popup" option won't work unless you have the PopUp
plugin installed.


-- LaunchInfo.txt Support --
LaunchInfo.txt is a file that's created by The All-Seeing Eye (www.udpsoft.com/eye) or GameSpy (www.gamespy.com)
and contains information about the game server the user connects on.
GamerStatus can filter all that information and provides it for your status message with many advanced options.


-- Requirements --
Miranda IM 0.1.2.1+
Recommended:	Popup plugin		[ http://miranda-im.org/download/details.php?action=viewfile&id=299  ]
	     or Popup+ plugin v2.0.3.8+	[ http://miranda-im.org/download/details.php?action=viewfile&id=1170 ]


-- Todo --
 + support for HLSW
 + cut extension of %exename%


-- Known Bugs --
This is a beta version, which means, that it has not been fully tested, yet.
So it is possible that it contains bugs, which are tried to get fixed.


-- Thanks --
Miranda IM team for Miranda.
disq for the basic gamerStatus-plugin
Tornado (author of AwaySys) and MatriX (author of IRC) for some help and support.


-- Contact --
Contact me for bug reports, suggestions and comments or check out the related thread in the Miranda IM forums (if there is one).
forums: http://forums.miranda-im.org/viewforum.php?f=1
E-mail: mistag@baze.de


-- ChangeLog --
v1.0.0.4 beta (2004-11-29)
 + added configurable delay before activation / deactivation (e.g. if a program quits for a short time and then starts again, the status is kept)
 + support for miranda uninstaller [ http://miranda-im.org/download/details.php?action=viewfile&id=969 ]
 + status-change can now be disabled (only to disable popups/sounds)
 + works with Popup Plus v2.0.3.8+
 + some small fixes/tweaks in the options


v1.0.0.3 beta (2004-11-02)
 + FIX: status-messages should now be correct again
 + FIX: occasional wrong load of groupname and awayMessage in proc-group-dialog (work-around)
 + FIX: fixed crashes on miranda shutdown
 + some changes in the options to make it easier and more comfortable

v1.0.0.2 beta (2004-10-15)
 + certain protocols can now be ignored

v1.0.0.1 beta (2004-10-14)
 + protocol-specific status-change (should now work with several protocols)
 + a few other minor changes

v1.0.0.0 beta (2004-10-07)
 + totally recoded (mainly internal effects):
	now development will be much more easier (I hope ^^)
 + introducing groups for all the games/processes:
	makes it easier to configure because you can make groups for games, media players, etc.
 + away-mode is saved in DB:
	popup-ups won't be disabled again if miranda crashed when gamerStatus was in away-mode


v0.9.9.0 (2004-03-17)
 + changed version-count-system
 + first realease from me (mistag) ^^
 + extended LaunchInfo options:
	+ browse for LaunchInfo.txt using the standard windows file-dialog
	+ variable %launchinfo% for a whole text including the desired launchinfo data
	+ option for max. age of launchinfo.txt
	+ option to disable LaunchInfo data if not all data is available
	+ preview of LaunchInfo data
	+ preview of the whole %launchinfo% string
 + option to disable miranda-popups while running a process
 + option to disable miranda-sounds while running a process
 + partly recoded: now properly commented --> better to understand
 + structured the translation file
 + reduced DLL-size to 30% using AgressiveOptimize and Multithreaded DLL


v0.0.1.2 (20030213)
 Memory leak fixed. (EnumProcs.c)
 Added error message to debug ProcessList not working on some systems.

v0.0.1.0 (20021126)
 Translation strings updated
 Clipboard support in away message (%clipboard%)
 Fixed a bug where GamerStatus won't work after launchinfo
 Set status back option didn't work correctly. Hopefully fixed.
 
v0.0.0.5 (20021123)
 Translation strings fixed & updated
 Pure LaunchInfo.txt support!
 Weather protocol support: Won't change weather status to offline
 Now asks for to save changes rather than a dialog to lose changes
 Also asks for to save changes if an entry is modified and the Options dialog is closing
 New variable list dialog
 
v0.0.0.4 (20021117)
 AwaySys 0.2.7.8 support (Will also support AwaySys macros if used with AwaySys) -- Thanks Tornado!
 Added "Plugin Enabled" checkbox in configuration
 "Back to normal" popup is not shown if popups are disabled in the process settings
 After a "save", it now selects the last edited entry
 Clicking "Notify through popup" checkbox wouldn't activate save button. fixed.
 While in away mode, changing status manually to online or offline resets away mode until "the" process ends

v0.0.0.3 (20021113)
 Added LaunchInfo.txt support for game server browsers, new macros
 Added variable list dialog
 
v0.0.0.2 (20021111)
 Got statusmessages working
 Added statusmessage macros: %statdesc% for status description and %exename% for running exe name
 Seperated status and statusmessage settings for each game
 Supressed status change popup
 Translation strings in translation_gamerstatus.txt

v0.0.0.1 (20021106)
 initial release.


-- Disclaimer --
The GamerStatus plugin works fine on my machine, it should work fine on yours too. NO WARRANTY though.
