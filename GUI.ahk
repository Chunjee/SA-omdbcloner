;/--\--/--\--/--\--/--\--/--\
; GUI
;\--/--\--/--\--/--\--/--\--/
GUI()
{
global
;Title
Gui, Font, s14 w70, Arial
Gui, Add, Text, x2 y4 w220 +Center, %The_ProjectName%
Gui, Font, s10 w70, Arial
Gui, Add, Text, x668 y0 w50 +Right, v%The_VersionNumb%

Gui, Font

; Gui,Add,Button,x0 y60 w43 h30 gStart, Start ;gMySubroutine


Gui,Add,ListView,x0 y100 w800 h480 Grid NoSort vGUI_Listview, Movie|Status
	; Gui, Add, ListView, x2 y70 w490 h536 Grid NoSort +ReDraw gDoubleClick vGUI_Listview, #|Status|RC|Name|Race|

Gui,Show,h600 w800, %The_ProjectName%


;Menu
Menu, FileMenu, Add, R&estart`tCtrl+R, Menu_File-Restart
Menu, FileMenu, Add, E&xit`tCtrl+Q, Menu_File-Quit
Menu, MenuBar, Add, &File, :FileMenu  ; Attach the sub-menu that was created above

Menu, HelpMenu, Add, &About, Menu_About
Menu, HelpMenu, Add, &Documentation`tCtrl+H, Menu_Documentation
Menu, MenuBar, Add, &Help, :HelpMenu

Gui, Menu, MenuBar
Return

;Menu Shortcuts
Menu_Documentation:
Run https://github.com/Chunjee/SA-omdbcloner
Return

Menu_About:
Msgbox, Please see the documentation
Return

Menu_File-Restart:
log.add("App being restarted by the user")
log.finalizeLog(The_ProjectName . " log completed.")
Reload

Menu_File-Quit:
GuiClose:
sb_exitapp()
}


sb_exitapp()
{
	global ;needs global acess to access log object

	log.add("GUI closed")
	log.finalizeLog(The_ProjectName . " log completed.")
	ExitApp
}