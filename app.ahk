;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
;Description
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/
; Performs Start of Day on the QA systems
; 
The_ProjectName := "MovieDBClone"
The_VersionNumb = 0.3.0

;~~~~~~~~~~~~~~~~~~~~~
;Compile Options
;~~~~~~~~~~~~~~~~~~~~~
SetBatchLines -1 ;Go as fast as CPU will allow
#NoTrayIcon ;No tray icon
#SingleInstance Force ;Do not allow running more then one instance at a time
ComObjError(False) ; Ignore any http timeouts

;Hide CMD window
DllCall("AllocConsole")
WinHide % "ahk_id " DllCall("GetConsoleWindow", "ptr")


;Dependencies
#Include %A_ScriptDir%\functions
#Include util_misc.ahk
#Include util_arrays.ahk
#Include json.ahk

;For Debug Only
; #Include ahk-unittest.ahk


;Classes
#Include %A_ScriptDir%\classes
#Include Logging.ahk


;Modules
#Include %A_ScriptDir%
#Include GUI.ahk


Sb_InstallFiles() ;Install included files and make any directories required

;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; StartUp
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

;;Creat Logging obj
log := new Log_class(The_ProjectName "-" A_YYYY A_MM A_DD, A_ScriptDir "\LogFiles")
log.maxSizeMBLogFile_Default := 99 ;Set log Max size to 99 MB
log.application := The_ProjectName
log.preEntryString := "%A_NowUTC% -- "
; log.postEntryString := "`r"
log.initalizeNewLogFile(false, The_ProjectName " v" The_VersionNumb " log begins...`n")
log.add(The_ProjectName " launched from user " A_UserName " on the machine " A_ComputerName ". Version: v" The_VersionNumb)

;;Create a blank GUI
GUI()
log.add("GUI launched.")

; Create Excel Object
Excel_obj := ComObjCreate("Excel.Application") ; create Excel Application object
Excel_obj.Visible := true ; make Excel Application invisible
Excel_obj.Workbooks.Open(A_ScriptDir . "\ExampleBook.xlsx") ;open an existing file
; Excel_obj.ActiveWorkbook.SaveAs(A_ScriptDir . "\ExampleBook.xlsx")
; Excel_obj.Workbooks.Add ; create a new workbook (oWorkbook := oExcel.Workbooks.Add)



;Read settings.JSON for global settings
FileRead, The_MemoryFile, % A_ScriptDir "\settings.json"
Settings := JSON.parse(The_MemoryFile)
The_MemoryFile := ""


;Create some god vars
AllMoviesDB := []


; Read all lines in the excel
Index := 1
KeepReading := true
While (KeepReading = true) {
    if (Index = 1) {
        ; do not read the column header
        Index++
        continue
    }

    rawtext := Excel_obj.Range("A" Index).Value
    TheMovie_title := Fn_QuickRegEx(rawtext,"([\w ]+)")
    TheMovie_year := Fn_QuickRegEx(rawtext,"\((\d+)\)")
    if (TheMovie_title != "null") {

    } else {
        KeepReading := false
        break
    }
    AllMoviesDB[Index,"rawtext"] := rawtext
    AllMoviesDB[Index,"excelindex"] := Index
    AllMoviesDB[Index,"title"] := TheMovie_title
    if (TheMovie_year != "null") {
        AllMoviesDB[Index,"year"] := TheMovie_year
    }
    Index++
}

SetTimer, PingIMDB, 1000
; Array_GUI(AllMoviesDB)
return



PingIMDB:
Loop, % AllMoviesDB.MaxIndex() {
    If (AllMoviesDB[A_Index, "checked"] != true) {
        if (AllMoviesDB[A_Index, "year"] ) {
            data := Fn_CheckIMDB(AllMoviesDB[A_Index, "title"], AllMoviesDB[A_Index, "year"])
        } else {
            data := Fn_CheckIMDB(AllMoviesDB[A_Index, "title"])
        }

        ;Set JSON data to true so it doesn't get anymore
        AllMoviesDB[A_Index, "checked"] := true

        excelindex := AllMoviesDB[A_Index, "excelindex"]
        msgbox, % data.Actors
        Excel_obj.Range("B" excelindex).Value := data.Actors
        Excel_obj.ActiveWorkbook.saved := true
        Excel_obj.ActiveWorkbook.SaveAs(A_ScriptDir . "\ExampleBook.xlsx")
    }
}
Excel_obj.Quit
return

; Array_GUI(AllMoviesDB)
ExitApp, 1



;Select the appropriot script or default
ReReadFile:
; GuiControl, Text , MainDropdown_List, % scriptdropdowns
LV_Delete() ;clear the listview



;;Fill GUI with whats to be done and assign index values to each item
INDEX := 0
for key, value in AllSteps_JSON { ;;- Each step
    INDEX++
    ; for key2, value2 in value.machines { ;;- Each machine
    ; }
}
;resize columns to fit all text
LV_ModifyCol()



;; Process each step after user presses Start
Start:
;Grab the selected script
GuiControlGet, OutputVar ,, MainDropdown_List, Text
INDEX := 0
log.add(The_ProjectName " Running " SelectedScript " with the " A_UserName " credentials from " A_ComputerName)
;;Create new Remote_Control class
RemoteControl := New RemoteControl_Class("TOP")

for key, value in AllItems_JSON { ; - Each step
; INDEX++
;     LV_Modify(INDEX,,,,,"IN PROGRESS")
}
Return





;/--\--/--\--/--\--/--\--/--\
; Subroutines
;\--/--\--/--\--/--\--/--\--/

;Create Directory and install needed file(s)
Sb_InstallFiles()
{
    ; FileCreateDir, %A_ScriptDir%\data\
}





;/--\--/--\--/--\--/--\--/--\
; Functions
;\--/--\--/--\--/--\--/--\--/

Fn_CheckIMDB(para_movietitle, para_year := "null")
{
    global 
    ; global Settings
    
    endpoint := Settings.endpoint "apikey=" Settings.key "&t=" para_movietitle
    if (para_year != "null") {
        endpoint := endpoint "&y=" para_year
    }
    if (Settings.optionals) {
        endpoint := endpoint Settings.optionals
    }

    clipboard := endpoint
    ; msgbox, % endpoint
    http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    http.Open("Get", endpoint, False)
    ; http.SetRequestHeader("Accept", "application/json")
    http.Send()
    ; msgbox, % http.ResponseText

    ;parse results and return if valid
    l_data := JSON.parse(http.ResponseText)
    if (l_data.Title) {
        return % l_data
    } else {
        return false
    }
    ;Save Raw just for later viewing
    
}
