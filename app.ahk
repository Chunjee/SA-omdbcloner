;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
;Description
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/
; Compares movie titles in an excel file to OMDB/IMDB for extra information which is re-saved to Excel
; 
The_ProjectName := "MovieDBClone"
The_VersionNumb = 1.0.0

;~~~~~~~~~~~~~~~~~~~~~
;Compile Options
;~~~~~~~~~~~~~~~~~~~~~
SetBatchLines -1 ;Go as fast as CPU will allow
#NoTrayIcon ;No tray icon
#SingleInstance Force ;Do not allow running more then one instance at a time
ComObjError(False) ; Ignore any http timeouts


;Dependencies
#Include %A_ScriptDir%\lib
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


; msgbox, % "test1 " Fn_StringSimilarityAttempt("test", "test")
; msgbox, % "test2 " Fn_StringSimilarityAttempt("tasddddasd", "teiiiiiter")
; msgbox, % "test3 " Fn_StringSimilarityAttempt("asdjkjerkquiqwue", "popiiklkpol..liki")


;Read settings.JSON for global settings
FileRead, The_MemoryFile, % A_ScriptDir "\settings.json"
Settings := JSON.parse(The_MemoryFile)
The_MemoryFile := ""


;Create some god vars
AllMoviesDB := []
The_ExcelPath := A_ScriptDir "\" Settings.excelfilename


; Create Excel Object
Excel_obj := ComObjCreate("Excel.Application") ; create Excel Application object
Excel_obj.Visible := true ; make Excel Application invisible
Excel_obj.Workbooks.Open(The_ExcelPath) ;open an existing file


;;Create a blank GUI
GUI()
log.add("GUI launched.")



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
    if (Excel_obj.Range("B" Index).Value != "") {
        AllMoviesDB[Index, "checked"] := true
    }

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


;;Fill GUI with whats being done
INDEX := 0
for key, value in AllMoviesDB { ;;- Each step
    INDEX++
    LV_Add(, AllMoviesDB[A_Index,"rawtext"])
}
;resize columns to fit all text
LV_ModifyCol()

return



PingIMDB:
Loop, % AllMoviesDB.MaxIndex() {
    If (AllMoviesDB[A_Index, "checked"] != true) {
        ;do not check blank titles
        if (AllMoviesDB[A_Index, "title"] = "") {
            continue
        }
        log.add("checking API for the following title:" AllMoviesDB[A_Index, "title"])
        if (AllMoviesDB[A_Index, "year"] ) {
            log.add(AllMoviesDB[A_Index, "title"] " being searched with the year: " AllMoviesDB[A_Index, "year"])
            data := Fn_CheckIMDB(AllMoviesDB[A_Index, "title"], AllMoviesDB[A_Index, "year"])
        } else {
            log.add(AllMoviesDB[A_Index, "title"] " being searched with the without a year")
            data := Fn_CheckIMDB(AllMoviesDB[A_Index, "title"])
        }
        ; Set JSON data to true so it doesn't get anymore
        AllMoviesDB[A_Index, "checked"] := true

        ; Verify that the titles match closely
        similarity := Fn_StringSimilarityAttempt(AllMoviesDB[A_Index, "title"], data.Title)
        if (similarity < Settings.titlematchsimilaritythreshold || !data.Title) {
            ; msgbox, % "titles too dissimiliar"
            return
        }


        excelcoumn := "A"
        excelindex := AllMoviesDB[A_Index, "excelindex"]
        ; Write values to excel
        for key, value in Settings.datapoints {
            thisvalue := Fn_SearchObj(data, value)
            excelcoumn := Fn_IncrementExcelColumn(excelcoumn,1)

            AllMoviesDB[A_Index, value] := thisvalue
            Excel_obj.Range(excelcoumn excelindex).Value := thisvalue
        }

        Excel_obj.ActiveWorkbook.saved := true
        ; Excel_obj.ActiveWorkbook.SaveAs(The_ExcelPath)
        return
    }
}
return

; Array_GUI(AllMoviesDB)
ExitApp, 1





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


Fn_SearchObj(para_obj, para_key)
{
    for l_key, l_value in para_obj {
        ; msgbox, % para_key " - " l_key
        if (para_key = l_key) {
            return l_value
        }
    }
}


Fn_StringSimilarityAttempt(para_string1, para_string2) {
    result := DamerauLevenshteinDistance(para_string1, para_string2)
    if (result >= 0) {
        return result
    } else {
        return 100
    }
}


DamerauLevenshteinDistance(s, t) {
	StringLen, m, s
	StringLen, n, t
	If m = 0
		Return, n
	If n = 0
		Return, m
	d0_0 = 0
	Loop, % 1 + m
		d0_%A_Index% = %A_Index%
	Loop, % 1 + n
		d%A_Index%_0 = %A_Index%
	ix = 0
	iy = -1
	Loop, Parse, s
	{
		sc = %A_LoopField%
		i = %A_Index%
		jx = 0
		jy = -1
		Loop, Parse, t
		{
			a := d%ix%_%jx% + 1, b := d%i%_%jx% + 1, c := (A_LoopField != sc) + d%ix%_%jx%
				, d%i%_%A_Index% := d := a < b ? a < c ? a : c : b < c ? b : c
			If (i > 1 and A_Index > 1 and sc == tx and sx == A_LoopField)
				d%i%_%A_Index% := d < c += d%iy%_%ix% ? d : c
			jx++
			jy++
			tx = %A_LoopField%
		}
		ix++
		iy++
		sx = %A_LoopField%
	}
	Return, d%m%_%n%
}


Fn_IncrementExcelColumn(para_Column,para_IncrementAmmount)
{
    ;Convert Column to a character code from its existing ASCII counterpart
    l_Column := Asc(para_Column)
    l_Column += %para_IncrementAmmount%
        If (l_Column > 122)
        {
        Msgbox, Columns greater than Z are not handled. The program will exit.
        ExitApp
        }
    Return Chr(l_Column)
}
