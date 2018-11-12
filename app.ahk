;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
;Description
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/
; Compares movie titles in an excel file to OMDB/IMDB for extra information which is re-saved to Excel
; 
The_ProjectName := "MovieDBClone"
The_VersionNumb = 1.0.9

;~~~~~~~~~~~~~~~~~~~~~
;Compile Options
;~~~~~~~~~~~~~~~~~~~~~
SetBatchLines -1 ;Go as fast as CPU will allow
#NoTrayIcon ;No tray icon
#SingleInstance Force ;Do not allow running more then one instance at a time
ComObjError(False) ; Ignore any http timeouts


;Dependencies
#Include %A_ScriptDir%\lib
#Include %A_ScriptDir%\lib\util-misc.ahk\export.ahk
#Include %A_ScriptDir%\lib\logs.ahk\export.ahk
#Include %A_ScriptDir%\lib\json.ahk\export.ahk
#Include %A_ScriptDir%\lib\sort-array.ahk\export.ahk
#Include %A_ScriptDir%\lib\string-similarity.ahk\export.ahk

;For Debug Only
; #Include %A_ScriptDir%\lib\unit-testing.ahk\export.ahk

;Modules
#Include %A_ScriptDir%
#Include GUI.ahk


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
log.add("Opened Excel file: " The_ExcelPath)

;;Create a blank GUI
GUI()
log.add("GUI launched.")



; Read all lines in the excel
Index := 1
KeepReading := true
While (KeepReading = true) {

    rawtext := Excel_obj.Range("A" Index).Value
    TheMovie_title := fn_GetTitle(rawtext)
    TheMovie_year := Fn_QuickRegEx(rawtext,"\((\d+)\)")
    if (Excel_obj.Range("B" Index).Value != "") {
        AllMoviesDB[Index, "checked"] := true
    }

    if (TheMovie_title) {
        
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
AllMoviesDB.RemoveAt(1, 1) ;because reading excel always seems to read the first header as the first line; remove it. 

SetTimer, PingIMDB, 1000
; SetTimer, WriteJSONDB, 30000
; Array_GUI(AllMoviesDB)


;;Fill GUI with whats being done
sleep, 2000
INDEX := 0
for key, value in AllMoviesDB { ;;- Each step
    INDEX++
    LV_Add(, AllMoviesDB[A_Index,"rawtext"])
}
;resize columns to fit all text
LV_ModifyCol()
return


WriteJSONDB:
sb_ExportJSON(AllMoviesDB, A_ScriptDir "\AllMoviesDB.json")
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
        similarity := stringSimilarity.compareTwoStrings(AllMoviesDB[A_Index, "title"], data.Title)
        if (similarity <= Settings.titlematchsimilaritythreshold || !data.Title) {
            msgbox, , The_ProjectName, % "When searching for " AllMoviesDB[A_Index, "title"] "; the return value ''" data.Title "'' was rated " similarity " which is below the settings threshold of " Settings.titlematchsimilaritythreshold "`n`nConsider lowering the threshold or set a negative number to accept all results", 10
            return
        }


        excelcoumn := "A"
        excelindex := AllMoviesDB[A_Index, "excelindex"]
        D_Index := A_Index
        ; Write values to excel
        for key, value in Settings.datapoints {
            thisvalue := Fn_SearchObjWithKey(data, value)
            excelcoumn := Fn_IncrementExcelColumn(excelcoumn,1)

            AllMoviesDB[D_Index, value] := thisvalue
            Excel_obj.Range(excelcoumn excelindex).Value := thisvalue
        }

        Excel_obj.ActiveWorkbook.saved := true
        ; Excel_obj.ActiveWorkbook.SaveAs(The_ExcelPath)

        ;Write to json file if at last line:
        if (A_Index = AllMoviesDB.MaxIndex()) {
            Gosub, WriteJSONDB
        }
        return
    }
}
return

; Array_GUI(AllMoviesDB)
ExitApp, 1





;/--\--/--\--/--\--/--\--/--\
; Subroutines
;\--/--\--/--\--/--\--/--\--/

sb_ExportJSON(para_DataObj, para_Filepath)
{
    global JSON

    l_memoryfile := JSON.stringify(para_DataObj)
    FileDelete, %para_Filepath%
    FileAppend, %l_memoryfile%, %para_Filepath%
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
        clipboard := endpoint
    }

    ; clipboard := endpoint
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


Fn_SearchObjWithKey(para_obj, para_key)
{
    for l_key, l_value in para_obj {
        ; msgbox, % para_key " - " l_key
        if (para_key = l_key) {
            return l_value
        }
    }
}


Fn_IncrementExcelColumn(para_Column, para_IncrementAmmount)
{
    ;Convert Column to a character code from its existing ASCII counterpart
    l_Column := Asc(para_Column)
    l_Column += %para_IncrementAmmount%
        If (l_Column > 122)
        {
            Msgbox, Columns greater than Z are not handled. The program will exit.
            ExitApp
        }
    return Chr(l_Column)
}


Fn_SDCSimilarity(para_string1,para_string2) {
    ;Sørensen-Dice coefficient
    vCount := 0
    oArray := {}
    oArray := {base:{__Get:Func("Abs").Bind(0)}} ;make default key value 0 instead of a blank string
    Loop, % vCount1 := StrLen(para_string1) - 1
        oArray["z" SubStr(para_string1, A_Index, 2)]++
    Loop, % vCount2 := StrLen(para_string2) - 1
        if (oArray["z" SubStr(para_string2, A_Index, 2)] > 0)
        {
            oArray["z" SubStr(para_string2, A_Index, 2)]--
            vCount++
        }
    vDSC := (2 * vCount) / (vCount1 + vCount2)
    ; MsgBox, % vCount " " vCount1 " " vCount2 "`r`n" vDSC
    return Round(vDSC,2)
}


fn_GetTitle(para_rawtext)
{
    global 

    text1 := Fn_QuickRegEx(para_rawtext,"(.+)\(")
    if (text1) {
        return Trim(text1," ")
    }
    text2 := Fn_QuickRegEx(para_rawtext,"(.+)")
    if (text2) {
        return Trim(text2," ")
    }
    return false
}
