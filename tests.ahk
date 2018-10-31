#NoTrayIcon


assert := new unittest_class()

;Test Excel column incrementing
assert.test(Fn_IncrementExcelColumn("C",1),"D")


;Test title pulling
assert.test(fn_GetTitle("The Mask (1999)"),"The Mask")
assert.test(fn_GetTitle("Déjà Vu"),"Déjà Vu")
assert.test(fn_GetTitle("Amélie (2000)"),"Amélie")
assert.test(fn_GetTitle("OSS 117: Cairo, Nest of Spies"),"OSS 117: Cairo, Nest of Spies")


;Test string similarity
assert.test(stringSimilarity.simpleBestMatch("Smart", ["smarts","marts","clip-art"]),"smarts")


msgbox, % assert.fullreport()
ExitApp

#Include app.ahk
#Include %A_ScriptDir%\lib\unit-testing.ahk\export.ahk