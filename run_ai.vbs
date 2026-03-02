' run_ai.vbs  (FIX: escape backslashes/quotes for DoJavaScript)
Option Explicit

Dim jsxPath, dataStr, progId
progId = "Illustrator.Application"

If WScript.Arguments.Count < 2 Then
  WScript.Echo "Usage: cscript //nologo run_ai.vbs ""path\script.jsx"" ""cx;path_or_folder"""
  WScript.Quit 2
End If

jsxPath = WScript.Arguments(0)
dataStr = WScript.Arguments(1)

WScript.Echo "DBG jsxPath=" & jsxPath
WScript.Echo "DBG dataStr=" & dataStr
WScript.Echo "DBG progId=" & progId

Function JsEscape(ByVal s)
  ' JS string escape (for content inside "...")
  s = Replace(s, "\", "\\")        ' IMPORTANT: windows path -> JS path
  s = Replace(s, """", "\""")      ' escape quotes
  s = Replace(s, vbCrLf, "\n")
  s = Replace(s, vbCr, "\n")
  s = Replace(s, vbLf, "\n")
  JsEscape = s
End Function

On Error Resume Next
Dim ai : Set ai = CreateObject(progId)
If Err.Number <> 0 Then
  WScript.Echo "ERR: CreateObject failed. " & Err.Number & " " & Err.Description
  WScript.Quit 3
End If
On Error GoTo 0

' 可选：让 AI 可见，避免有时后台初始化不完整
On Error Resume Next
ai.Visible = True
ai.UserInteractionLevel = 2 ' DontDisplayAlerts
On Error GoTo 0

WScript.Sleep 500

Dim js
js = ""
js = js & "var __args=[""" & JsEscape(dataStr) & """];" & vbLf
js = js & "try{" & vbLf
js = js & "  $.evalFile(""" & JsEscape(jsxPath) & """);" & vbLf
js = js & "  if (typeof main !== 'function') { 'ERR:main_not_found'; }" & vbLf
js = js & "  else { String(main(__args)); }" & vbLf
js = js & "}catch(e){ 'ERR:'+String(e); }"

On Error Resume Next
Dim ret : ret = ai.DoJavaScript(js)
If Err.Number <> 0 Then
  WScript.Echo "ERR: DoJavaScript failed. err=" & Err.Number & " desc=" & Err.Description
  WScript.Quit 4
End If
On Error GoTo 0

WScript.Echo "RET: " & ret
WScript.Quit 0