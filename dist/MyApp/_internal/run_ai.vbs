' run_ai.vbs
Option Explicit

Dim jsxPath, dataStr, progId

If WScript.Arguments.Count < 2 Then
  WScript.Echo "Usage: cscript //nologo run_ai.vbs ""path\script.jsx"" ""cx;aiFilePath"" [progId]"
  WScript.Quit 2
End If

jsxPath = WScript.Arguments(0)
dataStr = WScript.Arguments(1)
progId = "Illustrator.Application"
If WScript.Arguments.Count >= 3 Then
  progId = WScript.Arguments(2)
End If

WScript.Echo "DBG jsxPath=" & jsxPath
WScript.Echo "DBG dataStr=" & dataStr
WScript.Echo "DBG progId=" & progId

Function JsEscape(ByVal s)
  ' Escape for JS string content inside "..."
  s = Replace(s, "\", "\\")
  s = Replace(s, """", "\""")
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

On Error Resume Next
ai.Visible = True
ai.UserInteractionLevel = 2 ' aiDontDisplayAlerts
On Error GoTo 0

WScript.Sleep 500

Dim js
js = ""
js = js & "var __arg=""" & JsEscape(dataStr) & """;" & vbLf
js = js & "var __ret='';" & vbLf
js = js & "try{" & vbLf
js = js & "  $.evalFile(""" & JsEscape(jsxPath) & """);" & vbLf
js = js & "  if (typeof main !== 'function') { __ret = '500;ERR;main_not_found'; }" & vbLf
js = js & "  else { __ret = String(main(__arg)); }" & vbLf
js = js & "}catch(e){ __ret = '500;ERR;' + String(e) + ((e && e.line) ? (' @line=' + e.line) : ''); }" & vbLf
js = js & "__ret;"

On Error Resume Next
Dim ret : ret = ai.DoJavaScript(js)
If Err.Number <> 0 Then
  WScript.Echo "ERR: DoJavaScript failed. err=" & Err.Number & " desc=" & Err.Description
  WScript.Quit 4
End If
On Error GoTo 0

WScript.Echo "RET: " & ret
WScript.Quit 0
