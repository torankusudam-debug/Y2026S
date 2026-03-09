' run_AI1.vbs (robust)
Option Explicit

Dim jsxPath, ai, i, ok
jsxPath = WScript.Arguments(0)

' 统一：必须同权限运行（不要一个管理员一个非管理员）
' 建议：用同一个权限启动 AI 和运行本 vbs

ok = False

For i = 1 To 12   ' 最多等 12 次（约 1 分钟）
  On Error Resume Next

  ' 先尝试连接已打开的 AI
  Err.Clear
  Set ai = GetObject(, "Illustrator.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set ai = CreateObject("Illustrator.Application")
  End If

  ' 如果连上了，尝试激活并执行
  If Not ai Is Nothing Then
    Err.Clear
    ai.Activate  ' 不要 ai.Visible
    If Err.Number = 0 Then
      Err.Clear
      ai.DoJavaScriptFile jsxPath
      If Err.Number = 0 Then
        ok = True
        Exit For
      End If
    End If
  End If

  ' 失败：等待再试（常见是 AI 没就绪或卡弹窗）
  WScript.Sleep 5000
  On Error GoTo 0
Next

If Not ok Then
  WScript.Echo "FAILED. Please check Illustrator popups / permission / crash."
  WScript.Quit 2
End If

WScript.Echo "OK"
WScript.Quit 0