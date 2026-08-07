' Argument-ignoring, console-free launcher for the File Transformer ClickOnce
' deployment. Used by the "filetransformer:" URL protocol on the CMSE Dashboard
' (see Register-CMSE-Tools.reg).
'
' Run via wscript.exe, which - unlike cmd.exe - creates no console window, so
' clicking the dashboard link shows nothing but the app itself.
'
' Four traps are handled here, all confirmed by testing on 2026-08-07, and every
' one of them fails SILENTLY. Do not "simplify" them away:
'
'  1. dfshim does NOT strip surrounding quotes. The deployment URL must be
'     passed UNQUOTED or ShOpenVerbApplication just does nothing. Safe here
'     because the percent-encoded URL contains no spaces.
'  2. A URL protocol handler whose shell\open\command has no %1 gets the clicked
'     URL APPENDED by Windows anyway, and dfshim accepts exactly one argument.
'     This script never reads WScript.Arguments, so every argument is discarded.
'  3. dfshim rejects a plain UNC path - the percent-encoded file:// form is
'     mandatory.
'  4. Keeping the URL here rather than in the registry value avoids the shell
'     substituting its own %2 into the "%20" sequences of the URL. VBScript does
'     no percent expansion, so unlike a .cmd the URL needs no %% escaping.
'
' The URL is hard-coded and can never be influenced by the caller, which is what
' keeps the protocol from becoming an argument-injection hole.
'
' Fallback: LaunchFileTransformer.cmd in the repo does the same job for a
' cmd.exe-based handler, if VBScript is ever removed from Windows. It works but
' flashes a console window.

Option Explicit

Dim url, cmd
url = "file://trgfile1/Operations/Software/_Source/ToolBox/" & _
      "File%20Transformer/File%20Transformer.application"
cmd = "rundll32.exe dfshim.dll,ShOpenVerbApplication " & url

On Error Resume Next
CreateObject("WScript.Shell").Run cmd, 0, False
If Err.Number <> 0 Then
    MsgBox "Couldn't start File Transformer." & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & "Open it directly from:" & vbCrLf & _
           "\\trgfile1\Operations\Software\_Source\ToolBox\File Transformer", _
           48, "File Transformer"
End If
