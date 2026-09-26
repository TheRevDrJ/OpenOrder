' OpenOrder - launch the Vite dev server HEADLESS (no window) and detached.
' Copyright (c) 2026 Rev. Dr. Jonathan Mellette (TheRevDrJ). AGPL-3.0 - see LICENSE.
'
' node has no windowless runtime like Python's pythonw.exe, so this hidden-launch
' shim is the node analog: WScript.Shell.Run with window-style 0 (hidden) and
' bWaitOnReturn = False (don't wait). The launched cmd/node gets its own hidden
' console, independent of whatever terminal started OpenOrder.bat - so closing
' that terminal does NOT take the Vite server down. wscript itself is a
' GUI-subsystem host (no console), which is what makes the detachment clean.
'
' Runs from this file's own folder so it works regardless of caller cwd.

Dim fso, sh, scriptDir, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = scriptDir
cmd = "cmd /c npm run dev --prefix frontend > """ & scriptDir & "\frontend.log"" 2>&1"
sh.Run cmd, 0, False
