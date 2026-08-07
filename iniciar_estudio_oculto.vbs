' Inicia Estudio Contable en segundo plano (sin ventana ni barra de tareas).
Option Explicit
Dim sh, fso, root, ps1, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
root = fso.GetParentFolderName(WScript.ScriptFullName)
ps1 = root & "\iniciar_estudio_oculto.ps1"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1 & """"
sh.CurrentDirectory = root
' 0 = oculto, False = no esperar
sh.Run cmd, 0, False
