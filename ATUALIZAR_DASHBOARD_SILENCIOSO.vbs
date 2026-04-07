' Roda o .bat completamente invisivel - sem janela preta
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "cmd /c """ & "C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\7. Dashboard\ATUALIZAR_DASHBOARD.bat" & """", 0, True
Set shell = Nothing
