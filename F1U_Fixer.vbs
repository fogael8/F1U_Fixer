Option Explicit

Dim xlApp, xlBook, wsArgs, CurrentDirectory

CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Set xlApp = CreateObject("Excel.Application")

xlApp.Visible = 1

set xlBook = xlApp.Workbooks.Open(CurrentDirectory & "\F1U_Fixer.xlsm")

Set wsArgs = WScript.Arguments

xlApp.Run "F1U_Fixer.xlsm!F1U.F1U", wsArgs(0)

xlBook.Close True

xlApp.Quit

set xlBook = Nothing

Set xlApp = Nothing