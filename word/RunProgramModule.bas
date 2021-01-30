Attribute VB_Name = "RunProgramModule"

' Требует ссылки на Microsoft Scripting Runtime
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Type RunResult
    ExitCode As Integer
    StdOut As String
    StdErr As String
    StdIn As String
    ProcessID As Integer
End Type

Function RunProgram(sProgram As String, Optional sParams As String = "", _
    Optional sCurrentDir As String = "", Optional sStdIn As String = "") As RunResult
    ' sProgram - допускается использование пробелов в пути и в имени файла
    ' sParams - программист сам следит за разделением параметров пробелами (аналогично вызову из cmd.exe)

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    If sCurrentDir <> "" Then
        oShell.CurrentDirectory = sCurrentDir
    End If

    Dim sCmd As String
    sCmd = sProgram & " " & sParams

    Dim oExec As Object
    Set oExec = oShell.Exec(sCmd)
    ' Если есть данные стандартного потока ввода, то передаём их программе
    If sStdIn <> "" And oExec.Status = 0 Then
        oExec.StdIn.WriteLine sStdIn
    End If
    oExec.StdIn.Close
    ' Ожидаем окончания программы
    Do While oExec.Status <> 1
        Sleep 500
    Loop

    RunProgram.ExitCode = oExec.ExitCode
    RunProgram.StdOut = oExec.StdOut.ReadAll()
    RunProgram.StdErr = oExec.StdErr.ReadAll()
    RunProgram.StdIn = sStdIn
    RunProgram.ProcessID = oExec.ProcessID

    Set oShell = Nothing
End Function

Sub TestPerl()
    Dim Result As RunResult
    Result = RunProgram("c:\perl\bin\perl.exe", sParams:="1.pl")
    Result = RunProgram("c:\perl\bin\perl.exe", sParams:="""path with spaces\test.pl""", sStdIn:="12345")
    MsgBox ("STDOUT " + Result.StdOut)
    MsgBox ("STDERR " + Result.StdErr)
End Sub
