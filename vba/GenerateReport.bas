
Option Explicit

Sub GenerateMISReport()

    Dim pythonPath As String
    Dim projectPath As String
    Dim scriptPath As String
    Dim selectedMonth As String
    Dim command As String
    Dim result As Long

    pythonPath = "C:\Users\satya\Downloads\excel-report-automation\.venv\Scripts\python.exe"

    projectPath = "C:\Users\satya\Downloads\excel-report-automation"

    scriptPath = projectPath & "\src\main.py"

    selectedMonth = InputBox( _
        "Enter reporting month (YYYY-MM):", _
        "Superstore MIS Report", _
        "2017-12" _
    )

    If selectedMonth = "" Then Exit Sub

    If Not selectedMonth Like "####-##" Then
        MsgBox "Please enter month in YYYY-MM format.", vbExclamation
        Exit Sub
    End If

    command = """" & pythonPath & """ """ & scriptPath & _
              """ --month " & selectedMonth

    result = Shell(command, vbNormalFocus)

    If result = 0 Then
        MsgBox "Unable to start Python automation.", vbCritical
    Else
        MsgBox "Python report generation has started." & vbCrLf & _
               "Please wait for the process to finish.", vbInformation
    End If

End Sub
