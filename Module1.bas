Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
    On Error GoTo Salir
    
    Dim strValue As String
    Dim mFile As String
    Dim mLongitud As Integer
    Dim mIni  As Long
    
    mFile = ""
    mLongitud = Len(Command)
    mFile = Mid(Command, 2, mLongitud - 2)
    
    If mFile <> "" Then
        Open mFile For Input As #1
        Do While Not EOF(1)
            Input #1, strValue
            If Len(strValue) > 50 Then
                For mIni = 1 To Len(strValue) Step 50
                    Printer.Print Mid(strValue, mIni, 50)
                Next
            Else
                Printer.Print strValue
            End If
        Loop
        Close #1
    End If
    
Salir:
    If Err <> 0 Then
        MsgBox Err & " " & Error
        Printer.KillDoc
    Else
        Printer.EndDoc
    End If
    End
End Sub
