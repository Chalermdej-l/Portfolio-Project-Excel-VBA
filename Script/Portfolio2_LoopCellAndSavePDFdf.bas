Option Explicit

Sub Savepdf() 'Save file in a pdf format

    ' Declare variable
    Dim result As Range
    Dim Refno As String
    Dim folder As String
    Dim number As Integer
    number = 0
    Sheet3.Activate

    ' Accept input Ref no.
    Refno = Application.InputBox(Prompt:="Put in Ref No.", Title:="Save PDF Multiple file", Type:=1)
    Application.ScreenUpdating = False
    If Len(Refno) <> 7 Then 'Check if choose the correct value
        MsgBox ("Plase put in Ref no.")
        Exit Sub
    End If
    Call CreatedFolder
    Set result = Range("B:B").Find(what:=Refno, LookIn:=xlValues, MatchCase:=True, lookat:=xlWhole)
    If result Is Nothing Then
        MsgBox ("Ref doesn't exist or input wrong no.")
        Exit Sub
    Else
        result.Select
        Do While ActiveCell.Value <> "" 'Loop downward to save all document as pdf
            ActiveCell.Copy
            Sheet4.Range("f2").PasteSpecial xlPasteValues
            Sheet4.ExportAsFixedFormat xlTypePDF, Environ("Userprofile") & "\Desktop\Print\" & "TH01 - " & ActiveCell.Value & ".pdf"
            ActiveCell.Offset(1, 0).Select
        Loop
    End If
    Application.ScreenUpdating = True

    ' Open the folder the file output in
    folder = Environ("Userprofile") & "\Desktop\Print"
    ActiveWorkbook.FollowHyperlink Address:=folder, NewWindow:=True 'Open the print folder on the desktop
End Sub

Sub CreatedFolder() 'Created Folder on the desktop to put pdf file if folder don't exist created them
    Dim fso As Object
    Dim path As String
    Set fso = CreateObject("scripting.filesystemobject")
    path = Environ("Userprofile") & "\Desktop\Print"

    If Not fso.folderexists(path) Then
        fso.createfolder path
    End If

    Set fso = Nothing

End Sub
























