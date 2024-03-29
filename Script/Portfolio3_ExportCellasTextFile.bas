Option Explicit
'Please note this code use Microsoft Scripting Runtime please enable in the reference first to use this code
Sub ExportWhttoText03()

    ' Seeting variable
    Dim fso As Scripting.FileSystemObject
    Dim tx As Variant
    Set fso = New Scripting.FileSystemObject
    Dim Data As Range
    Dim WHTpath As String
    Dim row, i, s As Integer
    
    Application.ScreenUpdating = False
    Call CreatedFolder 'Check for folder if not exist created one
    Sheet7.Select

    If Range("a4").Value = "" Then 'Check if there is data or not
        Exit Sub
    End If

    ' Set folder to use
    Set tx = fso.OpenTextFile(Environ("Userprofile") & "\Desktop\WHTaxGenerate\TH03.txt", ForWriting, True, True)
        
    
    Range("a4").Select
    row = -1

    Do While ActiveCell.Value <> "" 'Count how many value in a row
        row = row + 1
        ActiveCell.Offset(1, 0).Select
    Loop
    

    For Each Data In Range("a4", Range("a4").Offset(row, 0)) 'Write data into text file
        For i = 1 To 23
            tx.Write Data.Offset(0, i - 1).Value & "|" 'Set a delimiter
        Next i
        tx.WriteLine
    Next Data

    WHTpath = Environ("Userprofile") & "\Desktop\WHTaxGenerate" 'Open The file
    ActiveWorkbook.FollowHyperlink Address:=WHTpath, NewWindow:=True
    tx.Close
        

    Set fso = Nothing
    Application.ScreenUpdating = True

 
End Sub

'Please note this code use Microsoft Scripting Runtime please enable in the reference first to use this code
Sub ExportWhttoText53()
    Dim fso As Scripting.FileSystemObject
    Dim tx As Variant
    Set fso = New Scripting.FileSystemObject
    Dim Data As Range
    Dim WHTpath As String
    Dim row, i, s As Integer
    
    Application.ScreenUpdating = False
    Call CreatedFolder 'Check for folder if not exist created one
    Sheet7.Select

    If Range("a106").Value = "" Then 'Check if there is data or not
        Exit Sub
    End If

    ' Set folder to use
    Set tx = fso.OpenTextFile(Environ("Userprofile") & "\Desktop\WHTaxGenerate\TH53.txt", ForWriting, True, True)
        
    
    Range("a106").Select
    row = -1

    Do While ActiveCell.Value <> "" 'Count how many value in a row
        row = row + 1
        ActiveCell.Offset(1, 0).Select
    Loop
    

    For Each Data In Range("a106", Range("a106").Offset(row, 0)) 'Write data into text file
        For i = 1 To 23
            tx.Write Data.Offset(0, i - 1).Value & "|" 'Set a delimiter
        Next i
        tx.WriteLine
    Next Data

    WHTpath = Environ("Userprofile") & "\Desktop\WHTaxGenerate" 'Open The file
    ActiveWorkbook.FollowHyperlink Address:=WHTpath, NewWindow:=True
    tx.Close
        

    Set fso = Nothing
    Application.ScreenUpdating = True

 
End Sub

Sub CreatedFolder() 'Created Folder on the desktop to put pdf file if folder don't exist created them
    Dim fso As Object
    Dim path As String
    Set fso = CreateObject("scripting.filesystemobject")
    path = Trim(Environ("Userprofile") & "\Desktop\WHTaxGenerate")

    If Not fso.folderexists(path) Then
        fso.createfolder path
    End If

    Set fso = Nothing

End Sub




