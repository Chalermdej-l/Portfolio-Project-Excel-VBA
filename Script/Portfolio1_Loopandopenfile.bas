Option Explicit

Sub OpenFileandImportData()
    ' Decalre variable to use in the code
    Dim DialogPicker As FileDialog
    Dim FilePicker As Boolean
    Dim files, check As String
    Dim n As Integer

    ' Disable screen update and alert to run code in the backgroup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Call function to clear the data in the sheet
    Call ClearData
    Sheet2.Select
    Set DialogPicker = Application.FileDialog(msoFileDialogFilePicker)


    With DialogPicker
        .InitialFileName = (Environ("Userprofile") & "\Downloads") 'Set Defualt Path for File picker
        .Filters.Clear
        .Filters.Add "Text Or Excel file", "*.txt,*.xls*"
        .Title = ("Please choose Excel or Text File to open")
        .AllowMultiSelect = True
        FilePicker = DialogPicker.Show
    End With

    If Not FilePicker Then 'Check if any file is select or not
        MsgBox ("No File Choosen.")
        Exit Sub
    End If

    If DialogPicker.SelectedItems.count = 1 Then 'Check If choose more than one file
        files = DialogPicker.SelectedItems(1)
        check = Right(files, Len(files) - InStrRev(files, "."))
        If Len(check) = 3 Then ' Check for file type
            Call Read_Text_File(files)
        Else
            Call Read_Excel_File(files)
        End If
    Else
        For n = 1 To DialogPicker.SelectedItems.count
            files = DialogPicker.SelectedItems(n)
            check = Right(files, Len(files) - InStrRev(files, "."))
            If Len(check) = 3 Then ' Check for file type
                Call Read_Text_File(files)
            Else
                Call Read_Excel_File(files)
            End If
        Next n
    End If

' Enable back the screen update and alert
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub Read_Text_File(ByVal File As String) 'Using adoStream to read text file
    Dim adoStream As Object
    Dim var_String As Variant
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "UTF-8"
    adoStream.Open
    adoStream.LoadFromFile File
    var_String = Split(adoStream.ReadText, vbCrLf)
    Range("A1048576").End(xlUp).Offset(1, 0).Resize(UBound(var_String) - LBound(var_String) + 1).Value = Application.Transpose(var_String)
End Sub

Sub Read_Excel_File(ByVal File As String) 'Read Excel File
    Dim openfile As Workbook
    Dim thisfile  As Workbook
    Set thisfile = ThisWorkbook
    Set openfile = Workbooks.Open(File)

    On Error GoTo Error
    openfile.Activate
    If Range("a1").Value = "" Then 'Check if file is empthy
        MsgBox ("File Name " & Right(File, Len(File) - InStrRev(File, "\")) & " is empthy. Skiiping this file.")
        openfile.Close
    Else
        Range("a1", Range("a1").End(xlDown)).EntireRow.Copy
        thisfile.Activate
        Range("A1048576").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        openfile.Close
    End If

Exit Sub

' Error if file can't be insert
Error:
    MsgBox ("Can't Import Data from  " & Right(File, Len(File) - InStrRev(File, "\")) & " . Please check if there are any space left. Or please check the format of the file " & Right(File, Len(File) - InStrRev(File, "\")))
Resume Next
End Sub

Sub ClearData()
Range("A3", Range("A1048576")).EntireRow.Clear
End Sub









