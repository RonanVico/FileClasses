Attribute VB_Name = "howtouseCSVFILE"
Option Compare Database
Option Explicit
'Made by ronanvico
'03 08 2018


Public Sub UsingTheClass()
'using the class to insert line by line without using a recordset
    Dim myFile As New CSVFile
    With myFile
        .Filename = CurrentProject.path & "\File1" ' Works with file1.csv too
        .useTheHeader = True
        .header = "COLUMN1,COLUMN2,COLUMN2"
        .InsertLine ("data1,data2,data3")
        .InsertLine ("data1,data2,data3")
        .InsertLine ("data1,data2,data3")
        'if u want to change the header
        .header = "COLUMN3,COLUMN4,COLUMN5)"
    End With
    'This line create the file
    If myFile.createFile Then
        MsgBox "File Created Sucefull!" & myFile.Filename
    End If
End Sub

Public Sub UsingTheClass2()
'Getting a recordset and putting that on a csv
    Dim i As Integer
    'Creating a exemple table to useit in file
    On Error Resume Next
        CurrentDb.Execute "Create TABLE tableExemple (Column1 Text, column2 TEXT ,column5 TEXT)"
        For i = 0 To 5
            CurrentDb.Execute "Insert into tableExemple values ('a','b','c')"
        Next i
    On Error GoTo 0
    Dim myFile As New CSVFile
    With myFile
        .Filename = CurrentProject.path & "\File2" ' Works with file1.csv too
        .getRstQueryData ("SELECT * FROM tableExemple")
        If .createFile Then MsgBox "File created Sucefull!" & .Filename
    End With
End Sub


