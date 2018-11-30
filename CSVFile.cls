VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Class to create CsvFile
'Made by RonanVico 03 08 2018 1194015-5925
'Feel Free to change the code, just give the credits for the one who started

Private Const ERR_Default = -1200
Private filename_ As String 'File name
Private header_ As String 'First line of the file
Private useHeader As Boolean 'Use header or not
Private content_ As String  'Content of the csv file
Private separator_ As String 'Separator , Default is ","
Private numColumns_ As Integer 'Number of columns who need to be the same all the lines
Private error_ As Integer 'Eerror number
Private path_ As String 'Path of the file
Private CreatedFile As Boolean 'if file is created or not

Public Enum CSVFileErrors
   ERR_NoError = 0
   err_noheader = (ERR_Default + 1)
   err_HeaderFalse = (ERR_Default + 2)
   err_ColumnsNotMatch = (ERR_Default + 3)
   err_HeaderAfterLines = (ERR_Default + 4)
   err_PathNotFound = (ERR_Default + 5)
   err_NotFileName = (ERR_Default + 6)
   err_FileAlreadExist = (ERR_Default + 7)
   Err_CantChangeHeaderAfter = (ERR_Default + 8)
   ERR_SeparatorUndefined = (ERR_Default + 9)
   err_ChangeSeparatorError = (ERR_Default + 10)

End Enum

'private
Private Sub Class_Initialize()
    useHeader = True
    separator_ = ","
    numColumns_ = -1
    error_ = 0
    content_ = ""
    path_ = "nopath"
    CreatedFile = False
End Sub

Property Get Separator() As String
    Separator = separator_
End Property

Property Let Separator(Character As String)
    'Allow just this 3 character , you cant put more here
    If Character <> "," And Character <> ";" And Character <> "|" Then
        raiseErr (ERR_SeparatorUndefined)
    End If
    
    If InStr(content_, Character) Then
        raiseErr (err_ChangeSeparatorError)
    End If
    
    content_ = Replace(content_, separator_, Character)
    separator_ = Character
    
End Property

Property Get path() As String
    path = path_
End Property

Property Get Filename() As String
    Filename = filename_
End Property

Property Let Filename(name As String)
    Dim path As String, xAux
    
    name = Replace(name, "/", "\")
    For xAux = 0 To UBound(Split(name, "\")) - 1
        path = path & Split(name, "\")(xAux) & "\"
    Next xAux
    
    path = Left(path, Len(path) - 1)
    If VBA.Dir$(path, vbDirectory) = "" Then
        raiseErr (CSVFileErrors.err_PathNotFound)
    End If
    
    path_ = path
    
    If Right(name, 4) <> ".csv" Then
        filename_ = name & ".csv"
    Else
        filename_ = name
    End If
End Property

Property Let useTheHeader(use As Boolean)
    If content_ <> "" Then
        raiseErr (CSVFileErrors.Err_CantChangeHeaderAfter)
    End If
    useHeader = use
End Property


Property Get content() As String
    content = content_
End Property

Public Function createFile(Optional ByVal killFileIfExist As Boolean = True) As Boolean
createFile = False
    If filename_ <> "" Then
        On Error Resume Next
            If killFileIfExist Then Kill filename_
        On Error GoTo 0
        If VBA.Dir$(path_, vbDirectory) = "" Then
            raiseErr (CSVFileErrors.err_PathNotFound)
        ElseIf Dir(filename_) <> "" Then
            raiseErr (CSVFileErrors.err_FileAlreadExist)
        End If
    Else
        raiseErr (err_NotFileName)
    End If
    
'    Dim fl As Object
'    Set fl = CreateObject("Scripting.FileSystemObject").CreateTextFile(filename_, True, True)
'    fl.writeline (content_)
'    fl.Close
'
Dim OutputFileNum, cell
OutputFileNum = FreeFile
  Open filename_ For Output Lock Write As #OutputFileNum
  For Each cell In Split(content_, vbCrLf)
    Print #OutputFileNum, cell
  Next
  Close OutputFileNum
  
  
createFile = True
    
End Function

Property Get header() As String
    header = header_
End Property

Property Let header(head As String)
    Dim xAux, AuxStr
    Dim i As Integer
    If useHeader = False Then
        raiseErr (err_HeaderFalse)
    End If
    
    head = Replace(Replace(head, Chr(13), ""), Chr(10), "")
    
    If content_ = "" Then
        content_ = head & vbNewLine
    ElseIf UBound(Split(header_, separator_)) + 1 <> numColumns_ Then
        raiseErr (CSVFileErrors.err_ColumnsNotMatch)
    Else
         xAux = Split(content_, vbNewLine)
        If UBound(xAux) = 1 Then
            content_ = header & vbNewLine
            numColumns_ = UBound(Split(head, separator_)) + 1
        Else
            AuxStr = head & vbNewLine
            For i = 1 To UBound(xAux)
                AuxStr = AuxStr & xAux(i) & vbNewLine
            Next i
            content_ = AuxStr
        End If
    End If
    
    If numColumns_ = -1 Then
        numColumns_ = UBound(Split(head, separator_)) + 1
    End If
    
    header_ = head
End Property


Public Sub InsertLine(line As String)
    If useHeader = True And header = "" Then
        raiseErr (CSVFileErrors.err_noheader)
    End If
    
    If numColumns_ = -1 Then
        numColumns_ = UBound(Split(line, separator_)) + 1
    End If
    
    If UBound(Split(line, separator_)) + 1 <> numColumns_ Then
        raiseErr (CSVFileErrors.err_ColumnsNotMatch)
    End If
    
    content_ = content_ & line & vbCrLf
End Sub

Public Function getRstQueryData(ByVal query As String)
    Dim csvRst As DAO.Recordset
    Dim i, newHead$, newContent$
    
    Set csvRst = CurrentDb.OpenRecordset(query)
    'clear all the text / header and num of columns
    clearContent
    
    If csvRst.EOF Then
        Exit Function
    End If
    
    For i = 0 To csvRst.Fields.Count - 1
        newHead = newHead & csvRst.Fields(i).name & IIf(csvRst.Fields.Count - 1 <> i, separator_, "")
    Next i
    header = newHead
    
    csvRst.MoveLast
    csvRst.MoveFirst
    
    While Not csvRst.EOF
        newContent = ""
        For i = 0 To csvRst.Fields.Count - 1
            newContent = newContent & csvRst.Fields(i) & IIf(csvRst.Fields.Count - 1 <> i, separator_, "")
        Next i
        
        InsertLine (newContent)
        csvRst.MoveNext
    Wend
    
    On Error Resume Next
        csvRst.Close
        Set csvRst = Nothing
    On Error GoTo 0
End Function

Public Sub clearContent()
    header_ = ""
    content_ = ""
    numColumns_ = -1
End Sub
Public Function getError() As String
    Select Case error_
        Case CSVFileErrors.err_noheader
            getError = "Need to insert a header before content lines if headerUse is true."
        Case CSVFileErrors.err_HeaderFalse
            getError = "You need to enabled property headerUse to get or set a header"
        Case CSVFileErrors.err_ColumnsNotMatch
            getError = "The number of columns inserted " & vbNewLine & "separated by the Separator Character '" & separator_ & "' is not equals the columns who was inserted before" & vbNewLine & "Check if your data in columns do not have the separator '" & separator_ & "'"
        Case CSVFileErrors.err_HeaderAfterLines
            getError = "You cannot insert a head after the lines , the header needs to be the first line!"
        Case CSVFileErrors.err_NotFileName
            getError = "The file needs a name with a Complete Path to create the CSV file."
        Case CSVFileErrors.err_NotFileName
            getError = "Cannot create a file, file already exists"
        Case CSVFileErrors.Err_CantChangeHeaderAfter
            getError = "Cant change this property after iniciate the lines"
        Case CSVFileErrors.ERR_SeparatorUndefined
            getError = "Separator undefined , is just possible to use ',' ';' '|' "
        Case CSVFileErrors.err_ChangeSeparatorError
            getError = "Cannot change separator if you inserted data with the character specified into file for secure purposes"
        Case CSVFileErrors.err_FileAlreadExist
            getError = "File Already Exists and cannot be Deleted because is open"
    End Select
End Function

Private Sub raiseErr(theErr As Integer)
    error_ = theErr
    err.Raise error_, , getError()
End Sub
