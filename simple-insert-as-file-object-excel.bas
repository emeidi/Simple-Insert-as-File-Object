Option Compare Text
Option Explicit

Dim previousFolder As String

Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const MAX_PATH As Long = 260
 
Type BrowseInfo
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszINSTRUCTIONS As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
 
Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
 
Declare Function SHGetPathFromIDListA Lib "shell32.dll" ( _
ByVal pidl As Long, _
ByVal pszBuffer As String) As Long
Declare Function SHBrowseForFolderA Lib "shell32.dll" ( _
lpBrowseInfo As BrowseInfo) As Long

Function emeidiBrowseFile() As Variant
    Dim fd As FileDialog
    Dim Path As String
    Dim vrtSelectedItem As Variant
    Dim Files() As Variant
    Dim numItems As Integer
    Dim strDebug As String
    Dim i As Integer
    Dim FirstRun As Boolean
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ReDim Files(0)
    
    With fd
        If Len(previousFolder) > 0 Then
            'Show previously used folder
            fd.InitialFileName = previousFolder
            'fd.Show
        End If
        
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                Path = vrtSelectedItem
                'Debug.Print "Path: " & Path & "(" & vrtSelectedItem
                
                numItems = UBound(Files)
                ReDim Preserve Files(numItems + 1)
                Files(numItems) = vrtSelectedItem
            Next vrtSelectedItem
        Else
        End If
    End With
    
    Set fd = Nothing
    
    previousFolder = emeidiGetDirname(Path)
    'MsgBox "previousFolder: " & previousFolder
    
    For i = 0 To (UBound(Files) - 1)
        strDebug = strDebug & ", " & i & ": " & Files(i)
    Next
    'MsgBox (strDebug)
    
    'emeidiBrowseFile = Path
    emeidiBrowseFile = Files
End Function

Function emeidiGetFileNameFromPath(Path As Variant)
    emeidiGetFileNameFromPath = Right(Path, Len(Path) - InStrRev(Path, "\"))
End Function

Function emeidiCleanFileName(Filename As String)
    'Might be a security risk ... Diable this if in doubt
    emeidiCleanFileName = Left(Filename, InStrRev(Filename, ".") - 1)
End Function

Function emeidiGetDirname(Path As String)
    emeidiGetDirname = Left(Path, InStrRev(Path, "\"))
End Function

Function emeidiGetIconFilename(Extension)
    Dim Icon(0 To 1) As Variant
    
    Extension = LCase(Extension)
    'MsgBox "Looking for: " & Extension
    
    Select Case Extension
        Case "doc", "docx"
            Icon(0) = Application.Path & "\winword.exe"
            Icon(1) = 1
        Case "xls", "xlsx", "csv"
            Icon(0) = Application.Path & "\excel.exe"
            Icon(1) = 1
        Case "ppt", "pptx"
            Icon(0) = Application.Path & "\powerpnt.exe"
            Icon(1) = 1
        Case "mdb", "accdb"
            Icon(0) = Application.Path & "\powerpnt.exe"
            Icon(1) = 1
        Case "gif", "jpg", "jpeg", "bmp", "tif", "tiff"
            Icon(0) = Application.Path & "\ois.exe"
            Icon(1) = 1
        Case "txt"
            ' Hardlinks, ugly ... but easiest way
            Icon(0) = "C:\WINDOWS\system32\notepad.exe"
            Icon(1) = 0
        Case "zip"
            ' Hardlinks, ugly ... but easiest way
            Icon(0) = "C:\Program Files\Winzip\winzip32.exe"
            Icon(1) = 0
        Case "pdf"
            ' Hardlinks, ugly ... but easiest way
            Icon(0) = "C:\Program Files\Adobe\Reader 9.0\Reader\AcroRd32.dll"
            Icon(1) = 5
    End Select
    
    emeidiGetIconFilename = Icon
End Function

Function emeidiGetExtensionFromFileName(Filename)
    Dim dotPos As Integer
    Dim Range As Integer
    
    dotPos = InStrRev(Filename, ".")
    Range = Len(Filename) - dotPos
    
    emeidiGetExtensionFromFileName = Right(Filename, Range)
End Function

Function InStrRev(ByVal pStr As String, pItem As String) As Integer
    Dim i As Integer, n As Integer, tLen As Integer

    n = 0
    tLen = Len(pItem)
    
    For i = Len(RTrim(pStr)) To 1 Step -1

        If Mid(pStr, i, tLen) = pItem Then
            n = i
            Exit For
        End If
    Next i

    InStrRev = n
End Function

Sub emeidiInsertDocs()
    Dim Prompt          As String
    Dim Title           As String
    Dim Path            As Variant
    Dim Filename        As String
    Dim MyResponse      As VbMsgBoxResult
    Dim IconLabel       As String
    Dim doc             As Object
    Dim shapeType       As String
    Dim Icon            As Variant
    Dim Extension       As String
    Dim Count           As Integer
    Dim Files()         As Variant
    Dim i               As Integer
    Dim iTop            As Integer
    Dim iLeft           As Integer
     
    'WordBasic.DisableAutoMacros True
    
    Files = emeidiBrowseFile()
    
    If UBound(Files) < 1 Then
        Prompt = "You didn't select a file. The procedure has been canceled."
        Title = "Procedure Canceled"
        MsgBox Prompt, vbCritical, Title
        
        GoTo Canceled:
    End If
    
    For i = 0 To UBound(Files) - 1
        Path = Files(i)
        
        If Path = "" Then
            Prompt = "Path is empty for "
            Title = "Procedure Canceled"
            MsgBox Prompt, vbCritical, Title
            
        End If
         
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        
        Filename = emeidiGetFileNameFromPath(Path)
        IconLabel = emeidiCleanFileName(Filename)
		
        If UBound(Files) < 4 Then
            ' If the user has selected less than 3 files, display the nameing dialog. Otherwise assume he just wants to batch insert
            IconLabel = InputBox(Prompt:="File Label:", Default:=IconLabel)
        End If

         
         If IconLabel = "" Then
            Prompt = "Empty Label not allowed"
            Title = "Procedure Canceled"
            MsgBox Prompt, vbCritical, Title
    GoTo Canceled:
        End If
         
         Extension = emeidiGetExtensionFromFileName(Filename)
         'MsgBox "Filename:" & Filename & ", Extension: " & Extension
         
         Icon = emeidiGetIconFilename(Extension)
         'MsgBox Icon(0) & ", " & Icon(1)
         
         ' Excel stacks the objects ontop of each other; add some distance to it
         iTop = i * 50
         iLeft = i * 50
         
         Set doc = ActiveWorkbook.ActiveSheet
         With doc.Shapes
            With .AddOLEObject(Filename:=Path, Link:=False, DisplayAsIcon:=True, IconFileName:=Icon(0), IconIndex:=Icon(1), IconLabel:=IconLabel, Left:=iLeft, Top:=iTop)
                'Not available in Excel
                '.ConvertToInlineShape
            End With
         End With
    Next
     
     
Canceled:
    'WordBasic.DisableAutoMacros False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
     
End Sub