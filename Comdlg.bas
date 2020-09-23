Attribute VB_Name = "MCommonDialog"

#If Win32 Then
''@B TOpenFileName32
Private Type TOpenFileName
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to vOwner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to vFilter
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long         ' Tied to vFilterIndex
    lpstrFile As String          ' Tied to vFileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to vFileTitle
    nMaxFileTitle As Long        ' Handle internally
    lpstrInitialDir As String    ' Tied to vInitDir
    lpstrTitle As String         ' Tied to vTitle
    Flags As Long                ' Tied to vFlags
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String        ' Tied to vDefExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (no hooks in Basic)
    lpTemplateName As Long       ' Ignored (no templates in Basic)
End Type
''@E TOpenFileName32
#Else
''@B TOpenFileName1
Private Type TOpenFileName
    lStructSize As Long
    hwndOwner As Integer
    hInstance As Integer
    lpstrFilter As String
    ''@E TOpenFileName1...
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
''@B TOpenFileName2
End Type
''@E TOpenFileName2
#End If

#If Win32 Then
Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As TOpenFileName) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (file As TOpenFileName) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long
#Else
Private Declare Function GetOpenFileName Lib "COMMDLG.DLL " ( _
    file As TOpenFileName) As Integer
Private Declare Function GetSaveFileName Lib "COMMDLG.DLL" ( _
    file As TOpenFileName) As Integer
Private Declare Function GetFileTitle Lib "COMMDLG.DLL" ( _
    ByVal szFile As String, ByVal szTitle As String, _
    ByVal cbBuf As Integer) As Integer
#End If


Public Const cdlOFNReadOnly = &H1
Public Const cdlOFNOverwritePrompt = &H2
Public Const cdlOFNHideReadOnly = &H4
Public Const cdlOFNNoChangeDir = &H8
Public Const cdlOFNFileShowHelp = &H10
Public Const cdlOFNNoValidate = &H100
Public Const cdlOFNAllowMultiselect = &H200
Public Const cdlOFNExtensionDifferent = &H400
Public Const cdlOFNPathMustExist = &H800
Public Const cdlOFNFileMustExist = &H1000
Public Const cdlOFNCreatePrompt = &H2000
Public Const cdlOFNShareAware = &H4000
Public Const cdlOFNNoReadOnlyReturn = &H8000
Public Const cdlOFNNoTestFileCreate = &H10000
' Win95 only
#If Win32 Then
Public Const cdlOFNExplorer = &H80000
Public Const cdlOFNNoDerefenceLinks = &H100000
Public Const cdlOFNLongnames = &H200000
#End If

' Common dialog errors

#If Win32 Then
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long
#Else
Private Declare Function CommDlgExtendedError Lib "COMMDLG.DLL" () As Long
#End If

Public Const CDERR_DIALOGFAILURE = &HFFFF
Public Const CDERR_GENERALCODES = &H0
Public Const CDERR_STRUCTSIZE = &H1
Public Const CDERR_INITIALIZATION = &H2
Public Const CDERR_NOHINSTANCE = &H4
Public Const CDERR_LOADSTRFAILURE = &H5
Public Const CDERR_FINDRESFAILURE = &H6
Public Const CDERR_LOADRESFAILURE = &H7
Public Const CDERR_LOCKRESFAILURE = &H8
Public Const CDERR_MEMALLOCFAILURE = &H9
Public Const CDERR_MEMLOCKFAILURE = &HA
Public Const CDERR_NOHOOK = &HB
Public Const CDERR_REGISTERMSGFAIL = &HC
Public Const PDERR_SETUPFAILURE = &H1001
Public Const PDERR_PARSEFAILURE = &H1002
Public Const PDERR_RETDEFFAILURE = &H1003
Public Const PDERR_LOADDRVFAILURE = &H1004
Public Const PDERR_GETDEVMODEFAIL = &H1005
Public Const PDERR_INITFAILURE = &H1006
Public Const PDERR_NODEVICES = &H1007
Public Const PDERR_NODEFAULTPRN = &H1008
Public Const PDERR_DNDMMISMATCH = &H1009
Public Const PDERR_CREATEICFAILURE = &H100A
Public Const PDERR_DEFAULTDIFFERENT = &H100C
Public Const CFERR_MAXLESSTHANMIN = &H2002
Public Const FNERR_FILENAMECODES = &H3000
Public Const FNERR_SUBCLASSFAILURE = &H3001
Public Const FNERR_INVALIDFILENAME = &H3002
Public Const FNERR_BUFFERTOOSMALL = &H3003


''@B VBGetOpenFileName1
Function VBGetOpenFileName(vFileName As Variant, _
                           Optional vFileTitle As Variant, _
                           Optional vFlags As Variant, _
                           Optional vOwner As Variant, _
                           Optional vFilter As Variant, _
                           Optional vFilterIndex As Variant, _
                           Optional vInitDir As Variant, _
                           Optional vTitle As Variant, _
                           Optional vDefExt As Variant) As Boolean
            
    Dim opfile As TOpenFileName, s As String
With opfile
    .lStructSize = Len(opfile)
    
    ' vFileName must get reference variable to receive result
    ' vFileTitle can get reference variable to receive title
    If IsMissing(vFileTitle) Then vFileTitle = sEmpty
    ' vFlags can get reference variable or constant with bit flags
    If IsMissing(vFlags) Then vFlags = 0
    ' vFilter can take list of filter strings separated by |
    If IsMissing(vFilter) Then vFilter = "All (*.*)| *.*"
    ' vFilterIndex can take initial filter index (one-based)
    If IsMissing(vFilterIndex) Then vFilterIndex = 1
    ' vOwner can take handle of owning window
    If Not IsMissing(vOwner) Then .hwndOwner = vOwner
    ' vInitDir can take initial directory string
    If Not IsMissing(vInitDir) Then .lpstrInitialDir = vInitDir
    ' vDefExt can take default extension
    If Not IsMissing(vDefExt) Then .lpstrDefExt = vDefExt
    ' vTitle can take dialog box title
    If Not IsMissing(vTitle) Then .lpstrTitle = vTitle
''@E VBGetOpenFileName1
    
''@B VBGetOpenFileName2
    ' To make Windows-style filter, replace pipes with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(vFilter)
        ch = Mid$(vFilter, i, 1)
        If ch = "|" Then
            s = s & sNullChr
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & sNullChr & sNullChr
    .lpstrFilter = s
    .nFilterIndex = vFilterIndex
''@E VBGetOpenFileName2

''@B VBGetOpenFileName3
    ' Pad file and file title buffers to maximum path
    s = vFileName & String$(cMaxPath - Len(vFileName), 0)
    .lpstrFile = s
    .nMaxFile = cMaxPath
    s = vFileTitle & String$(cMaxFile - Len(vFileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = cMaxFile
''@E VBGetOpenFileName3
    
''@B VBGetOpenFileName4
    ' Pass in flags, stripping out non-VB flags
    .Flags = vFlags And &H1FF1F
''@E VBGetOpenFileName4

    ' All other fields zero
    '.lpstrCustomFilter = sNullStr
    '.nMaxCustFilter = 0
    '.nFileOffset = 0
    '.nFileExtension = 0
    '.hInstance = hNull            ' No templates
    '.lpTemplateName = pNull
    '.lCustData = pNull            ' No hooks
    '.lpfnHook = pNull
    
''@B VBGetOpenFileName5
    If GetOpenFileName(opfile) Then
        VBGetOpenFileName = True
        vFileName = StrZToStr(.lpstrFile)
        vFileTitle = StrZToStr(.lpstrFileTitle)
        vFlags = .Flags
    Else
        VBGetOpenFileName = False
        vFileName = sEmpty
        vFileTitle = sEmpty
        vFlags = 0
    End If
''@E VBGetOpenFileName5
End With
End Function

Function VBGetSaveFileName(vFileName As String, _
                           Optional vFileTitle As Variant, _
                           Optional vFlags As Variant, _
                           Optional vOwner As Variant, _
                           Optional vFilter As Variant, _
                           Optional vFilterIndex As Variant, _
                           Optional vInitDir As Variant, _
                           Optional vTitle As Variant, _
                           Optional vDefExt As Variant) As Boolean
            
    Dim opfile As TOpenFileName, s As String
With opfile
    .lStructSize = Len(opfile)
    
    ' vFileName must get reference variable to receive result
    ' vFileTitle can get reference variable to receive title in BASE.EXT form
    If IsMissing(vFileTitle) Then vFileTitle = sEmpty
    ' vFlags can get reference variable or constant with bit flags
    If IsMissing(vFlags) Then vFlags = 0
    ' vFilter can take list of filter strings separated by | character
    If IsMissing(vFilter) Then vFilter = "All (*.*)| *.*"
    ' vFilterIndex can take initial filter index (one-based)
    If IsMissing(vFilterIndex) Then vFilterIndex = 1
    ' vOwner can take handle of owning window
    If Not IsMissing(vOwner) Then .hwndOwner = vOwner
    ' vInitDir can take initial directory string
    If Not IsMissing(vInitDir) Then .lpstrInitialDir = vInitDir
    ' vDefExt can take default extension
    If Not IsMissing(vDefExt) Then .lpstrDefExt = vDefExt
    ' vTitle can take dialog box title
    If Not IsMissing(vTitle) Then .lpstrTitle = vTitle
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(vFilter)
        ch = Mid$(vFilter, i, 1)
        If ch = "|" Then
            s = s & sNullChr
        Else
            s = s & ch
        End If
    Next
    s = s & sNullChr & sNullChr
    .lpstrFilter = s
    .nFilterIndex = vFilterIndex

    ' Pad file and file title buffers to maximum path
    s = vFileName & String$(cMaxPath - Len(vFileName), 0)
    .lpstrFile = s
    .nMaxFile = cMaxPath
    s = vFileTitle & String$(cMaxFile - Len(vFileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = cMaxFile
    
    ' Pass in flags, stripping out non-VB flags
    .Flags = vFlags And &H1FF1F

    ' All other fields zero
    '.lpstrCustomFilter = sNull
    '.nMaxCustFilter = 0
    '.nFileOffset = 0
    '.nFileExtension = 0
    '.hInstance = hNull            ' No templates
    '.lpTemplateName = pNull
    '.lCustData = pNull            ' No hooks
    '.lpfnHook = pNull
    
    If GetSaveFileName(opfile) Then
        VBGetSaveFileName = True
        vFileName = StrZToStr(.lpstrFile)
        vFileTitle = StrZToStr(.lpstrFileTitle)
        vFlags = .Flags
    Else
        VBGetSaveFileName = False
        vFileName = sEmpty
        vFileTitle = sEmpty
        vFlags = 0
    End If
End With
End Function

Function VBGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = cMaxPath
    sFileTitle = String$(cMaxPath, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, cMaxPath)
    If cFileTitle Then
        VBGetFileTitle = sEmpty
    Else
        VBGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, sNullChr) - 1)
    End If

End Function

Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, InStr(s, sNullChr) - 1)
End Function


