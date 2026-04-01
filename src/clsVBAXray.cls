VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAXray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  ||||||||||||||||||||||||||||||||||||||||||||
'  |||||||||||||    vbaXRay    ||||||||||||||||
'  ||||||||||||||||||||||||||||||||||||||||||||
'
'    AUTHOR:   Kallun Willock
'
'    URL:      https://github.com/KallunWillock/vbaXray
'
'    NOTES:    The class requires a vbaProject.bin file, which can be obtained
'              from the container .XLSM, .DOCM, etc. file.
'
'    LICENSE:  MIT
'
'    VERSION:  1.0

Option Explicit

#If Win64 Then
  Private Const PTR_SIZE As Long = 8
#Else
  Private Const PTR_SIZE As Long = 4
#End If

#If VBA7 Then
  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
  Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
  Private Declare PtrSafe Function apiCreateDirectoryW Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
  Private Declare PtrSafe Function apiGetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As LongPtr) As Long
  Private Declare PtrSafe Function StgOpenStorage Lib "ole32" (ByVal pwcsName As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare PtrSafe Function CreateILockBytesOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ppLkbyt As LongPtr) As Long
  Private Declare PtrSafe Function StgOpenStorageOnILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ppstgOpen As stdole.IUnknown) As Long
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
  Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
  Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
#Else
  Private Enum LongPtr
    [_]
  End Enum
  Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
  Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
  Private Declare Function apiCreateDirectoryW Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
  Private Declare Function apiGetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As LongPtr) As Long
  Private Declare Function StgOpenStorage Lib "ole32" (ByVal pwcsName As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare Function CreateILockBytesOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ppLkbyt As LongPtr) As Long
  Private Declare Function StgOpenStorageOnILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ppstgOpen As stdole.IUnknown) As Long
  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
  Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
  Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
#End If

Private Type UUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Type FILETIME
  dwLowDateTime  As Long
  dwHighDateTime As Long
End Type

Private Type STATSTG
  pwcsName          As LongPtr
  dwType            As Long
  cbSize            As Currency
  mtime             As FILETIME
  ctime             As FILETIME
  atime             As FILETIME
  grfMode           As Long
  grfLocksSupported As Long
  clsid             As UUID
  grfStateBits      As Long
  reserved          As Long
End Type

Private Enum IStorageInterface
  IStg_CreateStream = 3
  IStg_OpenStream = 4
  IStg_CreateStorage = 5
  IStg_OpenStorage = 6
  IStg_CopyTo = 7
  IStg_MoveElementTo = 8
  IStg_Commit = 9
  IStg_Revert = 10
  IStg_EnumElements = 11
  IStg_DestroyElement = 12
  IStg_RenameElement = 13
  IStg_SetElementTimes = 14
  IStg_SetClass = 15
  IStg_SetStateBits = 16
  IStg_stat = 17
End Enum

Private Enum IEnumSTATSTGInterface
  IEnum_Next = 3
  IEnum_Skip = 4
  IEnum_Reset = 5
  IEnum_Clone = 6
End Enum

Private Enum IStreamInterface
  IStm_Read = 3
  IStm_Write = 4
  IStm_Seek = 5
  IStm_SetSize = 6
  IStm_CopyTo = 7
  IStm_Commit = 8
  IStm_Revert = 9
  IStm_LockRegion = 10
  IStm_UnlockRegion = 11
  IStm_stat = 12
  IStm_Clone = 13
End Enum

Private Enum STGTY
  STGTY_STORAGE = 1
  STGTY_STREAM = 2
  STGTY_LOCKBYTES = 3
  STGTY_PROPERTY = 4
End Enum

Private Const STGM_READ                As Long = &H0
Private Const STGM_SHARE_DENY_WRITE    As Long = &H20
Private Const STGM_SHARE_EXCLUSIVE     As Long = &H10
Private Const STATFLAG_NONAME          As Long = 1
Private Const CC_STDCALL               As Long = 4
Private Const INVALID_FILE_ATTRIBUTES  As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const GMEM_MOVEABLE            As Long = &H2

Public Enum ModuleType
  modOther
  modStandard
  modClass
  modDesigner
  modDocument
End Enum

Private Type ModuleRecord
  Name       As String
  StreamName As String
  Offset     As Long
  ModType    As ModuleType
  IsForm     As Boolean
  SourceCode As String
End Type

Private Type ProjectRecord
  Name                As String
  FilePath            As String
  CodePage            As Long
  IsPasswordProtected As Boolean
  ModuleCount         As Long
  Modules()           As ModuleRecord
End Type

Private m_Project As ProjectRecord

' -------------------
'  Public Properties
' -------------------

Public Property Get ProjectName() As String
  ProjectName = m_Project.Name
End Property

Public Property Get CodePage() As Long
  CodePage = m_Project.CodePage
End Property

Public Property Get IsPasswordProtected() As Boolean
  IsPasswordProtected = m_Project.IsPasswordProtected
End Property

Public Property Get ModuleCount() As Long
  ModuleCount = m_Project.ModuleCount
End Property

Public Property Get ModuleName(ByVal Index As Long) As String
  If Index < 1 Or Index > m_Project.ModuleCount Then Debug.Print "ModuleName - Index out of range": Exit Property
  ModuleName = m_Project.Modules(Index).Name
End Property

Public Property Get ModuleType(ByVal Index As Long) As ModuleType
  If Index < 1 Or Index > m_Project.ModuleCount Then Debug.Print "ModuleType - Index out of range": Exit Property
  ModuleType = m_Project.Modules(Index).ModType
End Property

Public Property Get IsForm(ByVal Index As Long) As Boolean
  If Index < 1 Or Index > m_Project.ModuleCount Then Debug.Print "IsForm - Index out of range": Exit Property
  IsForm = m_Project.Modules(Index).IsForm
End Property

Public Property Get SourceCode(ByVal Index As Long) As String
  If Index < 1 Or Index > m_Project.ModuleCount Then Debug.Print "SourceCode - Index out of range": Exit Property
  SourceCode = m_Project.Modules(Index).SourceCode
End Property

Public Function FullSourceCode(Optional ByVal Separator As String = vbCrLf & vbCrLf & _
                               "' ------------------------------" & vbCrLf & _
                               "' Module: {0}" & vbCrLf & _
                               "' ------------------------------" & vbCrLf) As String
    
  If m_Project.ModuleCount = 0 Then Exit Function
  Dim parts() As String, i As Long
  ReDim parts(1 To m_Project.ModuleCount)
  For i = 1 To m_Project.ModuleCount
    parts(i) = Replace(Separator, "{0}", m_Project.Modules(i).Name) & m_Project.Modules(i).SourceCode
  Next i

  FullSourceCode = Join(parts, vbNullString)

End Function

' --------------------------
'  Public Methods - Loading
' --------------------------

Public Sub LoadFromFile(ByVal FilePath As String)
  If Not FileExists(FilePath) Then
    Debug.Print "LoadFromFile: File not found: " & FilePath
    Exit Sub
  End If

  Select Case LCase$(Right$(FilePath, 4))
    Case "xlsm", "xlam", "docm", "dotm", "pptm", "potm"
      Dim TempFolder As String
      Dim BinPath    As String
      TempFolder = Environ$("TEMP") & "\vbaXray_" & Format(Now, "hhnnss")
      If Not CreateFolderRecursive(TempFolder) Then
        Debug.Print "LoadFromFile: Failed to create temp folder"
        Exit Sub
      End If
      BinPath = ExtractVBAProjectBin(FilePath, TempFolder)
      If Len(BinPath) > 0 Then LoadFromStorage StorageOpenFile(BinPath)
      On Error Resume Next
      Kill BinPath
      RmDir TempFolder
    Case Else
      LoadFromStorage StorageOpenFile(FilePath)
  End Select
End Sub

Public Sub LoadFromByteArray(ByRef VBAProjectBinData() As Byte)
  LoadFromStorage StorageOpenBytes(VBAProjectBinData)
End Sub

Public Sub Reset()
  Dim EmptyStruct As ProjectRecord
  m_Project = EmptyStruct
  m_Project.CodePage = 1252
  m_Project.ModuleCount = 0
  m_Project.Name = ""
  m_Project.IsPasswordProtected = False
End Sub

' ----------------------------
'  Public Methods - Exporting
' ----------------------------

Public Sub ExportAll(ByVal ExportPath As String, _
                     Optional ByVal StoreVBAprojectBinInDest As Boolean = True)
                     
  If m_Project.ModuleCount = 0 Then
    Debug.Print "ExportAll - No project loaded - call LoadFromFile or LoadFromByteArray first"
  End If

  If Not FolderExists(ExportPath) Then CreateFolderRecursive ExportPath

  Dim i As Long
  For i = 1 To m_Project.ModuleCount
    ExportModule i, ExportPath
  Next i

  Debug.Print "Export complete. " & m_Project.ModuleCount & " modules exported to: " & ExportPath
End Sub

Public Sub ExportModule(ByVal Index As Long, ByVal ExportPath As String)
  If m_Project.ModuleCount = 0 Then
    err.Raise vbObjectError, "ExportModule", "No project loaded"
  End If
  If Index < 1 Or Index > m_Project.ModuleCount Then err.Raise 9

  If Not FolderExists(ExportPath) Then CreateFolderRecursive ExportPath
  WriteModule Index, ExportPath
End Sub

Public Sub ExportModuleByName(ByVal modName As String, ByVal ExportPath As String)
  If m_Project.ModuleCount = 0 Then
    Debug.Print "ExportModuleByName - No project loaded"
    Exit Sub
  End If

  Dim i As Long
  For i = 1 To m_Project.ModuleCount
    If StrComp(m_Project.Modules(i).Name, modName, vbTextCompare) = 0 Then
      ExportModule i, ExportPath
      Exit Sub
    End If
  Next

  Debug.Print "ExportModuleByName - Module not found: " & modName
End Sub

' ----------------------
'  Core logic (private)
' ----------------------

' Based off code I wote for a separate project:
' https://www.mrexcel.com/board/threads/alternative-approach-to-fix-those-problematic-excel-files-when-excel-sees-1048576-as-the-last-row-you-cant-fix-it.1199633/post-5856635
'
' Extracts vbaProject.bin from an Office Open XML file (.xlsm, .docm, .pptm etc.).
' NB: Access files (.accdb/.accde) are not supported - my understanding is that they use a
' completely different binary format and cannot be handled this way.
'
' Returns the path to the extracted file, or empty string on failure.

Private Function ExtractVBAProjectBin(ByVal OfficePath As String, ByVal DestFolder As Variant) As String
    
    Dim ZipPath    As Variant
    Dim ShellApp   As Object
    Dim ZipDir     As Object
    Dim BinFile    As Object
    Dim DestNS     As Object
    Dim InternalPath As String
    Dim Elapsed    As Single
    
    On Error GoTo Failed
    
    ' Determine internal ZIP path by file extension
    Select Case LCase$(Right$(OfficePath, 4))
      Case "xlsm", "xlam"
        InternalPath = "\xl\"
      Case "docm", "dotm"
        InternalPath = "\word\"
      Case "pptm", "potm"
        InternalPath = "\ppt\"
      Case Else
        If LCase(OfficePath) Like "*vbaProject.bin" Then
          ExtractVBAProjectBin = OfficePath
          Exit Function
        End If
    End Select
    
    ' Make a temporary ZIP copy
    ZipPath = Left$(OfficePath, InStrRev(OfficePath, ".")) & "zip"
    FileCopy OfficePath, ZipPath
    
    Set ShellApp = CreateObject("Shell.Application")
    Set ZipDir = ShellApp.nameSpace(ZipPath & InternalPath)
    
    If ZipDir Is Nothing Then GoTo Failed
    
    Set BinFile = ZipDir.items.Item("vbaProject.bin")
    If BinFile Is Nothing Then GoTo Failed  ' No VBA project in this file
    
    If Not FolderExists(DestFolder) Then
      If Not CreateFolderRecursive(DestFolder) Then
        Debug.Print "ExtractVBAProjectBin: Failed to create destination folder"
        Exit Function
      End If
    End If
    Set DestNS = ShellApp.nameSpace(DestFolder)
    If DestNS Is Nothing Then GoTo Failed
    
    DestNS.CopyHere BinFile.Path, 20  ' 20 = suppress dialogs and prompts
    
    ' Wait for async copy to complete (timeout after 30 seconds)
    Elapsed = Timer
    Do Until Not DestNS.items.Item("vbaproject.bin") Is Nothing
        DoEvents
        If Timer - Elapsed > 30 Then GoTo Failed
    Loop
    
    Kill ZipPath
    ExtractVBAProjectBin = DestFolder & "\vbaProject.bin"
    Exit Function

Failed:
    On Error Resume Next
    Kill ZipPath
    ExtractVBAProjectBin = ""
End Function

Private Sub LoadFromStorage(ByRef topStg As stdole.IUnknown)
  
  Reset

  Dim vbaStg As stdole.IUnknown
  Set vbaStg = StorageOpenStorage(topStg, "VBA")

  Dim dirBytes()  As Byte
  Dim decompDir() As Byte
  dirBytes = StreamReadBytes(StorageOpenStream(vbaStg, "dir"))
  decompDir = DecompressStream(dirBytes)
  ParseDirStream decompDir        ' Parse module records + codepage

  Dim ProjectBytes() As Byte      ' Parse PROJECT stream (not compressed)
  ProjectBytes = StreamReadBytes(StorageOpenStream(topStg, "PROJECT"))
  ProcessProject BytesToString(ProjectBytes, 0, UBound(ProjectBytes) + 1, m_Project.CodePage)

  If m_Project.ModuleCount = 0 Then Debug.Print "No modules found in dir stream": Exit Sub

  Dim i As Long
  For i = 1 To m_Project.ModuleCount
    ReadModuleSource vbaStg, i
  Next

  Debug.Print "Project loaded: " & m_Project.Name & " (" & m_Project.ModuleCount & " modules)"
End Sub

Private Sub ReadModuleSource(ByRef vbaStg As stdole.IUnknown, ByVal Index As Long)
  With m_Project.Modules(Index)
    Dim modBytes() As Byte
    modBytes = StreamReadBytes(StorageOpenStream(vbaStg, .StreamName))

    Dim compLen As Long
    compLen = UBound(modBytes) - .Offset + 1
    If compLen <= 0 Then Exit Sub

    Dim compSource() As Byte
    Dim data()       As Byte
    compSource = SliceBytes(modBytes, .Offset, compLen)
    data = DecompressStream(compSource)
    .SourceCode = BytesToString(data, 0, UBound(data) + 1, m_Project.CodePage)
  End With
End Sub

Private Sub WriteModule(ByVal Index As Long, ByVal ExportPath As String)
    With m_Project.Modules(Index)
        Dim Extension As String
        Dim SubFolder As String
        
        Select Case .ModType
            Case modStandard
                Extension = ".bas"
                SubFolder = "01_Standard"
            Case modClass
                Extension = ".cls"
                SubFolder = "02_Class"
            Case modDocument
                Extension = ".cls"
                SubFolder = "03_Document"
            Case modDesigner
                Extension = ".frm"
                SubFolder = "04_Designer"
            Case Else
                Extension = ".txt"
                SubFolder = "05_Other"
        End Select
        
        Dim TargetPath As String
        TargetPath = ExportPath & SubFolder & "\"
        If Not FolderExists(TargetPath) Then CreateFolderRecursive TargetPath
        
        WriteFile TargetPath & .Name & Extension, BytesToBytes(.SourceCode, m_Project.CodePage)
        
        Debug.Print "Exported: " & SubFolder & "\" & .Name & Extension
    End With
End Sub

Private Sub ParseDirStream(ByRef data() As Byte)

  Dim tempList() As ModuleRecord
  ReDim tempList(0 To 63)
  Dim count As Long: count = 0

  m_Project.CodePage = 1252

  Dim Pos        As Long: Pos = 0
  Dim RecordId   As Long
  Dim RecordSize As Long

  Do While Pos < UBound(data) - 5
    RecordId = CLng(data(Pos)) Or (CLng(data(Pos + 1)) * 256&)

    ' MS-OVBA: 2.3.4.2.1.5 PROJECTCODEPAGE (0x0003)
    If RecordId = &H3 Then
      RecordSize = CLng(data(Pos + 2)) Or (CLng(data(Pos + 3)) * 256&) Or (CLng(data(Pos + 4)) * 65536) Or (CLng(data(Pos + 5)) * 16777216)
      If RecordSize = 2 Then
        m_Project.CodePage = CLng(data(Pos + 6)) Or (CLng(data(Pos + 7)) * 256&)
        Debug.Print "Project CodePage: " & m_Project.CodePage
      End If
    End If

    ' Scan for the MODULENAME header (0x0019)
    If RecordId = &H19 Then
      RecordSize = CLng(data(Pos + 2)) Or (CLng(data(Pos + 3)) * 256&) Or (CLng(data(Pos + 4)) * 65536) Or (CLng(data(Pos + 5)) * 16777216)
      If RecordSize > 0 And RecordSize < 256 Then
        Dim modRec As ModuleRecord
        modRec.Name = ""
        modRec.StreamName = ""
        modRec.Offset = 0
        modRec.ModType = modOther
        modRec.IsForm = False
        modRec.SourceCode = ""

        Dim modPos As Long
        modPos = Pos

        Do While modPos < UBound(data) - 5
          Dim mId As Long, mSize As Long
          mId = CLng(data(modPos)) Or (CLng(data(modPos + 1)) * 256&)
          mSize = CLng(data(modPos + 2)) Or (CLng(data(modPos + 3)) * 256&) Or (CLng(data(modPos + 4)) * 65536) Or (CLng(data(modPos + 5)) * 16777216)
          modPos = modPos + 6
          If modPos + mSize > UBound(data) + 1 Then Exit Do

          If mId = &H19 Then        ' MODULENAME
            modRec.Name = BytesToString(data, modPos, mSize, m_Project.CodePage)
          ElseIf mId = &H1A Then    ' MODULESTREAMNAME
            modRec.StreamName = BytesToString(data, modPos, mSize, m_Project.CodePage)
          ElseIf mId = &H21 Then    ' standard module
            modRec.ModType = modStandard
          ElseIf mId = &H22 Then    ' class / designer / document (refined by ProcessProject)
            modRec.ModType = modClass
          ElseIf mId = &H31 Then    ' MODULEOFFSET
            modRec.Offset = CLng(data(modPos)) Or (CLng(data(modPos + 1)) * 256&) Or (CLng(data(modPos + 2)) * 65536) Or (CLng(data(modPos + 3)) * 16777216)
          ElseIf mId = &H2B Then    ' MODULETERMINATOR
            modPos = modPos + mSize
            Exit Do
          End If
          modPos = modPos + mSize
        Loop

        If Len(modRec.StreamName) > 0 And modRec.Offset > 0 Then
          If count = UBound(tempList) Then
            ReDim Preserve tempList(0 To count + 63)
          End If
          count = count + 1
          tempList(count) = modRec
        End If
        Pos = modPos
        GoTo ContinueScan
      End If
    End If

    Pos = Pos + 1
ContinueScan:
  Loop

  If count > 0 Then
    ReDim m_Project.Modules(1 To count)
    Dim i As Long
    For i = 1 To count
      m_Project.Modules(i) = tempList(i)
    Next
    m_Project.ModuleCount = count
  End If
End Sub

Private Sub ProcessProject(ByVal StreamString As String)
  Dim Lines()  As String
  Dim parts()  As String
  Dim Key      As String
  Dim modName  As String
  Dim idx      As Long
  Dim i        As Long

  Lines = Split(StreamString, vbCrLf)
  Dim Line As Variant
  For Each Line In Lines
    Line = Trim$(Line)
    If Len(Line) = 0 Then GoTo NextLine
    If Left$(Line, 1) = "'" Then GoTo NextLine
    parts = Split(Line, "=")
    If UBound(parts) < 1 Then GoTo NextLine
    Key = LCase$(Trim$(parts(0)))
    modName = Trim$(parts(1))
    If Left$(modName, 1) = Chr(34) Then modName = Split(modName, Chr(34))(1)
    
    ' Strip any /& annotation (e.g. "ThisWorkbook/&H00000000")
    Dim slashPos As Long
    slashPos = InStr(modName, "/")
    If slashPos > 0 Then modName = Left$(modName, slashPos - 1)

    Select Case Key
    Case "name"
      m_Project.Name = modName
      GoTo NextLine
    Case "dpb"
    ' Per MS-OVBA 2.3.1.16: DPB value length of 16 hex chars = Length of 1 = no password
    ' A password protected project will have a DPB value of 58 hex chars (Length = 29)
      m_Project.IsPasswordProtected = (Len(modName) > 16)
      GoTo NextLine
    End Select

    ' Find module by name - linear scan. I originally used a dictionary object
    ' but it felt overdone.
    idx = 0
    For i = 1 To m_Project.ModuleCount
      If LCase$(m_Project.Modules(i).Name) = LCase$(modName) Or LCase$(m_Project.Modules(i).StreamName) = LCase$(modName) Then
        idx = i
        Exit For
      End If
    Next
    If idx = 0 Then GoTo NextLine

    With m_Project.Modules(idx)
      Select Case Key
      Case "document"
        .ModType = modDocument
        .IsForm = False
      Case "baseclass"
        .ModType = modDesigner
        .IsForm = True
      Case "class"
        If .ModType = modOther Then .ModType = modClass
      Case "module"
        If .ModType = modOther Then .ModType = modStandard
      End Select
    End With
NextLine:
  Next
End Sub

' ------------------------
'  Storage/Stream helpers
' ------------------------
Private Function StorageOpenFile(ByVal FilePath As String) As stdole.IUnknown
  Dim hr As Long
  hr = StgOpenStorage(StrPtr(FilePath), 0, STGM_READ Or STGM_SHARE_DENY_WRITE, 0, 0, StorageOpenFile)
  If hr < 0 Then Debug.Print "StgOpenStorage - Error: " & hr
End Function

Private Function StorageOpenBytes(ByRef data() As Byte) As stdole.IUnknown
  Dim hMem          As LongPtr
  Dim lpMem         As LongPtr
  Dim pLockBytesPtr As LongPtr
  Dim hr            As Long
  Dim dataSize      As LongPtr

  dataSize = UBound(data) - LBound(data) + 1

  hMem = GlobalAlloc(GMEM_MOVEABLE, dataSize)
  If hMem = 0 Then Debug.Print "StorageOpenBytes - GlobalAlloc failed"

  lpMem = GlobalLock(hMem)
  If lpMem = 0 Then
    GlobalFree hMem
    Debug.Print "StorageOpenBytes - GlobalLock failed"
    Exit Function
  End If

  CopyMemory lpMem, VarPtr(data(LBound(data))), dataSize
  GlobalUnlock hMem

  hr = CreateILockBytesOnHGlobal(hMem, 1, pLockBytesPtr)
  If hr < 0 Then
    GlobalFree hMem
    Debug.Print "StorageOpenBytes - CreateILockBytesOnHGlobal failed: &H" & hex(hr)
    Exit Function
  End If

  hr = StgOpenStorageOnILockBytes(pLockBytesPtr, 0, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0, 0, StorageOpenBytes)
  If hr < 0 Then Debug.Print "StorageOpenBytes - StgOpenStorageOnILockBytes failed: &H" & hex(hr)
End Function

Private Function StorageOpenStorage(pStg As stdole.IUnknown, ByVal storageName As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_OpenStorage, StrPtr(storageName), 0&, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0&, 0&, VarPtr(StorageOpenStorage))
  If hr < 0 Then Debug.Print "IStorage.OpenStorage - hResult: &H" & hex(hr)
End Function

Private Function StorageOpenStream(pStg As stdole.IUnknown, ByVal StreamName As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_OpenStream, StrPtr(StreamName), 0&, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0&, VarPtr(StorageOpenStream))
  If hr < 0 Then Debug.Print "IStorage.OpenStream - hResult: &H" & hex(hr)
End Function

Private Function StorageEnumElements(ByRef pStg As stdole.IUnknown) As Collection
  Dim pEnum As stdole.IUnknown
  Dim hr    As Long
  Dim items As New Collection
  Dim stat  As STATSTG
  Dim entry As Object

  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_EnumElements, 0&, 0&, 0&, VarPtr(pEnum))
  If hr < 0 Then Debug.Print "IStorage.EnumElements - hResult: &H" & hex(hr)

  Do
    hr = DispCallByVtbl(pEnum, IEnumSTATSTGInterface.IEnum_Next, 1&, VarPtr(stat), 0&)
    If hr <> 0 Then Exit Do
    If stat.pwcsName <> 0 Then
      Set entry = CreateObject("Scripting.Dictionary")
      entry("name") = PtrToString(stat.pwcsName)
      entry("type") = stat.dwType
      entry("size") = CLng(stat.cbSize * 10000)
      CoTaskMemFree stat.pwcsName
      items.Add entry
    End If
  Loop

  Set StorageEnumElements = items
End Function

Private Function StorageExists(ByRef pStg As stdole.IUnknown, ByVal ElementName As String) As Boolean
  Dim items As Collection
  Dim entry As Variant
  Set items = StorageEnumElements(pStg)
  For Each entry In items
    If LCase$(entry("name")) = LCase$(ElementName) Then
      StorageExists = True
      Exit Function
    End If
  Next
End Function

Private Function StreamGetSize(ByRef pStm As stdole.IUnknown) As Long
  Dim stat As STATSTG
  Dim hr   As Long
  hr = DispCallByVtbl(pStm, IStreamInterface.IStm_stat, VarPtr(stat), STATFLAG_NONAME)
  If hr < 0 Then Debug.Print "IStream.Stat - hResult: &H" & hex(hr)
  StreamGetSize = CLng(stat.cbSize * 10000)
End Function

Private Function StreamReadBytes(pStm As stdole.IUnknown) As Byte()
  Dim Buffer() As Byte
  Dim chunk(0 To 32767) As Byte
  Dim bytesRead As Long
  Dim hr As Long
  Dim total As Long
  Dim oldUB As Long

  If pStm Is Nothing Then Exit Function

  hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Seek, 0@, 0@, 0&, 0&)

  Do
    hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Read, VarPtr(chunk(0)), 32768, VarPtr(bytesRead))
    If hr < 0 Then Debug.Print "IStream.Read - hResult: &H" & hex(hr)
    If bytesRead <= 0 Then Exit Do

    If total = 0 Then
      ReDim Buffer(0 To bytesRead - 1)
    Else
      oldUB = UBound(Buffer)
      ReDim Preserve Buffer(0 To oldUB + bytesRead)
    End If

    CopyMemory VarPtr(Buffer(total)), VarPtr(chunk(0)), bytesRead
    total = total + bytesRead
  Loop

  If total > 0 Then
    ReDim Preserve Buffer(0 To total - 1)
  Else
    ReDim Buffer(0 To 0)
  End If

  StreamReadBytes = Buffer
End Function

' LZ77 decompression
' MS-OVBA: 2.4.1 Compression and Decompression

Private Function DecompressStream(ByRef CompressedData() As Byte) As Byte()
  Dim output() As Byte, outPos As Long, inPos As Long
  ReDim output(0 To 65535)
  If UBound(CompressedData) < 2 Then DecompressStream = output: Exit Function
  If CompressedData(0) = 1 Then inPos = 1
  Do While inPos < UBound(CompressedData)
    If inPos + 1 > UBound(CompressedData) Then Exit Do

    Dim chunkHeader  As Long
    Dim isCompressed As Boolean
    Dim payloadSize  As Long
    Dim chunkEnd     As Long
    Dim chunkOutPos  As Long

    chunkHeader = CLng(CompressedData(inPos)) Or (CLng(CompressedData(inPos + 1)) * 256&)
    inPos = inPos + 2
    isCompressed = ((chunkHeader And &H8000&) = &H8000&)
    payloadSize = (chunkHeader And &HFFF&) + 1
    chunkEnd = inPos + payloadSize
    chunkOutPos = 0

    If Not isCompressed Then
      If outPos + 4096 > UBound(output) Then ReDim Preserve output(0 To outPos + 4096 + 65535)
      Dim i As Long
      For i = 0 To 4095
        If inPos <= UBound(CompressedData) Then
          output(outPos) = CompressedData(inPos)
          inPos = inPos + 1: outPos = outPos + 1: chunkOutPos = chunkOutPos + 1
        End If
      Next i
    Else
      Do While inPos < chunkEnd And inPos <= UBound(CompressedData)
        Dim flagByte As Byte
        flagByte = CompressedData(inPos)
        inPos = inPos + 1

        Dim bitIndex As Integer
        For bitIndex = 0 To 7
          If chunkOutPos >= 4096 Then Exit Do
          If inPos >= chunkEnd Or inPos > UBound(CompressedData) Then Exit Do

          If (flagByte And (2 ^ bitIndex)) = 0 Then
            If outPos > UBound(output) Then ReDim Preserve output(0 To outPos + 65535)
            output(outPos) = CompressedData(inPos)
            outPos = outPos + 1: chunkOutPos = chunkOutPos + 1: inPos = inPos + 1
          Else
            If inPos + 1 > UBound(CompressedData) Then Exit Do
            Dim token As Long
            token = CLng(CompressedData(inPos)) Or (CLng(CompressedData(inPos + 1)) * 256&)
            inPos = inPos + 2

            Dim bitCount As Integer
            Select Case chunkOutPos
            Case Is <= 16:   bitCount = 4
            Case Is <= 32:   bitCount = 5
            Case Is <= 64:   bitCount = 6
            Case Is <= 128:  bitCount = 7
            Case Is <= 256:  bitCount = 8
            Case Is <= 512:  bitCount = 9
            Case Is <= 1024: bitCount = 10
            Case Is <= 2048: bitCount = 11
            Case Else:       bitCount = 12
            End Select

            Dim lengthMask  As Long: lengthMask = (2 ^ (16 - bitCount)) - 1
            Dim copyLength  As Long: copyLength = (token And lengthMask) + 3
            Dim copyOffset  As Long: copyOffset = (token \ (2 ^ (16 - bitCount))) + 1

            If outPos + copyLength > UBound(output) Then ReDim Preserve output(0 To outPos + copyLength + 65535)

            Dim j As Long
            For j = 0 To copyLength - 1
              output(outPos) = output(outPos - copyOffset)
              outPos = outPos + 1: chunkOutPos = chunkOutPos + 1
            Next j
          End If
        Next bitIndex
      Loop
    End If

    inPos = chunkEnd
  Loop

  If outPos > 0 Then
    ReDim Preserve output(0 To outPos - 1)
  Else
    ReDim output(0 To 0)
  End If

  DecompressStream = output
End Function

Private Function DispCallByVtbl(ByRef pUnk As stdole.IUnknown, ByVal lIndex As Long, ParamArray a() As Variant) As Variant
  Dim lIdx           As Long
  Dim vParam()       As Variant
  Dim vType(0 To 63) As Integer
  Dim vPtr(0 To 63)  As LongPtr
  Dim hResult        As Long

  vParam = a
  For lIdx = 0 To UBound(vParam)
    vType(lIdx) = VarType(vParam(lIdx))
    vPtr(lIdx) = VarPtr(vParam(lIdx))
  Next

  hResult = DispCallFunc(ObjPtr(pUnk), lIndex * PTR_SIZE, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
  If hResult < 0 Then Debug.Print "DispCallByVtbl - hResult: &H" & hex(hResult)
End Function

Private Function PtrToString(ByVal stringPtr As LongPtr) As String
  Dim buf        As String
  Dim CharLength As Long
  CharLength = lstrlenW(stringPtr)
  If CharLength > 0 Then
    buf = VBA.String$(CharLength, vbNullChar)
    CopyMemory ByVal StrPtr(buf), ByVal stringPtr, CharLength * 2
  End If
  PtrToString = buf
End Function

Private Function BytesToString(ByRef data() As Byte, ByVal Pos As Long, ByVal Size As Long, ByVal CodePage As Long) As String
  If Size <= 0 Then Exit Function
  Dim charCount As Long, Result As String
  charCount = MultiByteToWideChar(CodePage, 0, VarPtr(data(Pos)), Size, 0, 0)
  If charCount = 0 Then Exit Function
  Result = VBA.String$(charCount, 0)
  MultiByteToWideChar CodePage, 0, VarPtr(data(Pos)), Size, StrPtr(Result), charCount
  BytesToString = Result
End Function

Private Function BytesToBytes(ByVal s As String, ByVal CodePage As Long) As Byte()
  ' Converts a VBA string back to bytes using the original code page
  ' Used to preserve encoding fidelity on export
  Dim Result() As Byte
  Dim Size As Long
  Size = WideCharToMultiByte(CodePage, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
  If Size = 0 Then Exit Function
  ReDim Result(0 To Size - 1)
  WideCharToMultiByte CodePage, 0, StrPtr(s), Len(s), VarPtr(Result(0)), Size, 0, 0
  BytesToBytes = Result
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
  Dim attr As Long
  attr = apiGetFileAttributesW(StrPtr(FilePath))
  FileExists = (attr <> INVALID_FILE_ATTRIBUTES) And ((attr And FILE_ATTRIBUTE_DIRECTORY) = 0)
End Function

Private Function FolderExists(ByVal FolderPath As String) As Boolean
  Dim attr As Long
  attr = apiGetFileAttributesW(StrPtr(FolderPath))
  FolderExists = (attr <> INVALID_FILE_ATTRIBUTES) And ((attr And FILE_ATTRIBUTE_DIRECTORY) <> 0)
End Function

Private Function CreateFolderRecursive(ByVal FolderPath As String) As Boolean
  FolderPath = VBA.Trim(FolderPath)
  If Len(FolderPath) Then
    If Not FolderExists(FolderPath) Then
      Dim parent As String
      parent = Left$(FolderPath, InStrRev(FolderPath, "\") - 1)
      If Len(parent) > 0 And Not FolderExists(parent) Then If Not CreateFolderRecursive(parent) Then Exit Function
      CreateFolderRecursive = (apiCreateDirectoryW(StrPtr(FolderPath & vbNullChar), 0) <> 0)
    End If
    CreateFolderRecursive = True
  End If
End Function

Private Function SliceBytes(ByRef data() As Byte, ByVal Pos As Long, ByVal Size As Long) As Byte()
  Dim Result() As Byte
  If Size <= 0 Then Exit Function
  ReDim Result(0 To Size - 1)
  CopyMemory VarPtr(Result(0)), VarPtr(data(Pos)), Size
  SliceBytes = Result
End Function

Private Function WriteFile(ByVal fileName As String, ByRef data As Variant, Optional ByVal WriteBinary As Boolean = True) As Boolean
  On Error GoTo errHandler
  Const adTypeBinary = 1
  Const adTypeText = 2
  Const adSaveCreateOverWrite = 2
  Dim stream As Object
  Set stream = CreateObject("ADODB.Stream")
  With stream
    .Type = IIf(WriteBinary, adTypeBinary, adTypeText)
    .Open
    If WriteBinary Then
      .Write data
    Else
      .Charset = "UTF-8"
      .WriteText CStr(data)
    End If
    .SaveToFile fileName, adSaveCreateOverWrite
    .Close
  End With
  WriteFile = True
  Exit Function
errHandler:
  WriteFile = False
End Function

Private Sub Class_Initialize()
  Reset
End Sub
