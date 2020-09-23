Attribute VB_Name = "modShellFileOperation"
'Shell File Operation , By Max Raskin

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'A Type for SHFileOperation API
Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type
'Enumarated Operations and Flags for SHFileOperation API
Public Enum SHOps
    Copy = &H2
    Delete = &H3
    SMove = &H1
    Rename = &H4
End Enum
'Flags for Shell File Operations
Public Enum SHFlags
    AllowUndo = &H40
    CopyOnlyFiles = &H80
    MultiDestinationFiles = &H1
    DontPromptOnOverwrite = &H10
    DontPromptOnCreateFolders = &H200
    RenameIfExists = &H8
    NoProgressDialog = &H4
    ProgressDialogWithNoFileNames = &H100
End Enum

'Preforme file operations just like in windows explorer -
'with progress dialog
Public Function ShellFileOp(hwnd As Long, Operation As SHOps, Source As String, Destination As String, Optional Flags As SHFlags, Optional ProgressDialogTitle As String)
    Dim shf As SHFILEOPSTRUCT
    If Flags <> 0 Then shf.fFlags = Flags
    If ProgressDialogTitle <> "" Then shf.lpszProgressTitle = ProgressDialogTitle
    shf.hwnd = hwnd
    shf.pFrom = Source
    shf.pTo = Destination
    shf.wFunc = Operation
    sh = SHFileOperation(shf)
End Function


