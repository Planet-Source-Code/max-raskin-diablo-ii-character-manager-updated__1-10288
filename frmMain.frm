VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Diablo II Character Manager"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8760
   Begin VB.Frame fraBkChar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Backed Up Characters:"
      ForeColor       =   &H0000FFFF&
      Height          =   5775
      Left            =   5460
      TabIndex        =   17
      Top             =   30
      Width           =   3255
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H8000000C&
         Caption         =   "&Delete"
         Height          =   345
         Left            =   2190
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Click here to remove a selected backup (Alt+D)"
         Top             =   5340
         Width           =   945
      End
      Begin VB.CommandButton cmdRestore 
         BackColor       =   &H8000000C&
         Caption         =   "R&estore"
         Height          =   345
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Click here to load a backupped character into Diablo II's save directory (Alt+E)"
         Top             =   5340
         Width           =   945
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H80000007&
         Caption         =   "Description:"
         ForeColor       =   &H0000FFFF&
         Height          =   1755
         Left            =   60
         TabIndex        =   21
         Top             =   3090
         Width           =   3135
         Begin VB.TextBox txtDesc 
            BackColor       =   &H80000007&
            ForeColor       =   &H80000009&
            Height          =   1455
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Text            =   "frmMain.frx":030A
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000009&
         Height          =   315
         Left            =   930
         TabIndex        =   20
         Top             =   4920
         Width           =   2265
      End
      Begin VB.CommandButton cmdBk 
         BackColor       =   &H8000000C&
         Caption         =   "&Backup"
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Click here to backup the character, you will also be prompted to enter a descripition for it (optional) (Alt+B)"
         Top             =   5340
         Width           =   945
      End
      Begin VB.ListBox lstBkChars 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000005&
         Height          =   2790
         Left            =   60
         TabIndex        =   18
         Top             =   270
         Width           =   3135
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   4980
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4980
      TabIndex        =   15
      ToolTipText     =   "Exit (ESC)"
      Top             =   180
      Width           =   375
   End
   Begin VB.Frame fraBkDir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Set Backup Directory:"
      ForeColor       =   &H80000005&
      Height          =   5745
      Left            =   5460
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H8000000C&
         Caption         =   "C&ancel"
         Height          =   345
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5220
         Width           =   1395
      End
      Begin VB.CommandButton cmdDone 
         BackColor       =   &H8000000C&
         Caption         =   "&Done"
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5220
         Width           =   1395
      End
      Begin VB.CommandButton cmdNewFolder 
         Height          =   315
         Left            =   2550
         MaskColor       =   &H0080FFFF&
         Picture         =   "frmMain.frx":0366
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Create New Folder"
         Top             =   210
         Width           =   345
      End
      Begin VB.DriveListBox drvMain 
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Top             =   210
         Width           =   2475
      End
      Begin VB.DirListBox dirMain 
         Height          =   4590
         Left            =   30
         TabIndex        =   8
         Top             =   570
         Width           =   3195
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2910
         MaskColor       =   &H0080FFFF&
         TabIndex        =   7
         ToolTipText     =   "Delete Selected Folder"
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H80000008&
      Caption         =   "v1.2"
      ForeColor       =   &H0000FFFF&
      Height          =   5805
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5385
      Begin VB.CommandButton cmdRunDiablo 
         BackColor       =   &H8000000C&
         Height          =   555
         Left            =   60
         Picture         =   "frmMain.frx":0C30
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Run Diablo II"
         Top             =   1530
         Width           =   615
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H8000000C&
         Caption         =   "&R"
         BeginProperty Font 
            Name            =   "MS SystemEx"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Refresh, click here to refresh the character listincase some of the save game names are changed (Alt+R)"
         Top             =   5010
         Width           =   405
      End
      Begin VB.ListBox lstChars 
         Height          =   450
         Left            =   90
         TabIndex        =   28
         Top             =   390
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Timer tmrUpdate 
         Interval        =   10
         Left            =   120
         Top             =   780
      End
      Begin VB.Frame fraCharList 
         BackColor       =   &H80000008&
         Caption         =   "Characters List:"
         ForeColor       =   &H0000FFFF&
         Height          =   2835
         Left            =   60
         TabIndex        =   26
         Top             =   2100
         Width           =   5265
         Begin VB.ListBox lstChars2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000007&
            ForeColor       =   &H80000005&
            Height          =   2565
            Left            =   30
            TabIndex        =   27
            Top             =   180
            Width           =   5205
         End
      End
      Begin VB.CommandButton cmdChangePath 
         BackColor       =   &H8000000C&
         Caption         =   "&Change Path"
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   $"frmMain.frx":0F3A
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H80000007&
         Enabled         =   0   'False
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5400
         Width           =   3735
      End
      Begin VB.CommandButton cmdSetBkDir 
         BackColor       =   &H8000000C&
         Caption         =   "&Set Backup Directory"
         Height          =   345
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Click here to set a backup directory, where backupped characters will be stored (Alt+S)"
         Top             =   5010
         Width           =   2295
      End
      Begin VB.CommandButton cmdBkChar 
         BackColor       =   &H8000000C&
         Caption         =   "&Backup/Restore Character"
         Height          =   345
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click here to backup or restore a character, select one from the list first (Alt+B)"
         Top             =   5010
         Width           =   2295
      End
      Begin VB.ListBox lstSaves 
         Height          =   645
         Left            =   210
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   0  'None
         Height          =   1740
         Left            =   870
         Picture         =   "frmMain.frx":0FCC
         ScaleHeight     =   1740
         ScaleWidth      =   3900
         TabIndex        =   1
         Top             =   180
         Width           =   3900
         Begin VB.Label lblCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Character Manager"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   405
            Left            =   420
            TabIndex        =   3
            Top             =   1380
            Width           =   2790
         End
      End
      Begin VB.Label lblC 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "By Max Raskin"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2130
         TabIndex        =   4
         Top             =   1890
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sINIFile As String, msgResult As VbMsgBoxResult, sRetVal As String, sBackupDir As String, m_cUnzip As New cUnzip
Dim sDescFile As String, sBkExt As String
Private WithEvents m_cZ As cZip
Attribute m_cZ.VB_VarHelpID = -1


'Diablo 2 Character Manager v1.2

'Purpose: Quickly backup / restore characters in several mouse clicks !

'Created by: Max Raskin

'NOTE: DON'T EVER PLACE THE EXE INTO A DIR CONTAINES D2S FILES - IT WILL DELETE THEM ALL!

'Code best viewed on 1024x768/1125x864 or higher resolution

'Credits for the Zip/Unzip classes and modules goes to http://vbaccelerator.com/
'and the Info-Zip group for making the this great file compression


'Updates in version 1.2:
'=======================
'Fixed a bug when running in the first time it wont show the description of a selected character
'Fixed issues with the getting characters info, now it gets also women's title fine =)

'Updates in version 1.1:
'=======================

'Minor bug fixes:
'~~~~~~~~~~~~~~~~
'* When clicking Backup, if a file u're trying to backup doesn't exists it wont display error messages such as 'select a character first', 'type a filename first' etc..
'* REALLY minor: Refresh of the file name text box , so it wont be color inverted =)

'New stuff added:
'~~~~~~~~~~~~~~~~
'* Now displays character status (Sir, Lord, Baron)
'* And also displays character level
'* A refresh button added incase you or some other program changed the name of the save file, it will simply refresh the character list
'* A Run Diablo II button added, it will look first of all for DLoad.exe (cracked) and if not found for Diablo II.exe (non cracked or cracked without a loader, which means without b.net, which means, never mind that b.net is a big lag!) =)

Private Sub cmdBk_Click()
On Error Resume Next
    Dim sDesc As String
    If lstChars.ListIndex = -1 Then
        MsgBox "Select A Character First!", vbExclamation, ""
        lstChars2.SetFocus
        Exit Sub
    End If
    If Trim(txtFileName) = "" Then
        MsgBox "Enter A Backup File Name First!", vbExclamation, ""
        txtFileName.SetFocus
        txtFileName.Refresh 'Refresh so the back color wont be inverted =)
        Exit Sub
    End If
    If Dir(sBackupDir & txtFileName.Text & sBkExt) <> "" Then
        msgResult = MsgBox("The file: " & txtFileName.Text & sBkExt & " already exists, do you want to contiune and overwrite it ?", vbQuestion + vbYesNo, "Already Exists")
            If msgResult = vbYes Then
            Kill sBackupDir & Filename & sBkExt
            'Save description text to file (if there is any):
            sDesc = InputBox("Enter A Short Description (Optional):", "Description", "")
            If Trim(sDesc) <> "" Then SaveFile sDescFile, sDesc
            With m_cZ 'Zip e'm ! :-)
                .ZipFile = sBackupDir & txtFileName.Text & sBkExt
                .StoreFolderNames = False
                .RecurseSubDirs = False
                .ClearFileSpecs
                .AddFileSpec txtPath.Text & "Save\" & lstChars.List(lstChars.ListIndex) & ".d2s"
                .AddFileSpec sDescFile
                .Zip
            End With
            Kill sDescFile
            ShowBkChar 'Refresh
            txtDesc.Text = "Select a character from the Backupped Characters list to get a character description."
        Else
            Exit Sub
        End If
    Else
         sDesc = InputBox("Enter A Short Description (Optional):", "Description", "")
            If Trim(sDesc) <> "" Then SaveFile sDescFile, sDesc
            With m_cZ 'Zip e'm ! :-)
                .ZipFile = sBackupDir & txtFileName.Text & sBkExt
                .StoreFolderNames = False
                .RecurseSubDirs = False
                .ClearFileSpecs
                .AddFileSpec txtPath.Text & "Save\" & lstChars.List(lstChars.ListIndex) & ".d2s"
                .AddFileSpec sDescFile
                .Zip
            End With
            Kill sDescFile
            ShowBkChar 'Refresh
            txtDesc.Text = "Select a character from the Backupped Characters list to get a character description."
    End If
End Sub

Private Sub cmdBkChar_Click() 'Show Backup Character Frame
    If fraBkDir.Visible = True Then fraBkDir.Visible = False
    ShowBkChar
End Sub

Private Sub cmdDel_Click()
    ShellFileOp Me.hwnd, Delete, sBackupDir & lstBkChars.List(lstBkChars.ListIndex) & sBkExt, ""
    ShowBkChar 'Refresh
End Sub

Private Sub cmdDone_Click() 'Set backup directory
    sBackupDir = dirMain.Path
    If sBackupDir = "" Then Exit Sub
    If Right(sBackupDir, 1) <> "\" Then sBackupDir = sBackupDir & "\"
    ShowBkChar
End Sub

Private Sub cmdCancel_Click()
    ShowBkChar
End Sub

Private Sub cmdChangePath_Click()
    Dim sPath As String
    sPath = BrowseForFolder(Me.hwnd, "Browse For Diablo II's Folder:", ReturnFileSystemFoldersOnly)
    If sPath <> "" Then txtPath.Text = sPath
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    EnumChars 'Refresh character list
End Sub

Private Sub cmdRestore_Click()
    On Error Resume Next
    Kill sBackupDir & "*.d2s"
    If lstBkChars.ListIndex <> -1 Then
        With m_cUnzip 'UnZip the character
            .UnzipFolder = sBackupDir
            .ZipFile = sBackupDir & lstBkChars.List(lstBkChars.ListIndex) & sBkExt
            .OverwriteExisting = True
            .Unzip
        End With
    Else
        MsgBox "Select A Character First!", vbExclamation, ""
        Exit Sub
    End If
    'Move it to Diablo II's save directory
    ShellFileOp Me.hwnd, SMove, sBackupDir & "*.d2s", txtPath.Text & "Save", NoProgressDialog
    Kill sDescFile
    EnumChars
End Sub

Private Sub cmdRunDiablo_Click()
    ChDir txtPath.Text
    If Dir(txtPath.Text & "DLoad.exe") <> "" Then
        Shell txtPath.Text & "DLoad.exe", vbNormalFocus 'Incase you are to lazy to insert the cd everytime, and you have the crack (like i do =) this is a usefull line
    Else
        Shell txtPath.Text & "Diablo II.exe", vbNormalFocus 'Anyway, if not ...
    End If
End Sub

Private Sub cmdSetBkDir_Click()
    If sBackupDir <> "" Then
        dirMain.Path = sBackupDir
    Else
        dirMain.Path = txtPath.Text
    End If
    fraBkChar.Visible = False
    fraBkDir.Visible = True
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Set m_cZ = New cZip
    Set m_cUnzip = New cUnzip
    sINIFile = App.Path & "Settings.ini" 'Set default INI file name
    If Right(App.Path, 1) <> "\" Then sINIFile = App.Path & "\" & "Settings.ini"
    LoadSettings
    sRetVal = GetINI(sINIFile, "Settings", "DiabloPath", "") 'Attempt to get path from INI
    'BrowseForFolder if no path saved in the INI
Browse:     If sRetVal = "" Then sRetVal = BrowseForFolder(Me.hwnd, "Browse For Diablo II's Folder:", ReturnFileSystemFoldersOnly)
    If sRetVal = "" Then
        'Make sure user selects the path
        msgResult = MsgBox("You Must Select A Path, Browse Again?", vbYesNo Or vbQuestion, "No Path Selected")
        If msgResult = vbYes Then
            GoTo Browse
        Else
            End
        End If
    Else
        txtPath.Text = sRetVal
        If Right(txtPath.Text, 1) <> "\" Then txtPath.Text = txtPath.Text & "\"
    End If
    EnumChars 'Get all characters from Diablo2Dir\Save directory
    ShowBkChar 'Enumerate backupped zip files to lstBkChars listbox
    SetVars 'Set some common variables
End Sub

Private Sub EnumChars()
    On Error Resume Next
    Dim i As Integer, l As Integer, Stats As String, CharClass As String
    tmrUpdate.Enabled = False 'to make sure, the timer wont work when its not suppost to
    lstChars2.Clear
    lstChars.Clear
    lstSaves.Clear 'Clear up the invisible text box
    EnumFilesByExt txtPath.Text & "Save", lstSaves, "d2s" 'Enum files by extension, diablo2's saves extension is "d2s"
    For i = 0 To lstSaves.ListCount - 1 'Remove the '.d2s' strings from the list items
        l = Len(lstSaves.List(i))
        lstChars.AddItem Left(lstSaves.List(i), l - 4)
        Stats = GetStatus(txtPath.Text & "Save\" & lstSaves.List(i))
        CharClass = GetClass(txtPath.Text & "Save\" & lstSaves.List(i))
        If Stats = "" Then
            lstChars2.AddItem Left(lstSaves.List(i), l - 4) & " (" & CharClass & " Level " & GetLevel(txtPath.Text & "Save\" & lstSaves.List(i)) & ")"
        Else
            lstChars2.AddItem Stats & " " & Left(lstSaves.List(i), l - 4) & " (" & CharClass & " Level " & GetLevel(txtPath.Text & "Save\" & lstSaves.List(i)) & ")"
        End If
    Next
    tmrUpdate.Enabled = True
End Sub

Private Sub SaveSettings()
    writeINI sINIFile, "Settings", "Top", Me.Top 'Remember form's Y
    writeINI sINIFile, "Settings", "Left", Me.Left 'Remember form's X
    writeINI sINIFile, "Settings", "DiabloPath", txtPath.Text 'Remember Diablo II's Path dir
    writeINI sINIFile, "Settings", "BackupPath", sBackupDir 'Remember Backup Directory
End Sub

Private Sub LoadSettings()
    sBackupDir = GetINI(sINIFile, "Settings", "BackupPath", "") 'Get Backup Directory
    Me.Top = GetINI(sINIFile, "Settings", "Top", Me.Top) 'Get Form's Y
    Me.Left = GetINI(sINIFile, "Settings", "Left", Me.Left) 'Get Form's X
End Sub

'Moving the form without a title bar:
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub fraCharList_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub fraMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub fraBkDir_DragDrop(Source As Control, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub lblC_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub lblCM_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

'On clicking on one of the backupped characters the description will be extracted from the
'zip file and loaded to the description text box
Private Sub lstBkChars_Click()
    On Error Resume Next
    SetVars
    If lstBkChars.ListIndex <> -1 Then
        With m_cUnzip 'Unzip selection, temp.
            .UnzipFolder = sBackupDir
            .ZipFile = sBackupDir & lstBkChars.List(lstBkChars.ListIndex) & sBkExt
            .OverwriteExisting = True
            .Unzip
        End With
    End If
    If Dir(sDescFile) <> "" Then
        OpenFile sDescFile, txtDesc 'Load it to the description text box
        Kill sDescFile 'Delete the description file
    End If
End Sub

Private Sub lstChars2_Click()
    On Error Resume Next
    lstChars.ListIndex = lstChars2.ListIndex
    If lstChars2.ListIndex <> -1 Then txtFileName.Text = lstChars.List(lstChars.ListIndex)
End Sub

Private Sub picLogo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub fraBkChar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub fradesc_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub lblFileName_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then MoveForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings 'Remember settings on exit
    End
End Sub

Private Sub ShowBkChar()
    Dim l As Integer
    If sBackupDir = "" Then
        dirMain.Path = txtPath.Text
        fraBkChar.Visible = False
        fraBkDir.Visible = True
        Exit Sub
    End If
    tmrUpdate.Enabled = False
    lstBkChars.Clear
    lstSaves.Clear
    EnumFilesByExt sBackupDir, lstSaves, "zip"
    For i = 0 To lstSaves.ListCount - 1 'Remove the '.zip' strings from the list items
        l = Len(lstSaves.List(i))
        lstBkChars.AddItem Left(lstSaves.List(i), l - 4)
    Next
    fraBkChar.Visible = True
    tmrUpdate.Enabled = True
End Sub

Private Sub cmdNewFolder_Click()
On Error GoTo ErrorHandler
    Dim sDirName As String, sDestPath As String
    sDestPath = dirMain.Path
    If Right(dirMain.Path, 1) <> "\" Then sDestPath = dirMain.Path & "\"
sDirName = InputBox("Enter the name of the Folder:", "Create New Folder", "Save.bak")
    If Trim(sDirName) <> "" Then
        MkDir sDestPath & sDirName
        dirMain.Path = sDestPath & sDirName
        dirMain.Refresh
    Else
        Exit Sub
    End If
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox sDestPath
        MsgBox "Unable to create the folder " & sDirName, vbExclamation, "Error"
        Exit Sub
    End If
End Sub

Private Sub drvMain_Change()
On Error GoTo ErrorHandler
    dirMain.Path = drvMain.Drive
ErrorHandler: If Err.Number <> 0 Then MsgBox "Drive Access Error", vbExclamation, "Error"
End Sub

Private Sub cmdDelete_Click()
    ShellFileOp Me.hwnd, Delete, dirMain.List(dirMain.ListIndex), ""
End Sub

'This timer will instantly update the characters in the list if the number of characters is changed
Private Sub tmrUpdate_Timer()
    lstSaves.Clear
    EnumFilesByExt txtPath.Text & "Save", lstSaves, "d2s" 'Enum files by extension, diablo2's saves extension is "d2s"
    If lstSaves.ListCount <> lstChars.ListCount Then
        EnumChars
    End If
End Sub

'Common Variables
Private Sub SetVars()
    sBkExt = ".zip" 'Backup default extension
    sDescFile = sBackupDir & "CharDesc.txt" 'Default name for the description file
End Sub

Private Sub txtFileName_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Refresh so the back color wont be inverted =)
    txtFileName.Refresh
End Sub
